from collections.abc import Iterable, Iterator
from contextlib import contextmanager
from hashlib import sha256
from io import BytesIO
import os
from pathlib import Path
import tempfile
from xml.etree.ElementTree import ParseError
from zipfile import BadZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException

from contracts import AMOUNT_LIMIT, COUNT_LIMIT, EXEMPTION_LABELS, Contract, Ledger, business_display


CONTRACT_SHEET = "계약 대장"
SUMMARY_SHEET = "업체별 현황"
SETTINGS_SHEET = "_설정"
AGGREGATION_RULES_VERSION = 1
YEAR_SHEET_HEADERS = ("회계연도", "연도별 계약 시트")
LEGACY_CONTRACT_HEADERS = (
    "순번", "계약 일자", "업체명", "사업자번호", "계약명", "계약 금액",
    "업체별 연간 누적 횟수", "업체별 연간 누적 금액",
)
CONTRACT_HEADERS = (
    *LEGACY_CONTRACT_HEADERS[:5], "공급가액", "부가세", "면세 여부",
    "업체별 연간 누적 횟수", "업체별 연간 누적 공급가액",
    "계약 상태", "회계연도",
)
SUMMARY_HEADERS = (
    "회계연도", "업체명", "사업자번호", "전체 계약 건수", "연간 누적 횟수",
    "연간 누적 공급가액", "남은 횟수", "남은 공급가액", "한도 상태",
)
EXEMPTION_VALUES = {label: exempt for exempt, label in EXEMPTION_LABELS.items()}


class StorageError(Exception):
    pass


class FileChangedError(StorageError):
    pass


def write_row(sheet, row_number: int, values: tuple) -> None:
    for column, value in enumerate(values, start=1):
        cell = sheet.cell(row_number, column, value)
        if isinstance(value, str):
            cell.data_type = "s"


def format_sheet(sheet, widths: tuple[int, ...], amount_columns: tuple[int, ...]) -> None:
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    sheet.row_dimensions[1].height = 24
    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor="183248")
        cell.font = Font(name="맑은 고딕", bold=True, color="FFFFFF")
    for column, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(column)].width = width
    body_font = Font(name="맑은 고딕")
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
        for cell in row:
            cell.font = body_font
        for column in amount_columns:
            row[column - 1].number_format = '#,##0'


def write_contract_sheet(sheet, ledger: Ledger, contracts: Iterable[Contract]) -> None:
    sheet.delete_rows(1, sheet.max_row)
    write_row(sheet, 1, CONTRACT_HEADERS)
    for row_number, contract in enumerate(contracts, start=2):
        totals = ledger.companies[contract.key].totals
        write_row(sheet, row_number, (
            contract.number, contract.signed_on.isoformat() if contract.signed_on else None, contract.company,
            business_display(contract.business_number), contract.title, contract.supply_amount,
            contract.vat_amount, EXEMPTION_LABELS[contract.tax_exempt],
            totals.count, totals.amount,
            contract.contract_status, contract.fiscal_year,
        ))
    format_sheet(sheet, (10, 15, 24, 18, 40, 17, 17, 14, 24, 28, 15, 12), (6, 7, 10))


def write_summary_sheet(summary, ledger: Ledger) -> None:
    summary.delete_rows(1, summary.max_row)
    write_row(summary, 1, SUMMARY_HEADERS)
    for row_number, ((year, business_number), company) in enumerate(sorted(ledger.companies.items()), start=2):
        totals = company.totals
        write_row(summary, row_number, (
            year, company.latest_contract.company, business_display(business_number),
            company.contract_count, totals.count, totals.amount,
            max(0, COUNT_LIMIT - totals.count), max(0, AMOUNT_LIMIT - totals.amount),
            totals.status,
        ))
    format_sheet(summary, (12, 24, 18, 18, 18, 22, 13, 20, 15), (6, 8))


def managed_year_sheets(workbook: Workbook) -> dict[int, str]:
    if SETTINGS_SHEET not in workbook.sheetnames:
        return {}
    settings = workbook[SETTINGS_SHEET]
    rows = settings.iter_rows(min_row=2, max_col=2, values_only=True)
    if next(rows, None) != YEAR_SHEET_HEADERS:
        return {}
    known_names = set(workbook.sheetnames) - {CONTRACT_SHEET, SUMMARY_SHEET, SETTINGS_SHEET}
    return {year: name for year, name in rows if type(year) is int and name in known_names}

def needs_aggregation_refresh(workbook: Workbook) -> bool:
    return (SETTINGS_SHEET not in workbook.sheetnames
            or workbook[SETTINGS_SHEET]["E1"].value != AGGREGATION_RULES_VERSION)


def write_workbook(ledger: Ledger, path: Path, previous: bytes | None = None, *,
                   layout_only: bool = False) -> None:
    workbook = Workbook() if previous is None else load_workbook(BytesIO(previous))
    try:
        if previous is None:
            workbook.active.title = CONTRACT_SHEET
        contract_sheet = workbook[CONTRACT_SHEET]
        upgrade_tax_columns = contract_sheet["F1"].value == LEGACY_CONTRACT_HEADERS[5]
        refresh_totals = needs_aggregation_refresh(workbook)
        refresh_contracts = (not layout_only or upgrade_tax_columns or refresh_totals
                             or (contract_sheet["K1"].value, contract_sheet["L1"].value) != CONTRACT_HEADERS[10:12])
        if refresh_contracts:
            write_contract_sheet(contract_sheet, ledger, ledger.contracts)
        summary = workbook[SUMMARY_SHEET] if SUMMARY_SHEET in workbook.sheetnames else None
        has_standard_summary = summary is None or tuple(cell.value for cell in summary[1][:5]) == SUMMARY_HEADERS[:5]
        if not layout_only or (refresh_totals and has_standard_summary):
            if summary is None:
                summary = workbook.create_sheet(SUMMARY_SHEET)
            write_summary_sheet(summary, ledger)
        elif upgrade_tax_columns and summary is not None:
            for cell, old_label, new_label in (("F1", "연간 누적 금액", SUMMARY_HEADERS[5]),
                                               ("H1", "남은 금액", SUMMARY_HEADERS[7])):
                if summary[cell].value == old_label:
                    summary[cell].value = new_label

        year_sheets = managed_year_sheets(workbook)
        year_contracts: dict[int, list[Contract]] = {year: [] for year in ledger.fiscal_years}
        for contract in ledger.contracts:
            year_contracts[contract.fiscal_year].append(contract)

        for year, name in year_sheets.items():
            if year not in year_contracts:
                workbook.remove(workbook[name])
        saved_names = {}
        for year, contracts in year_contracts.items():
            if year in year_sheets:
                sheet = workbook[year_sheets[year]]
                if refresh_contracts:
                    write_contract_sheet(sheet, ledger, contracts)
            else:
                name = f"{year}년"
                sheet = workbook.create_sheet(f"{name} 계약" if name in workbook.sheetnames else name)
                write_contract_sheet(sheet, ledger, contracts)
            saved_names[year] = sheet.title

        settings = workbook[SETTINGS_SHEET] if SETTINGS_SHEET in workbook.sheetnames else workbook.create_sheet(SETTINGS_SHEET)
        settings.delete_rows(1, settings.max_row)
        write_row(settings, 1, ("다음 순번", ledger.next_number))
        settings["D1"] = "집계 기준 버전"
        settings["E1"] = AGGREGATION_RULES_VERSION
        write_row(settings, 2, YEAR_SHEET_HEADERS)
        for row_number, info in enumerate(saved_names.items(), start=3):
            write_row(settings, row_number, info)
        settings.sheet_state = "hidden"
        workbook.save(path)
    finally:
        workbook.close()


def read_workbook(content: bytes) -> tuple[Ledger, bool, bool]:
    try:
        workbook = load_workbook(BytesIO(content), read_only=True, keep_links=False)
        try:
            if CONTRACT_SHEET not in workbook.sheetnames:
                raise StorageError(f"엑셀에 '{CONTRACT_SHEET}' 시트가 없습니다.")
            sheet = workbook[CONTRACT_SHEET]
            rows = sheet.iter_rows(max_col=len(CONTRACT_HEADERS), values_only=True)
            header = next(rows, ())
            legacy_amounts = header[:6] == LEGACY_CONTRACT_HEADERS[:6]
            has_plan_columns = header[10:12] == CONTRACT_HEADERS[10:12]
            if not legacy_amounts and header[:8] != CONTRACT_HEADERS[:8]:
                raise StorageError("엑셀의 계약 대장 열 이름 또는 순서가 다릅니다.")

            contracts = []
            numbers = set()
            for row_number, row in enumerate(rows, start=2):
                if all(value is None for value in row):
                    continue
                try:
                    number = row[0]
                    if type(number) is not int or number < 1 or number in numbers:
                        raise ValueError("순번은 중복되지 않는 양의 정수여야 합니다.")
                    vat_amount, tax_exempt = None, None
                    if not legacy_amounts:
                        vat_amount, status = row[6:8]
                        if status not in EXEMPTION_VALUES:
                            raise ValueError("면세 여부는 예, 아니오, 미확인 중 하나여야 합니다.")
                        tax_exempt = EXEMPTION_VALUES[status]
                        if tax_exempt is False and vat_amount is None:
                            raise ValueError("부가세를 입력하세요. 부가세가 0원이면 0을 입력하세요.")
                    contract_status, year = row[10:12] if has_plan_columns else (None, None)
                    if contract_status not in (None, "계약 계획", "계약 완료"):
                        raise ValueError("계약 상태는 계약 계획 또는 계약 완료여야 합니다.")
                    contract = Contract.from_fields(*row[:6], vat_amount=vat_amount, tax_exempt=tax_exempt,
                                                    planned=contract_status == "계약 계획", fiscal_year=year)
                except ValueError as error:
                    raise StorageError(f"계약 대장 {row_number}행\n{error}") from error
                contracts.append(contract)
                numbers.add(number)

            next_number = max(numbers, default=0) + 1
            if SETTINGS_SHEET in workbook.sheetnames:
                saved_next = workbook[SETTINGS_SHEET]["B1"].value
                if type(saved_next) is int:
                    next_number = max(next_number, saved_next)
            ledger = Ledger(tuple(sorted(contracts, key=lambda row: row.number)), next_number)
            needs_layout_update = (not has_plan_columns or needs_aggregation_refresh(workbook)
                                   or not set(ledger.fiscal_years).issubset(managed_year_sheets(workbook)))
            return ledger, needs_layout_update, legacy_amounts
        finally:
            workbook.close()
    except (BadZipFile, InvalidFileException, ParseError, KeyError, ValueError, TypeError) as error:
        raise StorageError("계약대장.xlsx를 읽을 수 없습니다. 파일 형식과 백업을 확인하세요.") from error


@contextmanager
def temporary_file(directory: Path) -> Iterator[Path]:
    fd, name = tempfile.mkstemp(dir=directory, prefix=".contract-", suffix=".xlsx")
    os.close(fd)
    path = Path(name)
    try:
        yield path
    finally:
        path.unlink(missing_ok=True)


class WorkbookStore:
    def __init__(self, path: Path):
        self.path = path
        self.backup_path = path.with_name(f"{path.stem}.backup.xlsx")
        self.lock_path = path.with_name(f".{path.stem}.lock")
        self.revision: bytes | None = None

    @contextmanager
    def lock(self) -> Iterator[None]:
        try:
            handle = self.lock_path.open("a+b")
        except OSError as error:
            raise StorageError("이 폴더에 쓸 수 없습니다. 프로그램을 쓰기 가능한 폴더로 옮겨 주세요.") from error
        with handle:
            try:
                if os.name == "nt":
                    import msvcrt
                    handle.seek(0)
                    msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
                else:
                    import fcntl
                    fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            except OSError as error:
                raise StorageError("같은 계약 대장이 다른 계약관리 창에서 열려 있습니다.") from error
            yield

    def _read(self) -> bytes | None:
        try:
            return self.path.read_bytes()
        except FileNotFoundError:
            return None
        except PermissionError as error:
            raise StorageError("계약대장.xlsx에 접근할 수 없습니다. 엑셀 창을 닫고 다시 시도하세요.") from error
        except OSError as error:
            raise StorageError(f"계약대장.xlsx를 열 수 없습니다.\n{error}") from error

    def load(self) -> Ledger:
        content = self._read()
        if content is None:
            ledger = Ledger()
            self.save(ledger)
            return ledger
        ledger, needs_layout_update, legacy_amounts = read_workbook(content)
        self.revision = sha256(content).digest()
        if legacy_amounts or needs_layout_update:
            self.save(ledger, layout_only=True)
        return ledger

    def save(self, ledger: Ledger, *, layout_only: bool = False) -> None:
        try:
            previous = self._read()
            self._check_revision(previous)
            with temporary_file(self.path.parent) as pending:
                write_workbook(ledger, pending, previous, layout_only=layout_only)
                with pending.open("rb+") as staged:
                    new_revision = sha256(staged.read()).digest()
                    os.fsync(staged.fileno())

                self._check_revision(self._read())

                if previous is not None:
                    with temporary_file(self.path.parent) as backup:
                        with backup.open("wb") as staged:
                            staged.write(previous)
                            staged.flush()
                            os.fsync(staged.fileno())
                        os.replace(backup, self.backup_path)
                os.replace(pending, self.path)
                self.revision = new_revision
        except PermissionError as error:
            raise StorageError(
                "계약 대장을 저장할 수 없습니다. 엑셀에서 대장 또는 백업 파일을 열어 둔 경우 닫고 다시 시도하세요."
            ) from error
        except OSError as error:
            raise StorageError(f"계약 대장을 저장할 수 없습니다.\n{error}") from error
        except (BadZipFile, InvalidFileException, ParseError, KeyError, ValueError, TypeError) as error:
            raise StorageError("계약대장.xlsx를 저장할 수 없습니다. 파일 형식과 백업을 확인하세요.") from error

    def _check_revision(self, content: bytes | None) -> None:
        current_revision = sha256(content).digest() if content is not None else None
        if current_revision != self.revision:
            raise FileChangedError("계약대장.xlsx가 외부에서 변경되었습니다. 프로그램을 다시 열어 최신 내용을 불러오세요.")
