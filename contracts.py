from dataclasses import dataclass
from datetime import date, datetime
import re


AMOUNT_THRESHOLD = 2_000_000
COUNT_THRESHOLD = 5_000_000
AMOUNT_LIMIT = 200_000_000
COUNT_LIMIT = 5
DEFAULT_FISCAL_YEARS = range(2025, 2031)
EXEMPTION_LABELS = {None: "미확인", False: "아니오", True: "예"}
MAX_EXCEL_INTEGER = 999_999_999_999_999

INVALID_EXCEL_TEXT = re.compile(r"[\x00-\x08\x0b-\x0c\x0e-\x1f\ufffe\uffff]")


def won(amount: int) -> str:
    return f"{amount:,}원"


def business_display(number: str) -> str:
    return f"{number[:3]}-{number[3:5]}-{number[5:]}"


def business_digits(value: str) -> str:
    return re.sub(r"[\s-]", "", value)


def required_text(value: object, label: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"{label} 항목을 입력하세요.")
    value = value.strip()
    if len(value) > 32_767 or INVALID_EXCEL_TEXT.search(value):
        raise ValueError(f"{label} 항목에 엑셀에 저장할 수 없는 문자가 있거나 입력 길이가 너무 깁니다.")
    return value


def parse_won(value: object, label: str) -> int:
    text = str(value).strip()
    if not re.fullmatch(r"(?:[0-9]+|[0-9]{1,3}(?:,[0-9]{3})+)", text):
        raise ValueError(f"{label}에 원 단위 정수를 입력하세요.")
    try:
        return int(text.replace(",", ""))
    except ValueError:
        raise ValueError(f"{label} 입력값이 너무 큽니다.") from None


def default_vat(amount: int) -> int:
    return amount // 10


@dataclass(frozen=True)
class Contract:
    number: int
    signed_on: date | None
    company: str
    business_number: str
    title: str
    supply_amount: int
    vat_amount: int | None = None
    tax_exempt: bool | None = None
    planned: bool = False
    undated_year: int | None = None

    @property
    def fiscal_year(self) -> int:
        if self.signed_on is not None:
            return self.signed_on.year
        assert self.undated_year is not None
        return self.undated_year

    @property
    def date_order(self) -> tuple[date, int]:
        return self.signed_on or date.min, self.number

    @property
    def contract_status(self) -> str:
        return "계약 계획" if self.planned else "계약 완료"

    @property
    def total_amount(self) -> int:
        return self.supply_amount + (self.vat_amount or 0)

    @property
    def key(self) -> tuple[int, str]:
        return self.fiscal_year, self.business_number

    @property
    def counted(self) -> int:
        return int(self.supply_amount >= COUNT_THRESHOLD)

    @property
    def amount_for_limit(self) -> int:
        return self.supply_amount if self.supply_amount >= AMOUNT_THRESHOLD else 0

    @classmethod
    def from_fields(
        cls,
        number: int,
        signed_on: object,
        company: object,
        business_number: object,
        title: object,
        supply_amount: object,
        vat_amount: object = None,
        tax_exempt: bool | None = False,
        *,
        planned: bool = False,
        fiscal_year: int | None = None,
    ) -> "Contract":
        if signed_on is None or isinstance(signed_on, str) and not signed_on.strip():
            if not planned:
                raise ValueError("완료 계약은 계약 일자를 입력하세요.")
            if type(fiscal_year) is not int or not 1 <= fiscal_year <= 9999:
                raise ValueError("날짜 없는 계약 계획에는 회계연도가 필요합니다.")
            signed_on = None
        elif isinstance(signed_on, datetime):
            signed_on = signed_on.date()
        elif not isinstance(signed_on, date):
            try:
                date_text = str(signed_on).strip()
                if not re.fullmatch(r"[0-9]{4}-[0-9]{2}-[0-9]{2}", date_text):
                    raise ValueError
                signed_on = date.fromisoformat(date_text)
            except ValueError:
                raise ValueError("계약 일자를 YYYY-MM-DD 형식으로 입력하세요.") from None

        company = required_text(company, "업체명")
        business_number = business_digits(str(business_number))
        if not re.fullmatch(r"[0-9]{10}", business_number):
            raise ValueError("사업자번호를 숫자 10자리로 입력하세요.")
        title = required_text(title, "계약명")

        supply_amount = parse_won(supply_amount, "공급가액")
        if supply_amount == 0:
            raise ValueError("공급가액은 0원을 초과해야 합니다.")

        if tax_exempt is None:
            if vat_amount is not None:
                raise ValueError("부가세가 있는 계약의 면세 여부를 확인하세요.")
        elif tax_exempt:
            if vat_amount is not None and parse_won(vat_amount, "부가세") != 0:
                raise ValueError("면세 계약의 부가세는 0원이어야 합니다.")
            vat_amount = 0
        else:
            vat_amount = default_vat(supply_amount) if vat_amount is None else parse_won(vat_amount, "부가세")
            if vat_amount > MAX_EXCEL_INTEGER:
                raise ValueError("부가세를 최대 15자리까지 입력하세요.")

        return cls(number, signed_on, company, business_number, title, supply_amount, vat_amount,
                   tax_exempt, planned, fiscal_year if signed_on is None else None)


@dataclass(frozen=True)
class Totals:
    amount: int = 0
    count: int = 0

    def add(self, contract: Contract) -> "Totals":
        return Totals(self.amount + contract.amount_for_limit, self.count + contract.counted)

    def remove(self, contract: Contract) -> "Totals":
        return Totals(self.amount - contract.amount_for_limit, self.count - contract.counted)

    def violations_since(self, before: "Totals") -> tuple[str, ...]:
        details = []
        if self.amount > max(AMOUNT_LIMIT, before.amount):
            details.append(f"공급가액 {won(self.amount - AMOUNT_LIMIT)} 초과")
        if self.count > max(COUNT_LIMIT, before.count):
            details.append(f"횟수 {self.count - COUNT_LIMIT}회 초과")
        return tuple(details)

    @property
    def status(self) -> str:
        if self.amount > AMOUNT_LIMIT or self.count > COUNT_LIMIT:
            return "한도 초과"
        if self.amount == AMOUNT_LIMIT or self.count == COUNT_LIMIT:
            return "한도 도달"
        return "정상"


@dataclass(frozen=True)
class CompanyYear:
    latest_contract: Contract
    totals: Totals
    contract_count: int


class Ledger:
    def __init__(self, contracts: tuple[Contract, ...] = (), next_number: int = 1):
        self.contracts = contracts
        self.next_number = next_number
        self.by_number = {contract.number: contract for contract in contracts}
        self.companies: dict[tuple[int, str], CompanyYear] = {}
        for contract in contracts:
            existing = self.companies.get(contract.key)
            if existing is None:
                self.companies[contract.key] = CompanyYear(contract, Totals().add(contract), 1)
                continue

            latest = existing.latest_contract
            if contract.date_order > latest.date_order:
                latest = contract
            self.companies[contract.key] = CompanyYear(
                latest, existing.totals.add(contract), existing.contract_count + 1
            )

    @property
    def fiscal_years(self) -> tuple[int, ...]:
        return tuple(sorted({*DEFAULT_FISCAL_YEARS, date.today().year, *(year for year, _ in self.companies)}))

    def preview(self, contract: Contract) -> tuple[Totals, Totals]:
        company = self.companies.get(contract.key)
        before = company.totals if company else Totals()
        after = before
        previous = self.by_number.get(contract.number)
        if previous is not None and previous.key == contract.key:
            after = after.remove(previous)
        return before, after.add(contract)

    def limit_violations(self, contract: Contract, before: Totals, after: Totals) -> tuple[str, ...]:
        if not contract.amount_for_limit:
            return ()
        previous = self.by_number.get(contract.number)
        if previous and previous.key == contract.key and previous.amount_for_limit:
            return after.violations_since(before)
        violations = after.violations_since(Totals())
        if violations:
            return violations
        if before.amount >= AMOUNT_LIMIT or before.count >= COUNT_LIMIT:
            return ("연간 한도 도달로 추가 계약을 등록할 수 없습니다.",)
        return ()

    def upsert(self, contract: Contract) -> "Ledger":
        if contract.number not in self.by_number and contract.number != self.next_number:
            raise ValueError("순번이 변경되었습니다. 계약을 다시 열어 주세요.")

        before, after = self.preview(contract)
        violations = self.limit_violations(contract, before, after)
        if violations:
            raise ValueError(" · ".join(violations))

        remaining = tuple(row for row in self.contracts if row.number != contract.number)
        return Ledger(
            tuple(sorted((*remaining, contract), key=lambda row: row.number)),
            max(self.next_number, contract.number + 1),
        )

    def delete(self, number: int) -> "Ledger":
        return Ledger(tuple(row for row in self.contracts if row.number != number), self.next_number)
