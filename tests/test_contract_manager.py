from datetime import date
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook, load_workbook

from contracts import Contract, Ledger, Totals
from storage import (
    CONTRACT_HEADERS, CONTRACT_SHEET, LEGACY_CONTRACT_HEADERS, SETTINGS_SHEET, SUMMARY_SHEET,
    FileChangedError, StorageError, WorkbookStore,
)


BUSINESS = "012-34-56789"


def contract(number, amount, signed_on="2026-02-01", business=BUSINESS, company="한빛"):
    return Contract.from_fields(number, signed_on, company, business, f"계약 {number}", amount)


class LedgerTests(unittest.TestCase):
    def test_boundaries(self):
        with self.assertRaisesRegex(ValueError, "0원"):
            contract(1, 0)
        ledger = Ledger()
        cases = (
            (1_999_999, Totals(0, 0)),
            (2_000_000, Totals(2_000_000, 0)),
            (4_999_999, Totals(6_999_999, 0)),
            (5_000_000, Totals(11_999_999, 1)),
            (5_000_001, Totals(17_000_000, 2)),
        )
        for number, (amount, expected) in enumerate(cases, start=1):
            entry = contract(number, amount)
            ledger = ledger.upsert(entry)
            self.assertEqual(ledger.companies[entry.key].totals, expected)

    def test_count_limit_blocks_new_eligible_contracts(self):
        ledger = Ledger()
        for number in range(1, 6):
            ledger = ledger.upsert(contract(number, 6_000_000))
        self.assertEqual(ledger.companies[(2026, "0123456789")].totals.status, "한도 도달")
        with self.assertRaisesRegex(ValueError, "연간 한도 도달"):
            ledger.upsert(contract(6, 2_000_000))
        ledger = ledger.upsert(contract(6, 1_999_999))
        self.assertEqual(ledger.companies[(2026, "0123456789")].totals, Totals(30_000_000, 5))
        with self.assertRaisesRegex(ValueError, "횟수 1회 초과"):
            ledger.upsert(contract(7, 5_000_000))

    def test_recounted_over_limit_record_can_be_corrected(self):
        rows = tuple(contract(number, 6_000_000) for number in range(1, 8))
        ledger = Ledger(rows, 8)
        reduced = ledger.upsert(contract(1, 4_999_999))
        self.assertEqual(reduced.companies[rows[0].key].totals, Totals(40_999_999, 6))
        with self.assertRaisesRegex(ValueError, "횟수 2회 초과"):
            reduced.upsert(rows[0])

    def test_amount_limit_and_both_violations(self):
        ledger = Ledger()
        for number in range(1, 6):
            ledger = ledger.upsert(contract(number, 40_000_000))
        totals = ledger.companies[(2026, "0123456789")].totals
        self.assertEqual(totals, Totals(200_000_000, 5))
        self.assertEqual(totals.status, "한도 도달")
        with self.assertRaisesRegex(ValueError, r"공급가액 2,000,001원 초과$"):
            ledger.upsert(contract(6, 2_000_001))
        with self.assertRaisesRegex(ValueError, "공급가액 5,000,001원 초과 · 횟수 1회 초과"):
            ledger.upsert(contract(6, 5_000_001))
        with self.assertRaisesRegex(ValueError, "공급가액 1원 초과"):
            Ledger().upsert(contract(1, 200_000_001))

    def test_edit_move_and_delete_recompute_both_groups(self):
        ledger = Ledger().upsert(contract(1, 5_000_001))
        ledger = ledger.upsert(contract(2, 4_000_000, company="한빛 주식회사"))
        changed = contract(1, 4_999_999)
        self.assertEqual(ledger.preview(changed), (Totals(9_000_001, 1), Totals(8_999_999, 0)))
        ledger = ledger.upsert(changed)
        moved = contract(1, 9_000_000, "2025-12-31", business="999-99-99999")
        ledger = ledger.upsert(moved)
        self.assertEqual(ledger.companies[(2026, "0123456789")].totals, Totals(4_000_000, 0))
        self.assertEqual(ledger.companies[moved.key].totals, Totals(9_000_000, 1))
        ledger = ledger.delete(1)
        self.assertNotIn(moved.key, ledger.companies)
        self.assertEqual(ledger.delete(2).next_number, 3)

    def test_same_business_is_independent_between_years(self):
        ledger = Ledger().upsert(contract(1, 200_000_000, "2025-12-31"))
        ledger = ledger.upsert(contract(2, 6_000_000, "2026-01-01"))
        self.assertEqual(ledger.companies[(2026, "0123456789")].totals, Totals(6_000_000, 1))

    def test_vat_never_affects_contract_counts_or_annual_limits(self):
        fields = ("2026-02-01", "한빛", BUSINESS, "계약")
        small = Contract.from_fields(1, *fields, 4_999_999, vat_amount=900_000)
        big = Contract.from_fields(2, *fields, 195_000_001)
        ledger = Ledger().upsert(small).upsert(big)
        self.assertEqual((small.counted, big.vat_amount), (0, 19_500_000))
        self.assertEqual(ledger.companies[small.key].totals, Totals(200_000_000, 1))

        exempt = Contract.from_fields(2, *fields, 195_000_001, tax_exempt=True)
        self.assertEqual((exempt.vat_amount, exempt.counted), (0, 1))
        self.assertEqual(ledger.preview(exempt), (Totals(200_000_000, 1), Totals(200_000_000, 1)))
        self.assertEqual(contract(1, 5_000_009).vat_amount, 500_000)
        with self.assertRaisesRegex(ValueError, "면세 계약의 부가세는 0원"):
            Contract.from_fields(2, *fields, 6_000_000, vat_amount=1, tax_exempt=True)
        with self.assertRaisesRegex(ValueError, "부가세에 원 단위 정수"):
            Contract.from_fields(2, *fields, 6_000_000, vat_amount="-1")

    def test_input_validation(self):
        for amount in ("2,00,001", "2500000.5"):
            with self.subTest(amount=amount), self.assertRaisesRegex(ValueError, "정수"):
                contract(1, amount)
        for signed_on in ("2026-02-29", "20260201"):
            with self.subTest(signed_on=signed_on), self.assertRaisesRegex(ValueError, "YYYY-MM-DD"):
                contract(1, 2_500_000, signed_on)
        with self.assertRaisesRegex(ValueError, "사업자번호"):
            contract(1, 2_500_000, business="1234")


class WorkbookTests(unittest.TestCase):
    def setUp(self):
        directory = TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.path = Path(directory.name) / "계약대장.xlsx"
        self.store = WorkbookStore(self.path)

    def test_round_trip_and_backup(self):
        with self.store.lock():
            empty = self.store.load()
            previous_file = self.path.read_bytes()
            original = Contract.from_fields(1, "2026-07-09", "한빛", BUSINESS, "=1+1", 4_000_000, vat_amount=123_456)
            exempt = Contract.from_fields(2, "2026-07-09", "한빛", BUSINESS, "면세 계약", 6_000_000, tax_exempt=True)
            ledger = empty.upsert(original).upsert(exempt)
            self.store.save(ledger)
            self.assertEqual(self.store.backup_path.read_bytes(), previous_file)

        workbook = load_workbook(self.path)
        self.addCleanup(workbook.close)
        sheet = workbook[CONTRACT_SHEET]
        self.assertEqual(sheet["D2"].value, BUSINESS)
        self.assertEqual((sheet["E2"].value, sheet["E2"].data_type), ("=1+1", "s"))
        for sheet in (sheet, workbook["2026년"]):
            self.assertEqual(tuple(cell.value for cell in sheet[2][5:]), (4_000_000, 123_456, "아니오", 1, 10_000_000))
            self.assertEqual(tuple(cell.value for cell in sheet[3][5:]), (6_000_000, 0, "예", 1, 10_000_000))
        self.assertEqual(workbook[SUMMARY_SHEET]["F2"].value, 10_000_000)
        reloaded = WorkbookStore(self.path).load()
        self.assertEqual(reloaded.contracts, ledger.contracts)
        self.assertEqual(reloaded.next_number, 3)

        self.store.save(ledger.delete(2))
        self.assertEqual(self.store.load().next_number, 3)

    def test_existing_workbook_recalculates_inclusive_count_boundary_once(self):
        ledger = self.store.load().upsert(contract(1, 5_000_000))
        self.store.save(ledger)
        previous = load_workbook(BytesIO(self.path.read_bytes()))
        previous[SETTINGS_SHEET]["E1"] = None
        previous[CONTRACT_SHEET]["I2"] = 0
        previous[SUMMARY_SHEET]["E2"] = 0
        previous["2026년"]["I2"] = 0
        previous.save(self.path)
        previous.close()
        old_content = self.path.read_bytes()

        self.assertEqual(self.store.load().contracts, ledger.contracts)
        self.assertEqual(self.store.backup_path.read_bytes(), old_content)
        refreshed_content = self.path.read_bytes()
        workbook = load_workbook(BytesIO(refreshed_content))
        self.addCleanup(workbook.close)
        self.assertEqual(workbook[CONTRACT_SHEET]["I2"].value, 1)
        self.assertEqual(workbook[SUMMARY_SHEET]["E2"].value, 1)
        self.assertEqual(workbook["2026년"]["I2"].value, 1)
        self.store.load()
        self.assertEqual(self.path.read_bytes(), refreshed_content)

    def test_new_workbook_contains_empty_fiscal_year_tabs(self):
        self.store.load()
        workbook = load_workbook(BytesIO(self.path.read_bytes()))
        self.addCleanup(workbook.close)
        years = sorted({*range(2025, 2031), date.today().year})
        self.assertEqual(workbook.sheetnames, [
            CONTRACT_SHEET, SUMMARY_SHEET, *(f"{year}년" for year in years), SETTINGS_SHEET,
        ])
        sheet = workbook["2030년"]
        self.assertEqual(tuple(sheet.values), (CONTRACT_HEADERS,))
        self.assertEqual((sheet.freeze_panes, sheet.auto_filter.ref), ("A2", "A1:J1"))

    def test_year_tabs_recalculate_when_contracts_move_or_are_deleted(self):
        ledger = self.store.load()
        ledger = ledger.upsert(contract(1, 6_000_000, "2025-12-31"))
        ledger = ledger.upsert(contract(2, 4_000_000, "2026-01-01"))
        ledger = ledger.upsert(contract(3, 5_000_000, "2024-12-31"))
        self.store.save(ledger)
        workbook = load_workbook(BytesIO(self.path.read_bytes()))
        self.addCleanup(workbook.close)
        self.assertEqual(list(workbook["2025년"].values)[1][5:], (6_000_000, 600_000, "아니오", 1, 6_000_000))
        self.assertEqual(list(workbook["2026년"].values)[1][5:], (4_000_000, 400_000, "아니오", 0, 4_000_000))
        self.assertEqual(workbook["2024년"]["A2"].value, 3)
        self.assertEqual(self.store.load().contracts, ledger.contracts)

        moved = ledger.upsert(contract(1, 6_000_000, "2026-02-01")).delete(3)
        self.store.save(moved)
        workbook = load_workbook(BytesIO(self.path.read_bytes()))
        self.addCleanup(workbook.close)
        self.assertEqual(tuple(workbook["2025년"].values), (CONTRACT_HEADERS,))
        self.assertEqual([row[8:] for row in list(workbook["2026년"].values)[1:]],
                         [(1, 10_000_000), (1, 10_000_000)])
        if date.today().year != 2024:
            self.assertNotIn("2024년", workbook.sheetnames)

    def test_legacy_workbook_adds_year_tabs_without_replacing_existing_sheets(self):
        legacy = Workbook()
        legacy.active.title = CONTRACT_SHEET
        legacy.active.append(LEGACY_CONTRACT_HEADERS)
        legacy.active.append((8, "2025-10-01", "한빛", BUSINESS, "기존 계약", 6_000_000, 1, 6_000_000))
        legacy.create_sheet(SUMMARY_SHEET)["A1"] = "기존 요약"
        legacy.create_sheet("2025년")["A1"] = "직접 작성한 메모"
        settings = legacy.create_sheet(SETTINGS_SHEET)
        settings.append(("다음 순번", 11))
        settings.sheet_state = "hidden"
        legacy.save(self.path)
        legacy.close()
        original = self.path.read_bytes()

        ledger = self.store.load()
        upgraded = self.path.read_bytes()
        self.assertEqual(self.store.backup_path.read_bytes(), original)
        self.assertEqual(ledger.next_number, 11)
        self.assertEqual((ledger.contracts[0].supply_amount, ledger.contracts[0].vat_amount,
                          ledger.contracts[0].tax_exempt), (6_000_000, None, None))
        self.assertEqual(self.store.load().contracts, ledger.contracts)
        self.assertEqual(self.path.read_bytes(), upgraded)
        workbook = load_workbook(BytesIO(upgraded))
        self.addCleanup(workbook.close)
        self.assertEqual(workbook["2025년"]["A1"].value, "직접 작성한 메모")
        self.assertEqual(workbook[SUMMARY_SHEET]["A1"].value, "기존 요약")
        self.assertEqual(workbook["2025년 계약"]["E2"].value, "기존 계약")
        self.assertEqual(tuple(cell.value for cell in workbook["2025년 계약"][2][5:8]),
                         (6_000_000, None, "미확인"))

        self.store.save(ledger.upsert(contract(11, 4_000_000, "2025-11-01")))
        workbook = load_workbook(BytesIO(self.path.read_bytes()))
        self.addCleanup(workbook.close)
        self.assertEqual(workbook["2025년"]["A1"].value, "직접 작성한 메모")
        self.assertEqual(workbook["2025년 계약"]["J3"].value, 10_000_000)

    def test_changed_file_is_preserved(self):
        ledger = self.store.load()
        external_content = b"externally replaced"
        self.path.write_bytes(external_content)
        with self.assertRaises(FileChangedError):
            self.store.save(ledger.upsert(contract(1, 5_000_000)))
        self.assertEqual(self.path.read_bytes(), external_content)
        self.assertFalse(self.store.backup_path.exists())

    def test_corrupt_workbook_is_preserved(self):
        self.path.write_bytes(b"not an xlsx")
        with self.assertRaises(StorageError):
            self.store.load()
        self.assertEqual(self.path.read_bytes(), b"not an xlsx")

    def test_same_workbook_cannot_be_opened_twice(self):
        with self.store.lock():
            with self.assertRaisesRegex(StorageError, "다른 계약관리 창"), WorkbookStore(self.path).lock():
                self.fail("second lock succeeded")
        with WorkbookStore(self.path).lock():
            pass


if __name__ == "__main__":
    unittest.main()
