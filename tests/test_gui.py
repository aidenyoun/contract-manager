from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory
import tkinter as tk
from tkinter import ttk
import unittest

from app import ContractApp, ContractEditor
from contracts import Contract, Ledger
from storage import WorkbookStore


def find_widgets(parent, widget_type):
    for child in parent.winfo_children():
        if isinstance(child, widget_type):
            yield child
        yield from find_widgets(child, widget_type)


class ContractGuiTests(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError as error:
            self.skipTest(f"Tk display unavailable: {error}")
        folder = TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.addCleanup(self.root.destroy)
        self.year = date.today().year
        store = WorkbookStore(Path(folder.name) / "계약대장.xlsx")
        store.load()
        records = (
            (self.year - 1, "기존 업체명", "0123456789", 6_000_000),
            (self.year - 1, "이전 연도 업체", "8321045678", 5_000_000),
            (self.year, "한빛산업", "0123456789", 200_000_000),
            (self.year, "한빛산업", "9999999999", 4_000_000),
            (self.year, "초과 업체", "5132456789", 200_000_001),
        )
        contracts = tuple(
            Contract.from_fields(number, date(year, 1, number), company, business, f"계약 {number}", amount)
            for number, (year, company, business, amount) in enumerate(records, start=1)
        )
        ledger = Ledger(contracts, len(contracts) + 1)
        store.save(ledger)
        self.app = ContractApp(self.root, store, ledger)
        self.root.update()

    def test_dashboard_cards_follow_selected_year(self):
        cards = [button for button in find_widgets(self.root, ttk.Button)
                 if button.cget("style") == "Metric.TButton"]

        self.app.query.set("검색 결과 없음")
        self.app.tabs.select(1)
        cards[0].invoke()
        self.assertEqual(self.app.tabs.select(), str(self.app.contract_tab))
        self.assertEqual(set(self.app.contract_table.get_children()), {"3", "4", "5"})

        self.app.query.set("검색 결과 없음")
        cards[1].invoke()
        popup = self.root.grab_current()
        table, = find_widgets(popup, ttk.Treeview)
        self.assertEqual(set(table.get_children()), {"0123456789", "9999999999", "5132456789"})
        popup.destroy()

        cards[2].invoke()
        popup = self.root.grab_current()
        table, = find_widgets(popup, ttk.Treeview)
        self.assertEqual(set(table.get_children()), {"0123456789", "5132456789"})
        self.assertEqual(table.item("0123456789", "values")[-1], "한도 도달")
        self.assertEqual(table.item("5132456789", "values")[-1], "한도 초과")
        popup.destroy()

        self.app.year.set(str(self.year - 1))
        cards[2].invoke()
        popup = self.root.grab_current()
        labels = [label.cget("text") for label in find_widgets(popup, ttk.Label)]
        self.assertTrue(any("해당 업체가 없습니다." in label for label in labels))
        self.assertEqual(self.app.query.get(), "검색 결과 없음")

    def test_fiscal_year_list_includes_empty_years_and_older_records(self):
        earlier = Contract.from_fields(6, date(2024, 1, 1), "이전 계약", "5555555555", "기존 계약", 4_000_000)
        self.app.ledger = self.app.ledger.upsert(earlier)
        self.app.refresh()
        options = {int(year) for year in self.app.year_box["values"]}
        self.assertTrue({2024, *range(2025, 2031)}.issubset(options))
        self.app.year_box.set("2030")
        if self.year not in (2030, 2031):
            self.assertFalse(self.app.contract_table.get_children())
        editor = ContractEditor(self.app, None)
        target_date = date.today() if self.year == 2030 else date(2030, 1, 1)
        self.assertEqual(editor.variables[0].get(), target_date.isoformat())

    def test_contract_headers_sort_and_keep_direction_after_filtering(self):
        table = self.app.contract_table
        click_date = table.heading("column1", "command")
        click_amount = table.heading("column5", "command")
        self.assertEqual(table.heading("column5", "text"), "총금액")
        self.assertEqual(table.heading("column6", "text"), "공급가액")
        self.assertEqual(table.get_children(), ("5", "4", "3"))
        table.tk.call(click_date)
        self.assertEqual(table.get_children(), ("3", "4", "5"))
        self.assertTrue(table.heading("column1", "text").endswith("↑"))

        table.selection_set("3")
        table.tk.call(click_amount)
        self.assertEqual(table.get_children(), ("4", "3", "5"))
        self.assertEqual(table.selection(), ("3",))
        table.tk.call(click_amount)
        self.assertEqual(table.get_children(), ("5", "3", "4"))
        self.assertTrue(table.heading("column5", "text").endswith("↓"))
        self.assertFalse(table.heading("column1", "text").endswith("↑"))
        self.app.query.set("한빛")
        self.assertEqual(table.get_children(), ("3", "4"))
        self.assertEqual(table.selection(), ("3",))

    def test_company_headers_sort_amounts_and_limit_status(self):
        table = self.app.company_table
        click_amount = table.heading("column4", "command")
        click_status = table.heading("column7", "command")
        ascending = ("9999999999", "0123456789", "5132456789")

        table.tk.call(click_amount)
        self.assertEqual(table.get_children(), ascending)
        table.tk.call(click_amount)
        self.assertEqual(table.get_children(), ascending[::-1])
        table.tk.call(click_status)
        self.assertEqual(table.get_children(), ascending)
        self.assertTrue(table.heading("column7", "text").endswith("↑"))
        self.assertIn("한도 도달", table.item("0123456789", "tags"))
        self.assertIn("alternate", table.item("0123456789", "tags"))

    def test_history_is_searchable_by_name_and_number_across_years(self):
        editor = ContractEditor(self.app, None)
        box = editor.history_box
        box.set("한빛")
        editor.filter_company_history()
        matches = box.cget("values")
        self.assertEqual(len(matches), 2)
        self.assertTrue(all("기존 업체명" not in label for label in matches))

        box.set("832-10")
        editor.filter_company_history()
        _, = box.cget("values")
        box.current(0)
        box.event_generate("<<ComboboxSelected>>")
        self.assertEqual(editor.variables[1].get(), "이전 연도 업체")
        self.assertEqual(editor.variables[2].get(), "832-10-45678")

    def test_amount_formats_as_typed_and_saved_with_numeric_value(self):
        editor = ContractEditor(self.app, None)
        for variable, value in zip(editor.variables[1:4], ("신규 업체", "1011123456", "점검 계약")):
            variable.set(value)
        entry = editor.entries[-1]
        entry.focus_force()
        self.root.update()
        for digit in "5000001":
            entry.event_generate("<KeyPress>", keysym=digit)
        self.assertEqual((entry.get(), entry.index(tk.INSERT)), ("5,000,001", 9))
        self.assertEqual(editor.supply_display.get(), "4,545,456")
        self.assertIn("0회 / 4,545,456원", editor.feedback.cget("text"))

        entry.icursor(2)
        entry.event_generate("<KeyPress>", keysym="BackSpace")
        self.assertEqual(entry.get(), "1")

        entry.selection_range(0, tk.END)
        entry.event_generate("<KeyPress>", keysym="BackSpace")
        entry.insert(0, "12,345678")
        self.assertEqual(entry.get(), "12,345,678")
        entry.icursor(2)
        entry.event_generate("<KeyPress>", keysym="Delete")
        self.assertEqual(entry.get(), "1,245,678")
        self.assertTrue(editor.save_button.instate(["!disabled"]))
        self.assertIn("집계 제외", editor.feedback.cget("text"))

        editor.variables[-1].set("2500000")
        self.assertEqual(entry.get(), "2,500,000")
        self.assertTrue(editor.save_button.instate(["!disabled"]))
        editor.save_button.invoke()
        saved = self.app.store.load().contracts[-1]
        self.assertEqual((saved.supply_amount, saved.vat_amount, saved.counted), (2_272_728, 227_272, 0))
        table = self.app.contract_table
        self.assertEqual(table.item(str(saved.number), "values")[5:8], ("2,500,000원", "2,272,728원", "227,272원"))

    def test_vat_can_be_overridden_and_exempted_in_contract_editor(self):
        editor = ContractEditor(self.app, None)
        for variable, value in zip(editor.variables[1:], ("신규 업체", "1011123456", "세금 확인", "2310000")):
            variable.set(value)
        self.assertEqual(editor.vat_entry.get(), "210,000")
        self.assertEqual(editor.supply_display.get(), "2,100,000")
        self.assertIn("0회 / 2,100,000원", editor.feedback.cget("text"))
        editor.variables[-1].set("5000000")
        editor.vat_variable.set("5000001")
        self.assertTrue(editor.save_button.instate(["disabled"]))
        self.assertIn("총금액을 초과", editor.feedback.cget("text"))
        editor.vat_variable.set("345678")
        self.assertEqual(editor.vat_entry.get(), "345,678")
        self.assertEqual(editor.supply_display.get(), "4,654,322")
        editor.exemption_checkbox.invoke()
        self.assertEqual(editor.vat_entry.get(), "0")
        self.assertTrue(editor.vat_entry.instate(["readonly"]))
        editor.exemption_checkbox.invoke()
        self.assertEqual(editor.vat_entry.get(), "345,678")
        editor.save_button.invoke()
        original = self.app.store.load().contracts[-1]
        self.assertEqual((original.vat_amount, original.tax_exempt), (345_678, False))

        editor = ContractEditor(self.app, original)
        self.assertEqual(editor.vat_entry.get(), "345,678")
        self.assertEqual(editor.entries[-1].get(), "5,000,000")
        editor.exemption_checkbox.invoke()
        editor.variables[-1].set("6000000")
        self.assertEqual(editor.vat_entry.get(), "0")
        self.assertIn("1회 / 6,000,000원", editor.feedback.cget("text"))
        editor.exemption_checkbox.invoke()
        self.assertEqual(editor.vat_entry.get(), "545,454")
        editor.exemption_checkbox.invoke()
        editor.save_button.invoke()
        saved = self.app.store.load().contracts[-1]
        self.assertEqual((saved.vat_amount, saved.tax_exempt, saved.counted), (0, True, 1))

    def test_editing_legacy_title_does_not_invent_tax_details(self):
        original = Contract.from_fields(6, date(self.year, 2, 1), "이전 업체", "1011123456", "이전 계약",
                                        5_000_000, tax_exempt=None)
        self.app.ledger = self.app.ledger.upsert(original)
        editor = ContractEditor(self.app, original)
        self.assertEqual(editor.vat_entry.get(), "")
        editor.variables[3].set("수정한 계약명")
        editor.save_button.invoke()
        saved = self.app.store.load().contracts[-1]
        self.assertEqual((saved.title, saved.vat_amount, saved.tax_exempt), ("수정한 계약명", None, None))


if __name__ == "__main__":
    unittest.main()
