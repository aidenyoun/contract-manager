from collections.abc import Callable, Iterable
from dataclasses import dataclass
from datetime import date
import json
import os
from pathlib import Path
import sys
import tkinter as tk
from tkinter import font, messagebox, ttk

from contracts import (
    AMOUNT_LIMIT, AMOUNT_THRESHOLD, COUNT_LIMIT, COUNT_THRESHOLD, EXEMPTION_LABELS, CompanyYear,
    Contract, Ledger, business_digits, business_display, parse_won, won,
)
from storage import CONTRACT_HEADERS, StorageError, WorkbookStore


BACKGROUND = "#edf2f6"
PANEL = "#ffffff"
INK = "#152b3d"
MUTED = "#506575"
BORDER = "#d4dfe7"
NAVY = "#183248"
TEAL = "#007c70"
DANGER = "#a22531"

VERSION = "v9.0"
APP_TITLE = f"계약 한도 관리 · {VERSION}"
METRIC_TITLES = ("전체 계약", "계약 업체", "한도 도달·초과 업체", "계약 계획 현황")


def numeric_value(text: str) -> int:
    return int(text.rstrip("원회건").replace(",", ""))


@dataclass(frozen=True)
class Column:
    title: str
    width: int
    anchor: str
    sort_key: Callable[[str], str | int] = str.casefold


CONTRACT_COLUMNS = tuple(Column(label, *settings) for label, settings in zip(
    (*CONTRACT_HEADERS[:2], "계약 상태", *CONTRACT_HEADERS[2:5], "총금액", *CONTRACT_HEADERS[5:10]),
    (
        (51, "center", numeric_value), (98, "center"), (92, "center"), (122, "w"), (110, "center"),
        (210, "w"), (107, "e", numeric_value), (101, "e", numeric_value),
        (97, "e", lambda text: -1 if text == "—" else numeric_value(text)), (90, "center"),
        (150, "e", numeric_value), (174, "e", numeric_value),
    ),
))
COMPANY_COLUMNS = (
    Column("업체명", 170, "w"), Column("사업자번호", 126, "center"),
    Column("전체 계약 건수", 102, "e", numeric_value), Column("연간 누적 횟수", 120, "e", numeric_value),
    Column("연간 누적 공급가액", 150, "e", numeric_value), Column("남은 횟수", 100, "e", numeric_value),
    Column("남은 공급가액", 136, "e", numeric_value),
    Column("한도 상태", 100, "center", ("정상", "한도 도달", "한도 초과").index),
)


def configure_style(root: tk.Tk) -> None:
    families = font.families(root)
    family = next((name for name in ("Malgun Gothic", "맑은 고딕", "Noto Sans CJK KR")
                   if name in families), font.nametofont("TkDefaultFont").actual("family"))
    root.option_add("*Listbox.background", PANEL)
    root.option_add("*Listbox.foreground", INK)

    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure(".", font=(family, 10), background=PANEL, foreground=INK)
    style.configure("Surface.TFrame", background=BACKGROUND)
    style.configure("Banner.TFrame", background=NAVY)
    style.configure("Banner.TLabel", background=NAVY, foreground=PANEL, font=(family, 23, "bold"))
    style.configure("Caption.TLabel", background=NAVY, foreground="#d9e8f1")
    style.configure("Muted.TLabel", foreground=MUTED)
    style.configure("Status.TLabel", background=BACKGROUND, foreground=MUTED)
    style.configure("Error.TLabel", background="#f9ecee", foreground=DANGER)
    style.configure("Preview.TLabel", background="#eaf5f2", foreground="#125c52")
    style.configure("TButton", padding=(12, 8), bordercolor=BORDER, background=PANEL)
    style.map("TButton", background=[("pressed", "#dce7ee"), ("active", "#eaf0f5")])
    style.configure("Metric.TButton", anchor="w", justify="left", padding=(16, 12),
                    font=(family, 12, "bold"), foreground=NAVY, width=0)
    style.configure("Primary.TButton", background=TEAL, foreground=PANEL, bordercolor=TEAL)
    style.map("Primary.TButton",
              background=[("disabled", "#dfe9e9"), ("pressed", "#00594f"), ("active", "#00695f")],
              foreground=[("disabled", "#506575")])
    style.configure("TEntry", padding=7, bordercolor=BORDER, fieldbackground=PANEL)
    style.configure("TCombobox", padding=6, bordercolor=BORDER)
    style.map("TCombobox", fieldbackground=[("readonly", PANEL)])
    style.configure("TNotebook", background=BACKGROUND, borderwidth=0, tabmargins=(0, 0, 0, 0))
    style.configure("TNotebook.Tab", padding=(21, 10), background="#e1e9ef", foreground=MUTED)
    style.map("TNotebook.Tab", background=[("selected", PANEL)],
              foreground=[("selected", NAVY)])
    style.configure("Treeview", background=PANEL, fieldbackground=PANEL,
                    foreground=INK, rowheight=max(29, font.Font(family=family, size=10).metrics("linespace") + 11),
                    bordercolor=BORDER)
    style.configure("Treeview.Heading", font=(family, 9, "bold"), background="#e7eff4",
                    foreground=NAVY, padding=(6, 10), relief="flat")
    style.map("Treeview", background=[("selected", "#cce9e4")], foreground=[("selected", INK)])


def clear_table(table: ttk.Treeview) -> None:
    rows = table.get_children()
    if rows:
        table.delete(*rows)


class SortableTable(ttk.Treeview):
    def __init__(self, parent: ttk.Frame, columns: tuple[Column, ...], default_sort: tuple[int, bool]):
        self.column_specs = columns
        self.column_ids = tuple(f"column{index}" for index in range(len(columns)))
        super().__init__(parent, columns=self.column_ids, show="headings", selectmode="browse")
        self.sort_index, self.descending = default_sort
        for index, (key, column) in enumerate(zip(self.column_ids, columns)):
            self.heading(key, anchor="w" if column.anchor == "w" else "center",
                         command=lambda index=index: self.sort_by(index))
            self.column(key, width=column.width, minwidth=45, anchor=column.anchor, stretch=False)
        self.update_sort_headings()
        self.tag_configure("alternate", background="#f5f8fa")
        self.tag_configure("한도 도달", foreground="#835200")
        self.tag_configure("한도 초과", foreground=DANGER)
        self.tag_configure("계약 계획", foreground="#00776d")

    def remember_widths(self, widths: dict, on_resize: Callable[[dict], None]) -> None:
        for key, column in zip(self.column_ids, self.column_specs):
            width = widths.get(column.title)
            if type(width) is int and 45 <= width <= 20_000:
                self.column(key, width=width)

        resizing = False

        def begin(event):
            nonlocal resizing
            resizing = self.identify_region(event.x, event.y) == "separator"

        def finish(_event):
            nonlocal resizing
            if resizing:
                resizing = False
                self.after_idle(lambda: on_resize({
                    column.title: self.column(key, "width")
                    for key, column in zip(self.column_ids, self.column_specs)
                }))

        self.bind("<ButtonPress-1>", begin, add="+")
        self.bind("<ButtonRelease-1>", finish, add="+")

    def update_sort_headings(self) -> None:
        arrow = "↓" if self.descending else "↑"
        for index, (key, column) in enumerate(zip(self.column_ids, self.column_specs)):
            self.heading(key, text=f"{column.title} {arrow}" if index == self.sort_index else column.title)

    def sort_by(self, index: int) -> None:
        self.descending = index == self.sort_index and not self.descending
        self.sort_index = index
        self.update_sort_headings()
        self.apply_sort()

    def apply_sort(self) -> None:
        key = self.column_ids[self.sort_index]
        sort_key = self.column_specs[self.sort_index].sort_key
        rows = sorted(self.get_children(), key=lambda row: sort_key(str(self.set(row, key))),
                      reverse=self.descending)
        self.set_children("", *rows)
        for index, row in enumerate(rows):
            tags = tuple(tag for tag in self.item(row, "tags") if tag != "alternate")
            self.item(row, tags=(*tags, "alternate") if index % 2 else tags)


def table_in(parent: ttk.Frame, columns: tuple[Column, ...],
             default_sort: tuple[int, bool] = (0, False)) -> SortableTable:
    frame = ttk.Frame(parent)
    frame.pack(fill="both", expand=True, padx=16)
    frame.columnconfigure(0, weight=1)
    frame.rowconfigure(0, weight=1)
    table = SortableTable(frame, columns, default_sort)

    vertical = ttk.Scrollbar(frame, orient="vertical", command=table.yview)
    horizontal = ttk.Scrollbar(frame, orient="horizontal", command=table.xview)
    table.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
    table.grid(row=0, column=0, sticky="nsew")
    vertical.grid(row=0, column=1, sticky="ns")
    horizontal.grid(row=1, column=0, sticky="ew")
    return table


def fill_company_table(table: SortableTable, companies: list[CompanyYear]) -> None:
    for company in sorted(companies, key=lambda company: company.latest_contract.business_number):
        totals = company.totals
        latest = company.latest_contract
        table.insert("", "end", iid=latest.business_number, tags=(totals.status,), values=(
            latest.company, business_display(latest.business_number), f"{company.contract_count}건",
            f"{totals.count}회", won(totals.amount), f"{max(0, COUNT_LIMIT - totals.count)}회",
            won(max(0, AMOUNT_LIMIT - totals.amount)), totals.status,
        ))
    table.apply_sort()


def fill_contract_table(table: SortableTable, rows: Iterable[Contract], companies: dict[str, CompanyYear]) -> None:
    for row in sorted(rows, key=lambda row: row.date_order, reverse=True):
        totals = companies[row.business_number].totals
        tags = ("계약 계획", totals.status) if row.planned else (totals.status,)
        table.insert("", "end", iid=str(row.number), tags=tags, values=(
            row.number, row.signed_on.isoformat() if row.signed_on else "", row.contract_status,
            row.company, business_display(row.business_number), row.title, won(row.total_amount), won(row.supply_amount),
            "—" if row.vat_amount is None else won(row.vat_amount), EXEMPTION_LABELS[row.tax_exempt],
            f"{totals.count}회", won(totals.amount),
        ))
    table.apply_sort()


def reveal_modal(window: tk.Toplevel, parent: tk.Tk) -> None:
    window.update_idletasks()
    max_width, max_height = window.winfo_screenwidth() - 32, window.winfo_screenheight() - 60
    window.maxsize(max_width, max_height)
    width = min(window.winfo_reqwidth(), max_width)
    height = min(window.winfo_reqheight(), max_height)
    left = parent.winfo_rootx() + (parent.winfo_width() - width) // 2
    top = parent.winfo_rooty() + (parent.winfo_height() - height) // 2
    window.geometry(f"+{max(0, left)}+{max(0, top)}")
    window.deiconify()
    window.grab_set()


class ContractApp:
    def __init__(self, root: tk.Tk, store: WorkbookStore, ledger: Ledger):
        self.root = root
        self.store = store
        self.ledger = ledger
        self.year = tk.StringVar(value=str(date.today().year))
        self.query = tk.StringVar()
        self.status = tk.StringVar(value=f"데이터 파일  {self.store.path}")
        self.contract_hint = tk.StringVar()
        self.company_hint = tk.StringVar()
        self.metrics = tuple(tk.StringVar() for _ in METRIC_TITLES)
        self.layout_path = store.path.with_suffix(".layout.json")
        try:
            layout = json.loads(self.layout_path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            layout = {}
        self.layout = layout if isinstance(layout, dict) else {}

        configure_style(root)
        root.title(APP_TITLE)
        width = min(1320, root.winfo_screenwidth() - 60)
        height = min(760, root.winfo_screenheight() - 90)
        root.geometry(f"{width}x{height}")
        root.minsize(880, 590)
        root.configure(background=BACKGROUND)
        self.build_screen()
        self.query.trace_add("write", lambda *_: self.refresh())
        self.year.trace_add("write", lambda *_: self.refresh())
        root.bind("<Control-n>", lambda _: self.open_editor())
        self.refresh()

    def build_screen(self) -> None:
        banner = ttk.Frame(self.root, style="Banner.TFrame", padding=(30, 20))
        banner.pack(fill="x")
        ttk.Label(banner, text="계약 한도 관리", style="Banner.TLabel").pack(anchor="w")
        ttk.Label(banner, style="Caption.TLabel",
                  text=f"업체별 회계연도 한도   공급가액 {won(AMOUNT_LIMIT)}  /  {COUNT_LIMIT}회").pack(anchor="w", pady=(6, 0))

        body = ttk.Frame(self.root, style="Surface.TFrame", padding=(24, 16, 24, 9))
        body.pack(fill="both", expand=True)
        controls = ttk.Frame(body, padding=(14, 10))
        controls.pack(fill="x")
        ttk.Label(controls, text="회계연도").pack(side="left", padx=(0, 8))
        self.year_box = ttk.Combobox(controls, textvariable=self.year, state="readonly", width=8)
        self.year_box.pack(side="left", padx=(0, 22))
        ttk.Label(controls, text="검색").pack(side="left", padx=(0, 8))
        search = ttk.Entry(controls, textvariable=self.query, width=32)
        search.pack(side="left", padx=(0, 6))
        search.bind("<Escape>", lambda _: self.query.set(""))
        ttk.Button(controls, text="지우기", command=lambda: self.query.set("")).pack(side="left")
        ttk.Button(controls, text="+ 계약 등록", style="Primary.TButton",
                   command=self.open_editor).pack(side="right")

        cards = ttk.Frame(body, style="Surface.TFrame")
        cards.pack(fill="x", pady=12)
        commands = (self.show_all_contracts, self.show_company_overview,
                    lambda: self.show_company_overview(at_limit=True), self.show_plan_overview)
        for index, command in enumerate(commands):
            cards.columnconfigure(index, weight=1, uniform="metric")
            card = ttk.Button(cards, textvariable=self.metrics[index], style="Metric.TButton",
                              command=command, cursor="hand2")
            card.grid(row=0, column=index, sticky="ew", padx=(0, 8 if index < len(commands) - 1 else 0))

        self.tabs = ttk.Notebook(body)
        self.tabs.pack(fill="both", expand=True)
        self.contract_tab = ttk.Frame(self.tabs)
        company_tab = ttk.Frame(self.tabs)
        self.tabs.add(self.contract_tab, text="계약 대장")
        self.tabs.add(company_tab, text="업체별 현황")

        actions = ttk.Frame(self.contract_tab, padding=(16, 8))
        actions.pack(fill="x")
        ttk.Label(actions, textvariable=self.contract_hint, style="Muted.TLabel").pack(side="left")
        self.delete_button = ttk.Button(actions, text="삭제", command=self.delete_selected)
        self.delete_button.pack(side="right")
        self.edit_button = ttk.Button(actions, text="수정", command=self.edit_selected)
        self.edit_button.pack(side="right", padx=(0, 6))
        self.contract_table = table_in(self.contract_tab, CONTRACT_COLUMNS, default_sort=(1, True))
        self.restore_column_layout(self.contract_table, "contracts")
        self.contract_table.bind("<<TreeviewSelect>>", lambda _: self.update_buttons())
        self.contract_table.bind("<Double-1>", self.edit_selected)
        self.contract_table.bind("<Return>", self.edit_selected)

        company_actions = ttk.Frame(company_tab, padding=(16, 8))
        company_actions.pack(fill="x")
        ttk.Label(company_actions, textvariable=self.company_hint, style="Muted.TLabel").pack(side="left")
        self.detail_button = ttk.Button(company_actions, text="업체 계약 내역", command=self.show_company)
        self.detail_button.pack(side="right")
        self.company_table = table_in(company_tab, COMPANY_COLUMNS)
        self.restore_column_layout(self.company_table, "companies")
        self.company_table.bind("<<TreeviewSelect>>", lambda _: self.update_buttons())
        self.company_table.bind("<Double-1>", self.show_company)
        self.company_table.bind("<Return>", self.show_company)

        self.tabs.bind("<<NotebookTabChanged>>", lambda _: self.update_buttons())
        ttk.Label(body, textvariable=self.status, style="Status.TLabel", padding=(2, 8, 0, 0)).pack(anchor="w")

    def restore_column_layout(self, table: SortableTable, name: str) -> None:
        def save(widths):
            self.layout[name] = widths
            temporary = self.layout_path.with_suffix(".tmp")
            try:
                temporary.write_text(json.dumps(self.layout, ensure_ascii=False), encoding="utf-8")
                temporary.replace(self.layout_path)
            except OSError:
                pass

        widths = self.layout.get(name)
        table.remember_widths(widths if isinstance(widths, dict) else {}, save)

    def matching(self, company: str, business_number: str, title: str = "") -> bool:
        query = self.query.get().strip().casefold()
        digits = business_digits(query)
        if len(digits) == 10 and digits.isascii() and digits.isdecimal():
            return business_number == digits
        return (query in company.casefold() or query in title.casefold()
                or (bool(digits) and digits in business_number))

    def companies_for_year(self, year: int) -> dict[str, CompanyYear]:
        return {business_number: company for (fiscal_year, business_number), company
                in self.ledger.companies.items() if fiscal_year == year}

    def refresh(self, selected_number: int | None = None) -> None:
        year = int(self.year.get())
        years = {year, *self.ledger.fiscal_years}
        self.year_box["values"] = sorted(years, reverse=True)
        companies = self.companies_for_year(year)
        rows = tuple(contract for contract in self.ledger.contracts if contract.fiscal_year == year)
        values = (f"{len(rows):,}건", f"{len(companies):,}곳",
                  f"{sum(company.totals.status != '정상' for company in companies.values()):,}곳",
                  f"{sum(row.planned for row in rows):,}건")
        for metric, title, value in zip(self.metrics, METRIC_TITLES, values):
            metric.set(f"{title}\n{value}")

        contract_selection = self.contract_table.selection()
        if selected_number is not None:
            contract_selection = (str(selected_number),)
        company_selection = self.company_table.selection()
        clear_table(self.contract_table)
        clear_table(self.company_table)

        matched_rows = [row for row in rows if self.matching(row.company, row.business_number, row.title)]
        fill_contract_table(self.contract_table, matched_rows, companies)
        if matched_rows:
            self.contract_hint.set(f"표시 {len(matched_rows):,}건  /  전체 {len(rows):,}건")
        else:
            self.contract_hint.set("검색 조건에 맞는 계약이 없습니다." if self.query.get().strip()
                                   else "이 회계연도에 등록된 계약이 없습니다.")

        matching_numbers = {row.business_number for row in matched_rows}
        matched_companies = [
            company for number, company in companies.items()
            if number in matching_numbers or self.matching(company.latest_contract.company, number)
        ]
        fill_company_table(self.company_table, matched_companies)
        if matched_companies:
            self.company_hint.set(f"표시 {len(matched_companies):,}곳  /  전체 {len(companies):,}곳")
        else:
            self.company_hint.set("검색 조건에 맞는 업체가 없습니다." if self.query.get().strip()
                                  else "이 회계연도에 등록된 업체가 없습니다.")

        for table, selection in ((self.contract_table, contract_selection), (self.company_table, company_selection)):
            if selection and table.exists(selection[0]):
                table.selection_set(selection[0])
                table.see(selection[0])
        self.update_buttons()

    def update_buttons(self) -> None:
        contract_state = "normal" if self.contract_table.selection() else "disabled"
        self.edit_button.configure(state=contract_state)
        self.delete_button.configure(state=contract_state)
        self.detail_button.configure(state="normal" if self.company_table.selection() else "disabled")

    def show_all_contracts(self) -> None:
        self.query.set("")
        self.tabs.select(self.contract_tab)
        self.contract_table.focus_set()

    def show_company_overview(self, at_limit: bool = False) -> None:
        year = int(self.year.get())
        companies = [company for company in self.companies_for_year(year).values()
                     if not at_limit or company.totals.status != "정상"]
        title = METRIC_TITLES[2 if at_limit else 1]
        window = tk.Toplevel(self.root)
        window.withdraw()
        window.title(f"{year}년 {title}")
        window.transient(self.root)
        window.bind("<Escape>", lambda _: window.destroy())
        panel = ttk.Frame(window)
        panel.pack(fill="both", expand=True)
        header = f"{year}년 {title}  ·  {len(companies):,}곳"
        if not companies:
            header += "\n해당 업체가 없습니다."
        ttk.Label(panel, text=header, padding=(16, 12)).pack(anchor="w")
        if companies:
            table = table_in(panel, COMPANY_COLUMNS)
            table.configure(height=min(10, len(companies)))
            fill_company_table(table, companies)
        close = ttk.Button(panel, text="닫기", command=window.destroy)
        close.pack(anchor="e", padx=16, pady=12)
        reveal_modal(window, self.root)
        close.focus_set()

    def show_plan_overview(self) -> None:
        year = int(self.year.get())
        plans = [row for row in self.ledger.contracts if row.planned and row.fiscal_year == year]
        window = tk.Toplevel(self.root)
        window.withdraw()
        window.title(f"{year}년 계약 계획 현황")
        window.transient(self.root)
        window.bind("<Escape>", lambda _: window.destroy())
        panel = ttk.Frame(window)
        panel.pack(fill="both", expand=True)
        amount = sum(row.amount_for_limit for row in plans)
        count = sum(row.counted for row in plans)
        summary = (f"{year}년 계약 계획 {len(plans):,}건  ·  집계 공급가액 {won(amount)}  ·  제한 횟수 {count}회"
                   if plans else f"{year}년에 등록된 계약 계획이 없습니다.")
        ttk.Label(panel, text=summary, padding=(16, 12)).pack(anchor="w")
        if plans:
            table = table_in(panel, CONTRACT_COLUMNS, default_sort=(1, True))
            table.configure(height=min(10, len(plans)))
            fill_contract_table(table, plans, self.companies_for_year(year))
        close = ttk.Button(panel, text="닫기", command=window.destroy)
        close.pack(anchor="e", padx=16, pady=12)
        reveal_modal(window, self.root)
        close.focus_set()

    def show_company(self, event=None) -> None:
        if event and event.type == tk.EventType.ButtonPress and not self.company_table.identify_row(event.y):
            return
        selection = self.company_table.selection()
        if selection:
            self.query.set(business_display(selection[0]))
            self.tabs.select(self.contract_tab)

    def open_editor(self, contract: Contract | None = None) -> None:
        ContractEditor(self, contract)

    def edit_selected(self, event=None) -> None:
        if event and event.type == tk.EventType.ButtonPress and not self.contract_table.identify_row(event.y):
            return
        if selection := self.contract_table.selection():
            self.open_editor(self.ledger.by_number[int(selection[0])])

    def save_contract(self, contract: Contract, parent: tk.Toplevel) -> bool:
        is_new = contract.number not in self.ledger.by_number
        try:
            ledger = self.ledger.upsert(contract)
            self.store.save(ledger)
        except (ValueError, StorageError) as error:
            messagebox.showerror("저장할 수 없습니다", str(error), parent=parent)
            return False

        self.ledger = ledger
        if not self.matching(contract.company, contract.business_number, contract.title):
            self.query.set("")
        self.year.set(str(contract.fiscal_year))
        self.tabs.select(self.contract_tab)
        self.refresh(contract.number)
        self.status.set(f"{contract.number}번 계약 {'등록' if is_new else '수정'} 완료")
        return True

    def delete_selected(self) -> None:
        selection = self.contract_table.selection()
        if not selection:
            return
        number = int(selection[0])
        contract = self.ledger.by_number[number]
        if not messagebox.askyesno("계약 삭제", f"{number}번 계약을 삭제하시겠습니까?\n\n{contract.company}\n{contract.title}",
                                   parent=self.root, default=messagebox.NO):
            return

        ledger = self.ledger.delete(number)
        try:
            self.store.save(ledger)
        except StorageError as error:
            messagebox.showerror("삭제할 수 없습니다", str(error), parent=self.root)
            return
        self.ledger = ledger
        self.refresh()
        self.status.set(f"{number}번 계약 삭제 완료")


class ContractEditor:
    def __init__(self, app: ContractApp, contract: Contract | None):
        self.app = app
        self.number = contract.number if contract else app.ledger.next_number
        self.window = tk.Toplevel(app.root)
        self.window.withdraw()
        self.window.title(f"{'계약 수정' if contract else '계약 등록'} · {VERSION}")
        self.window.transient(app.root)
        self.window.resizable(False, False)

        today = date.today()
        self.fiscal_year = contract.fiscal_year if contract else int(app.year.get())
        self.planned = tk.BooleanVar(value=contract.planned if contract else False)
        default_date = today if self.fiscal_year == today.year else date(self.fiscal_year, 1, 1)
        initial_date = contract.signed_on if contract else default_date
        fields = (
            ("계약 일자", initial_date.isoformat() if initial_date else "", ""),
            ("업체명", contract.company if contract else "", ""),
            ("사업자번호", business_display(contract.business_number) if contract else "", "숫자 10자리"),
            ("계약명", contract.title if contract else "", ""),
            ("총금액", f"{contract.total_amount:,}" if contract else "", "부가세 포함"),
        )
        form = ttk.Frame(self.window, padding=(22, 16))
        form.pack(fill="both", expand=True)
        form.columnconfigure(1, weight=1)
        self.variables = []
        self.entries = []
        if not contract:
            self.add_company_history(form)
        first_row = int(contract is None)
        for row, (label, initial, hint) in enumerate(fields, start=first_row):
            ttk.Label(form, text=label).grid(row=row, column=0, sticky="w", padx=(0, 14), pady=4)
            variable = tk.StringVar(value=initial)
            entry = ttk.Entry(form, textvariable=variable, width=35)
            entry.grid(row=row, column=1, sticky="ew", pady=4)
            if row == first_row:
                ttk.Checkbutton(form, text="계약 계획", variable=self.planned, command=self.update_preview).grid(
                    row=row, column=2, sticky="w", padx=(9, 0)
                )
            elif hint:
                ttk.Label(form, text=hint, style="Muted.TLabel").grid(row=row, column=2, sticky="w", padx=(9, 0))
            self.variables.append(variable)
            self.entries.append(entry)

        supply_row = first_row + len(fields)
        self.supply_display = tk.StringVar()
        ttk.Label(form, text="공급가액").grid(row=supply_row, column=0, sticky="w", padx=(0, 14), pady=4)
        ttk.Entry(form, textvariable=self.supply_display, width=35, state="readonly", takefocus=False).grid(
            row=supply_row, column=1, sticky="ew", pady=4
        )

        vat_row = supply_row + 1
        self.legacy_tax = contract is not None and contract.tax_exempt is None
        vat_text = f"{contract.vat_amount:,}" if contract is not None and contract.vat_amount is not None else ""
        self.vat_variable = tk.StringVar(value=vat_text)
        self.tax_exempt = tk.BooleanVar(value=contract is not None and contract.tax_exempt is True)
        self.vat_before_exemption = "" if self.tax_exempt.get() else vat_text
        ttk.Label(form, text="부가세").grid(row=vat_row, column=0, sticky="w", padx=(0, 14), pady=4)
        self.vat_entry = ttk.Entry(form, textvariable=self.vat_variable, width=35)
        self.vat_entry.grid(row=vat_row, column=1, sticky="ew", pady=4)
        if self.tax_exempt.get():
            self.vat_entry.state(["readonly"])
        ttk.Label(form, text="공급가액의 10% · 직접 수정 가능", style="Muted.TLabel").grid(
            row=vat_row, column=2, sticky="w", padx=(9, 0)
        )
        self.exemption_checkbox = ttk.Checkbutton(form, text="면세", variable=self.tax_exempt,
                                                  command=self.toggle_tax_exemption)
        self.exemption_checkbox.grid(row=vat_row + 1, column=1, sticky="w", pady=(4, 5))
        ttk.Label(form, text=f"공급가액 {won(AMOUNT_THRESHOLD)} 미만 제외  ·  {won(COUNT_THRESHOLD)} 이상 1회",
                  style="Muted.TLabel").grid(row=vat_row + 2, column=0, columnspan=3, sticky="w", pady=(6, 7))
        self.feedback = ttk.Label(form, style="Preview.TLabel", padding=11, width=54, wraplength=460)
        self.feedback.grid(row=vat_row + 3, column=0, columnspan=3, sticky="ew")
        buttons = ttk.Frame(form)
        buttons.grid(row=vat_row + 4, column=0, columnspan=3, sticky="e", pady=(13, 0))
        ttk.Button(buttons, text="취소", command=self.window.destroy).pack(side="left", padx=(0, 7))
        self.save_button = ttk.Button(buttons, text="저장", command=self.submit, style="Primary.TButton")
        self.save_button.pack(side="left")
        for variable in self.variables[:-1]:
            variable.trace_add("write", self.update_preview)
        self.variables[-1].trace_add("write", self.total_amount_changed)
        self.vat_variable.trace_add("write", self.vat_changed)
        for entry in (self.entries[-1], self.vat_entry):
            entry.bind("<BackSpace>", self.skip_amount_separator)
            entry.bind("<Delete>", self.skip_amount_separator)
        self.window.bind("<Escape>", lambda _: self.window.destroy())
        self.window.bind("<Return>", self.submit)

        self.update_preview()
        reveal_modal(self.window, app.root)
        self.entries[1 if not contract else 0].focus_set()

    def add_company_history(self, form: ttk.Frame) -> None:
        latest_per_year = sorted((company.latest_contract for company in self.app.ledger.companies.values()),
                                 key=lambda row: (row.fiscal_year, row.date_order))
        latest_per_number = {row.business_number: row for row in latest_per_year}
        self.company_history = {
            f"{row.company}  ({business_display(row.business_number)})": row
            for row in sorted(latest_per_number.values(), key=lambda row: (row.company.casefold(), row.business_number))
        }
        ttk.Label(form, text="기존 업체").grid(row=0, column=0, sticky="w", padx=(0, 14), pady=(0, 8))
        self.history_box = ttk.Combobox(form, width=35, values=tuple(self.company_history),
                                        postcommand=self.filter_company_history)
        self.history_box.grid(row=0, column=1, sticky="ew", pady=(0, 8))
        hint = "업체명·번호로 검색"
        if not self.company_history:
            self.history_box.state(["disabled"])
            hint = "등록 내역 없음"
        ttk.Label(form, text=hint, style="Muted.TLabel").grid(
            row=0, column=2, sticky="w", padx=(9, 0), pady=(0, 8)
        )
        self.history_box.bind("<<ComboboxSelected>>", self.choose_company_history)
        self.history_box.bind("<Return>", self.choose_company_history)

    def history_matches(self) -> tuple[str, ...]:
        query = self.history_box.get().strip().casefold()
        digits = business_digits(query)
        return tuple(label for label, row in self.company_history.items()
                     if query in row.company.casefold() or (digits and digits in row.business_number))

    def filter_company_history(self) -> None:
        if self.history_box.get() in self.company_history:
            self.history_box["values"] = tuple(self.company_history)
        else:
            self.history_box["values"] = self.history_matches()

    def choose_company_history(self, _event) -> str:
        label = self.history_box.get()
        if label not in self.company_history:
            matches = self.history_matches()
            if len(matches) != 1:
                return "break"
            label = matches[0]
            self.history_box.set(label)
        company = self.company_history[label]
        self.variables[1].set(company.company)
        self.variables[2].set(business_display(company.business_number))
        self.entries[3].focus_set()
        return "break"

    @staticmethod
    def format_amount(variable: tk.StringVar, entry: ttk.Entry) -> None:
        text = variable.get()
        digits = text.strip().replace(",", "")
        if not digits or (digits.isascii() and digits.isdecimal()):
            try:
                formatted = f"{int(digits):,}" if digits else ""
            except ValueError:
                formatted = text
            if text != formatted:
                caret = entry.index(tk.INSERT) if entry.get() == text else len(text)
                digits_before_caret = len(text[:caret].strip().replace(",", ""))
                leading_zeros_removed = len(digits) - len(formatted.replace(",", ""))
                remaining = max(0, digits_before_caret - leading_zeros_removed)
                position = 0
                while remaining:
                    remaining -= formatted[position] != ","
                    position += 1
                # Update the entry before restoring its caret so a newly added comma counts in its length.
                entry.delete(0, tk.END)
                entry.insert(0, formatted)
                variable.set(formatted)
                entry.icursor(position)

    @staticmethod
    def skip_amount_separator(event) -> None:
        entry = event.widget
        if entry.selection_present():
            return
        caret = entry.index(tk.INSERT)
        text = entry.get()
        if event.keysym == "BackSpace" and caret > 0 and text[caret - 1] == ",":
            entry.icursor(caret - 1)
        elif event.keysym == "Delete" and text[caret:caret + 1] == ",":
            entry.icursor(caret + 1)

    def suggested_vat(self) -> str:
        try:
            amount = parse_won(self.variables[-1].get(), "총금액")
        except ValueError:
            return ""
        # One eleventh of the VAT-inclusive amount is tax. Discard sub-won fractions.
        return f"{amount // 11:,}"

    def total_amount_changed(self, *_args) -> None:
        self.format_amount(self.variables[-1], self.entries[-1])
        suggested_vat = self.suggested_vat()
        if self.tax_exempt.get():
            self.vat_before_exemption = suggested_vat
            self.update_preview()
        else:
            self.vat_variable.set(suggested_vat)

    def vat_changed(self, *_args) -> None:
        if not self.tax_exempt.get():
            self.format_amount(self.vat_variable, self.vat_entry)
        self.update_preview()

    def toggle_tax_exemption(self) -> None:
        if self.tax_exempt.get():
            self.vat_before_exemption = self.vat_variable.get()
            self.vat_entry.state(["readonly"])
            self.vat_variable.set("0")
        else:
            self.vat_entry.state(["!readonly"])
            self.vat_variable.set(self.vat_before_exemption or self.suggested_vat())

    def split_amount(self) -> tuple[int, int | None]:
        total = parse_won(self.variables[-1].get(), "총금액")
        tax_text = self.vat_variable.get()
        if self.legacy_tax and not self.tax_exempt.get() and not tax_text:
            return total, None
        vat = parse_won(tax_text, "부가세")
        if vat > total:
            raise ValueError("부가세는 총금액을 초과할 수 없습니다.")
        return total - vat, vat

    def candidate(self, supply: int, vat: int | None) -> Contract:
        tax_exempt = self.tax_exempt.get()
        return Contract.from_fields(self.number, *(value.get() for value in self.variables[:-1]),
                                    supply_amount=supply, vat_amount=vat,
                                    tax_exempt=None if vat is None else tax_exempt,
                                    planned=self.planned.get(), fiscal_year=self.fiscal_year)

    def update_preview(self, *_args) -> None:
        self.supply_display.set("")
        try:
            supply, vat = self.split_amount()
            self.supply_display.set(f"{supply:,}")
            candidate = self.candidate(supply, vat)
        except ValueError as error:
            self.feedback.configure(text=str(error), style="Muted.TLabel")
            self.save_button.state(["disabled"])
            return

        before, after = self.app.ledger.preview(candidate)
        line = f"{candidate.fiscal_year}년 누적   {before.count}회 / {won(before.amount)} → {after.count}회 / {won(after.amount)}"
        violations = self.app.ledger.limit_violations(candidate, before, after)
        status = "집계 제외" if not candidate.amount_for_limit else after.status
        self.feedback.configure(text=f"{line}\n{' · '.join(violations) if violations else status}",
                                style="Error.TLabel" if violations else "Preview.TLabel")
        self.save_button.state(["disabled"] if violations else ["!disabled"])

    def submit(self, _event=None) -> None:
        if self.save_button.instate(["disabled"]):
            return
        if self.app.save_contract(self.candidate(*self.split_amount()), self.window):
            self.window.destroy()


def main() -> int:
    data_directory = Path(sys.executable if getattr(sys, "frozen", False) else __file__).resolve().parent
    store = WorkbookStore(data_directory / "계약대장.xlsx")
    root = tk.Tk()
    root.withdraw()
    try:
        with store.lock():
            ledger = store.load()
            ContractApp(root, store, ledger)
            root.deiconify()
            root.mainloop()
    except StorageError as error:
        messagebox.showerror(APP_TITLE, str(error), parent=root)
        root.destroy()
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
