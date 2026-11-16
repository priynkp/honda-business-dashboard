"""Honda Business Management Dashboard: a local, Excel-backed desktop app."""
from __future__ import annotations

import os
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from types import SimpleNamespace
from zipfile import BadZipFile
import tkinter as tk
from tkinter import messagebox, ttk

import pandas as pd
from analysis import analyze, export_analysis
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

BASE_DIR = Path(__file__).resolve().parent
# Sample workbooks are read-only seeds. Local additions live in ignored data/.
SCHEMAS = {
    "Rentals": ("Customer Name", "Car Model", "Cost", "Rental Date"),
    "Repairs": ("DATE", "VEHICLE NO", "CUSTOMER NAME", "CAR MODEL", "ISSUES", "COST"),
    "Sales": ("CUSTOMER_NAME", "CAR_MODEL", "SALE_DATE", "SALE_PRICE"),
}
AMOUNT_COLUMNS = {"Rentals": "Cost", "Repairs": "COST", "Sales": "SALE_PRICE"}
DATE_COLUMNS = {"DATE", "SALE_DATE", "Rental Date"}

RED = '#c80f0f'
RED_DARK = '#9c0b0b'
INK = '#1f2937'
MUTED = '#6b7280'
PAPER = '#ffffff'
SOFT = '#f7f7f8'
BORDER = '#e5e7eb'
STRIPE = '#f4f4f5'

USERS = {'admin': '1234'}


def normalize_record(category: str, values: list) -> list:
    """Validate a row without changing the caller's data."""
    headers = SCHEMAS[category]
    if len(values) != len(headers):
        raise ValueError(f"{category}: expected {len(headers)} fields.")
    result = []
    for header, raw in zip(headers, values):
        if pd.isna(raw) or not str(raw).strip():
            raise ValueError(f"{header.replace('_', ' ').title()} is required.")
        text = str(raw).strip()
        if header == AMOUNT_COLUMNS[category]:
            try:
                amount = Decimal(text)
            except InvalidOperation as exc:
                raise ValueError("Amount must be a number, without commas or currency symbols.") from exc
            if not amount.is_finite() or amount < 0 or amount > Decimal('1000000000000'):
                raise ValueError("Amount must be between 0 and 1,000,000,000,000.")
            if amount != amount.quantize(Decimal('0.01')):
                raise ValueError("Amount can have at most two decimal places.")
            result.append(float(amount))
        elif header in DATE_COLUMNS:
            if isinstance(raw, datetime):
                text = raw.strftime('%Y-%m-%d')
            try:
                parsed = datetime.strptime(text, '%Y-%m-%d')
            except ValueError as exc:
                raise ValueError("Use a valid date in YYYY-MM-DD format.") from exc
            result.append(parsed.strftime('%Y-%m-%d'))
        else:
            # Avoid storing Excel formulas from form inputs.
            if text.startswith(('=', '+', '-', '@')):
                raise ValueError("Text fields cannot start with =, +, - or @.")
            result.append(text)
    return result


class WorkbookStore:
    """Load sample workbooks, then atomically save local working copies per category."""

    def __init__(self, base_dir: Path = BASE_DIR):
        self.base_dir = Path(base_dir)
        self.local_dir = self.base_dir / 'data'
        self.frames = {}
        for category, headers in SCHEMAS.items():
            local = self.local_dir / f'{category}.xlsx'
            path = local if local.exists() else self.base_dir / f'{category}.xlsx'
            frame = pd.read_excel(path, engine='openpyxl')
            if tuple(frame.columns) != headers:
                raise ValueError(f'{path.name}: expected columns {", ".join(headers)}.')
            rows = []
            for number, row in enumerate(frame.itertuples(index=False, name=None), 2):
                try:
                    rows.append(normalize_record(category, list(row)))
                except ValueError as exc:
                    raise ValueError(f'{path.name}, row {number}: {exc}') from exc
            self.frames[category] = pd.DataFrame(rows, columns=headers)

    def filtered(self, category: str, query: str = '') -> pd.DataFrame:
        frame = self.frames[category]
        query = query.strip().casefold()
        if not query:
            return frame.copy()
        mask = frame.astype(str).apply(
            lambda column: column.str.casefold().str.contains(query, regex=False)
        ).any(axis=1)
        return frame.loc[mask].copy()

    def add(self, category: str, values: list) -> None:
        record = normalize_record(category, values)
        candidate = pd.concat([
            self.frames[category], pd.DataFrame([record], columns=SCHEMAS[category])
        ], ignore_index=True)
        self.local_dir.mkdir(exist_ok=True)
        temporary = self.local_dir / f'.{category}.tmp.xlsx'
        try:
            candidate.to_excel(temporary, index=False, engine='openpyxl')
            os.replace(temporary, self.local_dir / f'{category}.xlsx')
        finally:
            temporary.unlink(missing_ok=True)
        # Only publish the change to the UI after a successful save.
        self.frames[category] = candidate


class HondaApp(tk.Tk):
    """Login screen, then a red Honda-themed dashboard with tables and analytics."""

    def __init__(self, store: WorkbookStore):
        super().__init__()
        self.store = store
        self.title('Honda | Business Management Dashboard')
        self.configure(bg=PAPER)
        self._style()
        self.query = tk.StringVar()
        self.status = tk.StringVar(value='Ready')
        self.year = tk.StringVar(value='All years')
        self.insight_tables = {}
        self.tables = {}
        self.row_labels = {}
        self.entries = {}
        self.metric_vars = {}
        self.query.trace_add('write', lambda *_: self.refresh_tables())
        self.show_login()

    # ---------- shared styling ----------
    def _style(self):
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure('.', font=('Arial', 10), background=PAPER, foreground=INK)
        style.configure('TFrame', background=PAPER)
        style.configure('TLabel', background=PAPER)
        style.configure('Muted.TLabel', foreground=MUTED)
        style.configure('TNotebook', background=PAPER, borderwidth=0)
        style.configure('TNotebook.Tab', padding=(24, 10), font=('Arial', 10, 'bold'))
        style.map('TNotebook.Tab',
                  background=[('selected', PAPER)],
                  foreground=[('selected', RED), ('!selected', MUTED)])
        style.configure('Treeview', rowheight=28, background=PAPER,
                        fieldbackground=PAPER, font=('Arial', 10))
        style.configure('Treeview.Heading', background=RED, foreground='white',
                        font=('Arial', 10, 'bold'), padding=8, relief='flat')
        style.map('Treeview.Heading', background=[('active', RED_DARK)])
        style.configure('TLabelframe', background=PAPER, bordercolor=BORDER)
        style.configure('TLabelframe.Label', background=PAPER, foreground=INK,
                        font=('Arial', 10, 'bold'))
        style.configure('TEntry', padding=5)
        style.configure('TCombobox', padding=4)

    def _red_button(self, parent, text, command):
        return tk.Button(parent, text=text, command=command, bg=RED, fg='white',
                         activebackground=RED_DARK, activeforeground='white',
                         relief='flat', bd=0, padx=18, pady=8, cursor='hand2',
                         font=('Arial', 10, 'bold'))

    def _clear(self):
        for child in self.winfo_children():
            child.destroy()

    # ---------- login view ----------
    def show_login(self):
        self._clear()
        self.geometry('440x460')
        self.minsize(420, 440)
        band = tk.Frame(self, bg=RED, height=110)
        band.pack(fill='x')
        band.pack_propagate(False)
        tk.Label(band, text='HONDA', bg=RED, fg='white',
                 font=('Arial', 26, 'bold')).pack(pady=(22, 0))
        tk.Label(band, text='BUSINESS MANAGEMENT DASHBOARD', bg=RED, fg='#ffe1e1',
                 font=('Arial', 9, 'bold')).pack()
        card = tk.Frame(self, bg=PAPER)
        card.pack(fill='both', expand=True, padx=48, pady=28)
        tk.Label(card, text='Welcome back', bg=PAPER, fg=INK,
                 font=('Arial', 16, 'bold')).pack(anchor='w')
        tk.Label(card, text='Sign in to continue', bg=PAPER, fg=MUTED,
                 font=('Arial', 10)).pack(anchor='w', pady=(2, 18))
        tk.Label(card, text='Username', bg=PAPER, fg=INK,
                 font=('Arial', 10, 'bold')).pack(anchor='w')
        username = ttk.Entry(card, width=30, font=('Arial', 11))
        username.pack(fill='x', pady=(4, 12))
        tk.Label(card, text='Password', bg=PAPER, fg=INK,
                 font=('Arial', 10, 'bold')).pack(anchor='w')
        password = ttk.Entry(card, show='*', width=30, font=('Arial', 11))
        password.pack(fill='x', pady=(4, 16))

        def verify(_event=None):
            if USERS.get(username.get().strip()) == password.get().strip():
                self.show_dashboard()
            else:
                messagebox.showerror('Login Failed', 'Incorrect username or password', parent=self)

        self._red_button(card, 'Login', verify).pack(fill='x')
        tk.Label(card, text='Use admin / 1234 to sign in', bg=PAPER, fg=MUTED,
                 font=('Arial', 9)).pack(pady=(14, 0))
        password.bind('<Return>', verify)
        username.focus_set()

    # ---------- dashboard view ----------
    def show_dashboard(self):
        self._clear()
        self.geometry('1280x860')
        self.minsize(1080, 760)
        header = tk.Frame(self, bg=RED, height=72)
        header.pack(fill='x')
        header.pack_propagate(False)
        inner = tk.Frame(header, bg=RED)
        inner.pack(fill='both', expand=True, padx=24)
        brand = tk.Frame(inner, bg=RED)
        brand.pack(side='left', pady=12)
        tk.Label(brand, text='HONDA', bg=RED, fg='white',
                 font=('Arial', 20, 'bold')).pack(anchor='w')
        tk.Label(brand, text='Business Management Dashboard', bg=RED, fg='#ffe1e1',
                 font=('Arial', 10)).pack(anchor='w')
        tk.Button(inner, text='Logout', command=self.show_login, bg=PAPER, fg=RED,
                  activebackground='#f3e8e8', activeforeground=RED_DARK,
                  relief='flat', bd=0, padx=16, pady=6, cursor='hand2',
                  font=('Arial', 10, 'bold')).pack(side='right')

        search = tk.Frame(self, bg=PAPER)
        search.pack(fill='x', padx=24, pady=(16, 6))
        tk.Label(search, text='Search records', bg=PAPER, fg=INK,
                 font=('Arial', 10, 'bold')).pack(side='left', padx=(0, 12))
        ttk.Entry(search, textvariable=self.query, font=('Arial', 11)).pack(
            side='left', fill='x', expand=True)
        self._red_button(search, 'Clear', lambda: self.query.set('')).pack(side='left', padx=(10, 0))
        tk.Label(search, text='Searches every field in all record tabs', bg=PAPER,
                 fg=MUTED, font=('Arial', 9)).pack(side='left', padx=(14, 0))

        footer = tk.Frame(self, bg=PAPER)
        footer.pack(side='bottom', fill='x', padx=24, pady=(4, 10))
        tk.Label(footer, textvariable=self.status, bg=PAPER, fg=MUTED,
                 font=('Arial', 9)).pack(side='left')
        tk.Label(footer, text='Honda Business Management System © 2025',
                 bg=PAPER, fg=MUTED, font=('Arial', 9)).pack(side='right')

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill='both', expand=True, padx=24, pady=(6, 6))
        for category in SCHEMAS:
            self._record_tab(category)
        self._analytics_tab()
        self.refresh_tables()
        self.refresh_analytics()

    def _record_tab(self, category):
        tab = ttk.Frame(self.notebook, padding=18)
        self.notebook.add(tab, text=category)
        form = ttk.LabelFrame(tab, text='Add a record', padding=14)
        form.pack(fill='x', pady=(0, 14))
        self.entries[category] = []
        for index, header in enumerate(SCHEMAS[category]):
            row, col = divmod(index, 3)
            cell = ttk.Frame(form)
            cell.grid(row=row, column=col, sticky='ew', padx=(0, 14), pady=(0, 10))
            form.columnconfigure(col, weight=1)
            label = header.replace('_', ' ').title()
            if header in DATE_COLUMNS:
                label += ' (YYYY-MM-DD)'
            ttk.Label(cell, text=label).pack(anchor='w', pady=(0, 4))
            field = ttk.Entry(cell)
            field.pack(fill='x')
            self.entries[category].append(field)
        self._red_button(form, 'Add & save locally',
                         lambda: self.add_record(category)).grid(
            row=2, column=0, sticky='w', pady=(4, 0))
        ttk.Label(form, text='All fields required. Amounts: numbers only.',
                  style='Muted.TLabel').grid(row=2, column=1, columnspan=2,
                                             sticky='w', pady=(4, 0))
        count = tk.StringVar()
        self.row_labels[category] = count
        ttk.Label(tab, textvariable=count, style='Muted.TLabel').pack(anchor='w', pady=(0, 8))
        grid = ttk.Frame(tab)
        grid.pack(fill='both', expand=True)
        grid.rowconfigure(0, weight=1)
        grid.columnconfigure(0, weight=1)
        table = ttk.Treeview(grid, columns=SCHEMAS[category], show='headings',
                             selectmode='browse')
        for header in SCHEMAS[category]:
            table.heading(header, text=header.replace('_', ' ').title())
            table.column(header, width=270 if header == 'ISSUES' else 165,
                         minwidth=120, stretch=True)
        table.grid(row=0, column=0, sticky='nsew')
        vertical = ttk.Scrollbar(grid, orient='vertical', command=table.yview)
        vertical.grid(row=0, column=1, sticky='ns')
        horizontal = ttk.Scrollbar(grid, orient='horizontal', command=table.xview)
        horizontal.grid(row=1, column=0, sticky='ew')
        table.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        table.tag_configure('stripe', background=STRIPE)
        self.tables[category] = table

    def _analytics_tab(self):
        tab = ttk.Frame(self.notebook, padding=18)
        self.notebook.add(tab, text='Analytics')
        self.analytics_tab = tab

        controls = ttk.Frame(tab)
        controls.pack(fill='x', pady=(0, 12))
        ttk.Label(controls, text='Analysis period',
                  font=('Arial', 10, 'bold')).pack(side='left', padx=(0, 10))
        years = sorted({str(pd.Timestamp(value).year)
                        for category, frame in self.store.frames.items()
                        for column in SCHEMAS[category] if column in DATE_COLUMNS
                        for value in frame[column]})
        self.year_picker = ttk.Combobox(controls, textvariable=self.year,
                                        values=['All years', *years],
                                        state='readonly', width=12)
        self.year_picker.pack(side='left')
        self.year_picker.bind('<<ComboboxSelected>>', lambda _: self.refresh_analytics())
        self._red_button(controls, 'Export analysis CSVs',
                         self.export_reports).pack(side='right')

        cards = tk.Frame(tab, bg=PAPER)
        cards.pack(fill='x', pady=(0, 14))
        for col, category in enumerate(SCHEMAS):
            card = tk.Frame(cards, bg=PAPER, highlightbackground=BORDER,
                            highlightthickness=1)
            card.grid(row=0, column=col, sticky='ew', padx=(0, 14 if col < 2 else 0))
            cards.columnconfigure(col, weight=1)
            tk.Label(card, text=category.upper(), bg=PAPER, fg=MUTED,
                     font=('Arial', 9, 'bold')).pack(anchor='w', padx=14, pady=(10, 2))
            var = tk.StringVar()
            self.metric_vars[category] = var
            tk.Label(card, textvariable=var, bg=PAPER, fg=RED,
                     font=('Arial', 13, 'bold'), justify='left').pack(
                anchor='w', padx=14, pady=(0, 10))
        avg_card = tk.Frame(cards, bg=PAPER, highlightbackground=BORDER,
                            highlightthickness=1)
        avg_card.grid(row=0, column=3, sticky='ew')
        cards.columnconfigure(3, weight=1)
        tk.Label(avg_card, text='AVG REPAIR COST', bg=PAPER, fg=MUTED,
                 font=('Arial', 9, 'bold')).pack(anchor='w', padx=14, pady=(10, 2))
        self.avg_var = tk.StringVar()
        tk.Label(avg_card, textvariable=self.avg_var, bg=PAPER, fg=RED,
                 font=('Arial', 13, 'bold')).pack(anchor='w', padx=14, pady=(0, 10))

        paned = ttk.PanedWindow(tab, orient='horizontal')
        paned.pack(fill='both', expand=True)
        charts = ttk.Frame(paned, padding=(0, 0, 8, 0))
        paned.add(charts, weight=3)
        self.figure = Figure(figsize=(7.2, 5.2), dpi=100, layout='constrained',
                             facecolor=PAPER)
        self.canvas = FigureCanvasTkAgg(self.figure, master=charts)
        widget = self.canvas.get_tk_widget()
        widget.bind('<Configure>', lambda _: self.after_idle(self.resize_chart), add='+')
        widget.bind('<Map>', lambda _: self.after_idle(self.resize_chart), add='+')
        widget.pack(fill='both', expand=True)

        side = ttk.Frame(paned, padding=(8, 0, 0, 0))
        paned.add(side, weight=2)
        reports = ttk.Notebook(side)
        reports.pack(fill='both', expand=True)
        for key, title, headers in [
            ('model_sales', 'Sales by model', ('Model', 'Sales', 'Amount (INR)')),
            ('top_customers', 'Top customers', ('Customer', 'Transactions', 'Amount (INR)')),
            ('repair_summary', 'Repair breakdown', ('Issue', 'Jobs', 'Amount (INR)', 'Avg cost (INR)')),
        ]:
            page = ttk.Frame(reports, padding=10)
            reports.add(page, text=title)
            table = ttk.Treeview(page, columns=headers, show='headings')
            for header in headers:
                table.heading(header, text=header)
                table.column(header, width=140, minwidth=90)
            scrollbar = ttk.Scrollbar(page, orient='vertical', command=table.yview)
            scrollbar.pack(side='right', fill='y')
            table.configure(yscrollcommand=scrollbar.set)
            table.pack(fill='both', expand=True)
            self.insight_tables[key] = table

    # ---------- data refresh ----------
    def refresh_tables(self):
        if not self.tables:
            return
        for category, table in self.tables.items():
            items = table.get_children()
            if items:
                table.delete(*items)
            frame = self.store.filtered(category, self.query.get())
            for index, row in enumerate(frame.itertuples(index=False, name=None)):
                values = [f'{value:,.2f}' if header == AMOUNT_COLUMNS[category] else value
                          for header, value in zip(SCHEMAS[category], row)]
                table.insert('', 'end', values=values, tags=('stripe',) if index % 2 else ())
            total = len(self.store.frames[category])
            self.row_labels[category].set(
                f'{len(frame)} of {total} records shown' + (' | No matches' if frame.empty else ''))

    def add_record(self, category):
        try:
            self.store.add(category, [entry.get() for entry in self.entries[category]])
        except (ValueError, OSError, ImportError) as exc:
            messagebox.showerror('Could not save record', str(exc), parent=self)
            return
        for entry in self.entries[category]:
            entry.delete(0, tk.END)
        self.refresh_tables()
        years = sorted({str(pd.Timestamp(value).year)
                        for key, frame in self.store.frames.items()
                        for column in SCHEMAS[key] if column in DATE_COLUMNS
                        for value in frame[column]})
        self.year_picker.configure(values=['All years', *years])
        self.refresh_analytics()
        self.status.set(f'Saved to data/{category}.xlsx | Clear search if the new record is hidden.')

    def resize_chart(self):
        # Keep the figure inside the Tk canvas on displays with non-default DPI.
        widget = self.canvas.get_tk_widget()
        self.canvas.resize(SimpleNamespace(width=widget.winfo_width(),
                                           height=widget.winfo_height()))

    def refresh_analytics(self):
        if not self.metric_vars:
            return
        year = None if self.year.get() == 'All years' else int(self.year.get())
        self.results = analyze(self.store.frames, year)
        kpis = self.results['kpis']
        for category in SCHEMAS:
            row = kpis.loc[category]
            self.metric_vars[category].set(
                f"{int(row['records']):,} records\nINR {row['recorded_amount']:,.0f}")
        self.avg_var.set(f"INR {kpis.loc['Repairs', 'average_amount']:,.2f}")
        for key, table in self.insight_tables.items():
            items = table.get_children()
            if items:
                table.delete(*items)
            for index, row in self.results[key].iterrows():
                values = [index] + [f'{value:,.2f}' if 'amount' in column or 'cost' in column
                                    else int(value) for column, value in row.items()]
                table.insert('', 'end', values=values)
        self._draw_charts(kpis)

    def _draw_charts(self, kpis):
        self.figure.clear()
        axes = self.figure.subplots(2, 2)
        monthly = self.results['monthly']
        positions = list(range(len(monthly)))
        labels = monthly.index.strftime('%b %y')
        axes[0, 0].plot(positions, monthly['Sales'] / 100000, color=RED, linewidth=2)
        axes[0, 0].set_title('Monthly sales value', loc='left', fontsize=10)
        axes[0, 0].set_ylabel('INR lakh (100,000)', fontsize=8)
        for category, color in [('Rentals', '#1f2937'), ('Repairs', '#9ca3af')]:
            axes[0, 1].plot(positions, monthly[category] / 1000, label=category,
                            color=color, linewidth=1.8)
        axes[0, 1].set_title('Monthly service amounts', loc='left', fontsize=10)
        axes[0, 1].set_ylabel('INR thousand', fontsize=8)
        axes[0, 1].legend(fontsize=7, frameon=False)
        ticks = positions[::max(1, len(positions) // 6)] if positions else []
        for ax in axes[0]:
            ax.set_xticks(ticks, [labels[i] for i in ticks], fontsize=7)
        model = self.results['model_sales'].sort_values('recorded_amount')
        axes[1, 0].barh(model.index, model['recorded_amount'] / 100000,
                        color=RED, height=.6)
        axes[1, 0].set_title('Sales value by car model', loc='left', fontsize=10)
        axes[1, 0].set_xlabel('INR lakh (100,000)', fontsize=8)
        counts = kpis['records']
        bars = axes[1, 1].bar(counts.index, counts,
                              color=['#1f2937', '#9ca3af', RED], width=.5)
        axes[1, 1].set_title('Transaction mix', loc='left', fontsize=10)
        axes[1, 1].set_ylabel('Record count', fontsize=8)
        axes[1, 1].bar_label(bars, labels=[f'{value:,.0f}' for value in counts],
                             padding=3, fontsize=8)
        axes[1, 1].margins(y=.2)
        for ax in axes.flat:
            ax.spines[['top', 'right']].set_visible(False)
            ax.set_axisbelow(True)
            ax.grid(axis='y', alpha=.15)
            ax.tick_params(labelsize=8)
        self.canvas.draw_idle()

    def export_reports(self):
        period = self.year.get().lower().replace(' ', '-')
        directory = self.store.base_dir / 'exports' / period
        try:
            paths = export_analysis(self.results, directory)
        except OSError as exc:
            messagebox.showerror('Could not export', str(exc), parent=self)
            return
        self.status.set(f'Exported {len(paths)} analysis CSVs to {directory}')


def main():
    try:
        store = WorkbookStore()
    except (OSError, ValueError, ImportError, BadZipFile) as exc:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            'Cannot load workbooks',
            f'{exc}\n\nKeep the three sample .xlsx files beside app.py.\n'
            'If a local file is invalid, back it up before repairing it.',
            parent=root)
        root.destroy()
        return
    HondaApp(store).mainloop()


if __name__ == '__main__':
    main()
