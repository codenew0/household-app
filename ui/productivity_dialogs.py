"""明細一覧・入力支援・予算・復元の画面。"""
import copy
import datetime
import tkinter as tk
from tkinter import ttk, messagebox
from ui.base_dialog import BaseDialog
from ui.form_widgets import scrollable_form, polish_form
from ui.month_picker import MonthPicker
from models.transactions import normalize, cash_amount, is_pending


def refreshed(app):
    app._show_month(app.current_month)


class LedgerListDialog(BaseDialog):
    def __init__(self, parent, app, pending=False):
        super().__init__(parent, '未確定の一括確認' if pending else '明細一覧', 1100, 700)
        self.app = app
        self.manager = app.data_manager
        self.body, self.footer = scrollable_form(self)
        self.references = {}
        self.sort_column = 'date'
        self.sort_reverse = False
        filters = ttk.Frame(self.body)
        filters.pack(fill='x', padx=10, pady=10)
        self.start = ttk.Entry(filters, width=11)
        self.end = ttk.Entry(filters, width=11)
        self.start.insert(0, '1900-01' if pending else f'{app.current_year}-{app.current_month:02d}')
        self.end.insert(0, '9999-12' if pending else f'{app.current_year}-{app.current_month:02d}')
        for field in (self.start, self.end):
            field.configure(state='readonly', cursor='hand2')
        for label, widget in (('開始年月', self.start), ('終了年月', self.end)):
            ttk.Label(filters, text=label).pack(side='left'); widget.pack(side='left', padx=5)
        self.start_picker = MonthPicker(filters, app.current_year, app.current_month, lambda y, m: self.pick_period(self.start, y, m), anchor=self.start)
        self.end_picker = MonthPicker(filters, app.current_year, app.current_month, lambda y, m: self.pick_period(self.end, y, m), anchor=self.end)
        self.start.bind('<Button-1>', lambda e: self.open_period(self.start, self.start_picker))
        self.end.bind('<Button-1>', lambda e: self.open_period(self.end, self.end_picker))
        self.query = ttk.Entry(filters, width=20)
        self.query.pack(side='left', padx=5)
        self.only_pending = tk.BooleanVar(value=pending)
        ttk.Checkbutton(filters, text='未確定のみ', variable=self.only_pending).pack(side='left')
        ttk.Button(filters, text='検索・更新', command=self.refresh).pack(side='left', padx=5)
        table = ttk.Frame(self.body)
        table.pack(fill='both', expand=True, padx=10)
        table.rowconfigure(0, weight=1)
        table.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(table, columns=('date', 'category', 'partner', 'amount', 'points', 'cash', 'method', 'memo', 'status'), show='headings', selectmode='extended', height=17)
        self.column_titles = {}
        for name, title in zip(self.tree['columns'], ('日付', '分類', '支払先', '利用前金額', 'ポイント', '現金', '支払方法', 'メモ', '状態')):
            self.column_titles[name] = title
            self.tree.heading(name, text=title, command=lambda column=name: self.sort_by(column))
            self.tree.column(name, width=105, minwidth=65, stretch=True)
        self.tree.tag_configure('pending', foreground='red')
        self.tree.tag_configure('duplicate', background='#ffcccc')
        self.tree.grid(row=0, column=0, sticky='nsew')
        vertical = ttk.Scrollbar(table, orient='vertical', command=self.tree.yview)
        vertical.grid(row=0, column=1, sticky='ns')
        horizontal = ttk.Scrollbar(table, orient='horizontal', command=self.tree.xview)
        horizontal.grid(row=1, column=0, sticky='ew')
        self.tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.tree.bind('<<TreeviewSelect>>', lambda e: self.total())
        self.tree.bind('<Control-a>', self.select_all)
        self.tree.bind('<Double-1>', lambda e: self.edit())
        self.tree.bind('<Button-3>', self.context_menu)
        self.bind('<Escape>', lambda e: self.destroy())
        self.summary = ttk.Label(self.body)
        self.summary.pack(fill='x', padx=10, pady=10)
        for title, command in (('選択を確定', self.confirm), ('選択を削除', self.delete), ('閉じる', self.destroy)):
            ttk.Button(self.footer, text=title, command=command).pack(side='left', padx=4)
        polish_form(self)
        self.refresh()
        self.show_ready()

    def open_period(self, field, picker):
        try:
            date = datetime.datetime.strptime(field.get(), '%Y-%m')
            picker.set(date.year, date.month)
        except ValueError:
            pass
        picker.open()
        return 'break'

    def pick_period(self, field, year, month):
        field.configure(state='normal')
        field.delete(0, 'end')
        field.insert(0, f'{year}-{month:02d}')
        field.configure(state='readonly')
        self.refresh()

    def refresh(self):
        try:
            start = datetime.datetime.strptime(self.start.get(), '%Y-%m').date()
            end = datetime.datetime.strptime(self.end.get(), '%Y-%m').date()
            if start > end: raise ValueError()
        except ValueError:
            messagebox.showwarning('期間', '開始・終了をYYYY-MM形式で正しく入力してください。', parent=self); return
        self.tree.delete(*self.tree.get_children()); self.references.clear()
        columns = self.app.get_all_columns()
        for key in sorted(self.manager.data, key=lambda k: self.manager._parse_key(k) or (0, 0, 0, 0)):
            parsed = self.manager._parse_key(key)
            if not parsed or not (start.year, start.month) <= parsed[:2] <= (end.year, end.month): continue
            year, month, day, column = parsed
            for index, row in enumerate(self.manager.data[key]):
                value = normalize(row)
                if self.only_pending.get() and not is_pending(value): continue
                category = '収入' if day == 0 else columns[column] if column < len(columns) else str(column)
                if self.query.get().casefold() not in (' '.join(value) + category).casefold(): continue
                iid = str(len(self.references))
                self.references[iid] = (key, index, copy.deepcopy(row))
                self.tree.insert('', 'end', iid=iid, values=(f'{year}-{month:02d}-{day:02d}', category, value[0], value[1], value[3], cash_amount(value), value[4], value[2], value[5]), tags=('pending',) if is_pending(value) else ())
        self.highlight_duplicates()
        self.apply_sort()
        self.total()

    def sort_by(self, column):
        self.sort_reverse = not self.sort_reverse if self.sort_column == column else False
        self.sort_column = column
        self.apply_sort()

    def apply_sort(self):
        column = self.sort_column
        def key(iid):
            if column == 'date':
                return self.manager._parse_key(self.references[iid][0])[:3]
            value = self.tree.set(iid, column)
            return int(value.replace(',', '') or '0') if column in ('amount', 'points', 'cash') else value.casefold()
        for index, iid in enumerate(sorted(self.tree.get_children(), key=key, reverse=self.sort_reverse)):
            self.tree.move(iid, '', index)
        for name, title in self.column_titles.items():
            self.tree.heading(name, text=title + (' ▼' if self.sort_reverse else ' ▲') if name == column else title)

    def highlight_duplicates(self):
        groups = {}
        for iid in self.tree.get_children():
            signature = tuple(str(value) for value in self.tree.item(iid, 'values'))
            groups.setdefault(signature, []).append(iid)
        for members in groups.values():
            if len(members) > 1:
                for iid in members:
                    tags = tuple(self.tree.item(iid, 'tags'))
                    self.tree.item(iid, tags=tags + ('duplicate',))

    def context_menu(self, event):
        old_menu = getattr(self, '_context_menu', None)
        if old_menu is not None:
            old_menu.destroy()
        menu = tk.Menu(self, tearoff=False)
        self._context_menu = menu
        if self.tree.identify_region(event.x, event.y) == 'heading':
            menu.add_command(label='列幅を初期値に戻す', command=lambda: [self.tree.column(name, width=105) for name in self.tree['columns']])
        else:
            iid = self.tree.identify_row(event.y)
            if not iid:
                menu.destroy()
                self._context_menu = None
                return
            self.tree.selection_set(iid)
            menu.add_command(label='詳細を編集', command=lambda: self.after_idle(self.edit))
            menu.add_command(label='メイン表の該当セルへ移動', command=lambda: self.after_idle(self.navigate))
        # tk_popupは選択完了を待たずに戻る。ここで破棄・grab変更しない。
        menu.bind('<Unmap>', lambda e: self.app.root.after_idle(self.restore_menu_grab))
        menu.tk_popup(event.x_root, event.y_root)

    def restore_menu_grab(self):
        if self.winfo_exists():
            grab = self.grab_current()
            if grab is None or grab is getattr(self, '_context_menu', None):
                self.grab_set()

    def navigate(self):
        if len(self.chosen()) != 1:
            return
        year, month, day, column = self.manager._parse_key(self.chosen()[0][0])
        self.app.current_year = year
        self.app.update_year_display()
        self.app.select_month(month)
        self.app.navigate_to_cell(day, column)
        self.destroy()

    def chosen(self):
        return [self.references[iid] for iid in self.tree.selection()]

    def select_all(self, event=None):
        self.tree.selection_set(self.tree.get_children())
        self.total()
        return 'break'

    def total(self):
        references = self.chosen() or list(self.references.values())
        expense = sum(cash_amount(row) for key, _, row in references if self.manager._parse_key(key)[2] > 0)
        income = sum(cash_amount(row) for key, _, row in references if self.manager._parse_key(key)[2] == 0)
        confirmed_expenses = [row for key, _, row in references if self.manager._parse_key(key)[2] > 0 and not is_pending(row)]
        average = expense / len(confirmed_expenses) if confirmed_expenses else 0
        self.summary.configure(text=f'{"選択" if self.chosen() else "検索結果"} {len(references)}件 / 支出 {expense:,}円 / 収入 {income:,}円 / 支出平均 {average:,.0f}円（未確定は除外）\n薄赤の背景は表示内容が一致する重複候補。見出しクリックで並べ替え、右クリックでメイン表へ移動できます。')

    def mutate(self, action):
        references = self.chosen()
        if action == 'confirm': references = [r for r in references if is_pending(r[2])]
        if not references: return
        verb = '確定' if action == 'confirm' else '削除'
        if not messagebox.askyesno(verb, f'{len(references)}件を{verb}しますか？\n操作前の復元ポイントを保存します。', parent=self): return
        try:
            self.manager.edit_entries(references, action)
            # 一覧での一括操作はメイン画面の旧Undo履歴とは混在させない。
            self.app.undo_stack.clear()
            refreshed(self.app); self.refresh()
        except (ValueError, OSError) as error:
            messagebox.showerror('操作エラー', str(error), parent=self)

    def confirm(self): self.mutate('confirm')
    def delete(self): self.mutate('delete')

    def edit(self):
        if len(self.chosen()) != 1: return
        from ui.transaction_dialog import TransactionDialog
        key, _, _ = self.chosen()[0]
        column = self.manager._parse_key(key)[3]
        dialog = TransactionDialog(self, self.app, key, self.app.get_all_columns()[column])
        self.wait_window(dialog)
        self.grab_set(); self.refresh()



class BudgetDialog(BaseDialog):
    def __init__(self, parent, app):
        super().__init__(parent, '予算・残額・前月比較', 900, 650)
        self.app = app
        self.manager = app.data_manager
        self.body, self.footer = scrollable_form(self)
        bar = ttk.Frame(self.body); bar.pack(fill='x', padx=12, pady=8)
        self.month = ttk.Entry(bar, width=12)
        self.month.insert(0, f'{app.current_year}-{app.current_month:02d}')
        self.month.configure(state='readonly', cursor='hand2')
        self.month.pack(side='left')
        self.month_picker = MonthPicker(
            bar, app.current_year, app.current_month,
            self.pick_month, anchor=self.month, theme='light')
        self.month.bind('<Button-1>', self.open_month_picker)
        ttk.Button(bar, text='年月を表示', command=self.refresh).pack(side='left', padx=8)
        self.tree = ttk.Treeview(self.body, columns=('name', 'budget', 'spent', 'remaining', 'previous', 'change'), show='headings', height=15)
        for column, title in zip(self.tree['columns'], ('分類', '予算', '支出', '残額', '前月支出', '前月との差')):
            self.tree.heading(column, text=title); self.tree.column(column, width=130)
        self.tree.pack(fill='both', expand=True, padx=12)
        self.tree.tag_configure('over', foreground='red')
        self.tree.bind('<<TreeviewSelect>>', self.select)
        self.amount = ttk.Entry(self.body)
        ttk.Label(self.body, text='行を選び、予算（円）を入力。空欄で予算解除。月全体と分類別は独立した予算です。').pack(pady=8)
        self.amount.pack(fill='x', padx=12)
        ttk.Button(self.footer, text='選択行の予算を保存', command=self.save).pack(side='left')
        ttk.Button(self.footer, text='閉じる', command=self.destroy).pack(side='right')
        polish_form(self); self.refresh(); self.show_ready()

    def open_month_picker(self, event=None):
        try:
            date = datetime.datetime.strptime(self.month.get(), '%Y-%m')
            self.month_picker.set(date.year, date.month)
        except ValueError:
            pass
        self.month_picker.open()
        return 'break'

    def pick_month(self, year, month):
        self.month.configure(state='normal')
        self.month.delete(0, 'end')
        self.month.insert(0, f'{year}-{month:02d}')
        self.month.configure(state='readonly')
        self.refresh()

    def refresh(self):
        try: date = datetime.datetime.strptime(self.month.get(), '%Y-%m').date()
        except ValueError:
            messagebox.showwarning('年月', 'YYYY-MMで入力してください。', parent=self); return
        self.period = date.strftime('%Y-%m')
        previous = date - datetime.timedelta(days=1)
        now = self.manager.monthly_spending(date.year, date.month)
        before = self.manager.monthly_spending(previous.year, previous.month)
        budgets = self.manager.budgets.get(self.period, {})
        self.tree.delete(*self.tree.get_children())
        for column, name in enumerate(['月全体'] + self.app.get_all_columns()[1:]):
            key = str(column)
            spent = now[column] if column else sum(now.values())
            last = before[column] if column else sum(before.values())
            budget = budgets.get(key)
            self.tree.insert('', 'end', iid=key, values=(name, budget if budget is not None else '未設定', spent, budget - spent if budget is not None else '—', last, f'{spent-last:+,}'), tags=('over',) if budget is not None and spent > budget else ())

    def select(self, event=None):
        if self.tree.selection():
            value = self.manager.budgets.get(self.period, {}).get(self.tree.selection()[0], '')
            self.amount.delete(0, 'end'); self.amount.insert(0, str(value))

    def save(self):
        if not self.tree.selection(): return
        previous = copy.deepcopy(self.manager.budgets)
        try:
            value = self.amount.get().strip()
            if value and int(value) < 0: raise ValueError('予算は0以上で入力してください。')
            budgets = self.manager.budgets.setdefault(self.period, {})
            key = self.tree.selection()[0]
            if value: budgets[key] = int(value)
            else: budgets.pop(key, None)
            self.manager.save_settings(); self.refresh()
        except (ValueError, OSError) as error:
            self.manager.budgets = previous
            messagebox.showerror('保存エラー', str(error), parent=self)


class RecoveryDialog(BaseDialog):
    def __init__(self, parent, app):
        super().__init__(parent, 'バックアップから復元', 900, 560)
        self.app = app
        self.points = app.data_manager.restore_points()
        self.tree = ttk.Treeview(self, columns=('time', 'label', 'count'), show='headings', selectmode='browse')
        for key, title in zip(self.tree['columns'], ('保存日時', '操作', '明細数')):
            self.tree.heading(key, text=title)
        self.tree.pack(fill='both', expand=True, padx=12, pady=12)
        for index, (_, created, label, count) in enumerate(self.points):
            self.tree.insert('', 'end', iid=str(index), values=(created, label, count))
        ttk.Label(self, text='復元すると全期間の明細・設定を選択した時点に戻します。復元直前の状態も保存します。').pack(pady=8)
        ttk.Button(self, text='選択した時点へ復元', command=self.restore).pack(pady=6)
        ttk.Button(self, text='閉じる', command=self.destroy).pack(pady=6)
        polish_form(self); self.show_ready()

    def restore(self):
        if not self.tree.selection(): return
        point = self.points[int(self.tree.selection()[0])]
        if not messagebox.askyesno('全体を復元', f'{point[1]} の状態に戻します。\nその後の変更も戻ります。よろしいですか？', parent=self): return
        try:
            self.app.data_manager.restore_snapshot(point[0])
            self.app._recreate_treeview(); refreshed(self.app)
            self.app.undo_stack.clear()
            self.destroy()
        except (OSError, ValueError, KeyError) as error:
            messagebox.showerror('復元エラー', str(error), parent=self)
