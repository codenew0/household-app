"""サブスク、家賃、公共料金、分割払いなどの定期支払い登録。"""
import datetime
import tkinter as tk
from tkinter import ttk, messagebox

from config import DefaultColumns
from ui.base_dialog import BaseDialog
from ui.form_widgets import scrollable_form, polish_form, PaymentMethodPicker


class SubscriptionDialog(BaseDialog):
    """保存形式との互換性のためクラス名は維持し、画面上は定期支払いと呼ぶ。"""

    def __init__(self, parent, app):
        super().__init__(parent, '定期支払い', 1000, 680)
        self.app = app
        self.body, self.footer = scrollable_form(self)
        self.selected_id = None

        table = ttk.Frame(self.body)
        table.pack(fill='both', expand=True, padx=12, pady=12)
        columns = ('kind', 'name', 'category', 'partner', 'amount', 'interval',
                   'count', 'day', 'method', 'enabled', 'next')
        self.tree = ttk.Treeview(table, columns=columns, show='headings',
                                 selectmode='browse', height=8)
        titles = ('種類', '名前', '分類', '支払先', '金額', '間隔', '回数',
                  '支払日', '支払方法', '状態', '次回支払日')
        widths = (120, 140, 110, 130, 85, 80, 70, 70, 120, 65, 105)
        for key, title, width in zip(columns, titles, widths):
            self.tree.heading(key, text=title)
            self.tree.column(key, width=width, minwidth=55, anchor='center')
        self.tree.column('name', anchor='w')
        self.tree.column('partner', anchor='w')
        xscroll = ttk.Scrollbar(table, orient='horizontal', command=self.tree.xview)
        self.tree.configure(xscrollcommand=xscroll.set)
        self.tree.pack(fill='both', expand=True)
        xscroll.pack(fill='x')
        self.tree.bind('<<TreeviewSelect>>', self.select)

        self.forecast = ttk.Label(self.body)
        self.forecast.pack(fill='x', padx=12, pady=(0, 10))

        ttk.Label(self.body, text='登録内容', font=('Yu Gothic UI', 11, 'bold')).pack(
            anchor='w', padx=17, pady=(2, 4))
        form = ttk.Frame(self.body, padding=8)
        form.pack(fill='x', padx=12)
        for column in range(3):
            form.columnconfigure(column, weight=1, uniform='fields')

        specs = (
            ('kind', '種類'), ('name', '名前'), ('partner', '支払先'),
            ('amount', '1回の金額（円）'), ('day', '支払日（1〜31）'), ('method', '支払方法'),
            ('start', '開始日 YYYY-MM-DD'), ('interval_months', '支払・更新間隔（月）'),
            ('category', '登録先の分類'), ('payment_count', '支払回数（空欄＝継続）'),
        )
        self.fields = {}
        for index, (key, title) in enumerate(specs):
            row, column = divmod(index, 3)
            ttk.Label(form, text=title).grid(row=row * 2, column=column, sticky='w', padx=5)
            if key == 'method':
                entry = PaymentMethodPicker(form, app.data_manager)
            elif key == 'kind':
                entry = ttk.Combobox(form, state='readonly',
                                     values=app.data_manager.RECURRING_PAYMENT_TYPES)
                entry.bind('<<ComboboxSelected>>', self._kind_changed)
            elif key == 'category':
                entry = ttk.Combobox(form, state='readonly', values=self._categories())
            elif key in ('day', 'interval_months'):
                entry = ttk.Spinbox(form, from_=1, to=120 if key == 'interval_months' else 31)
            else:
                entry = ttk.Entry(form)
            entry.grid(row=row * 2 + 1, column=column, padx=5, pady=(3, 9), sticky='ew')
            self.fields[key] = entry

        self.enabled = tk.BooleanVar(value=True)
        ttk.Checkbutton(form, text='自動登録を有効にする', variable=self.enabled).grid(
            row=8, column=1, columnspan=2, sticky='w', padx=5)
        ttk.Label(
            self.body,
            text='支払日がない月は月末に登録します。金額が変わる公共料金は、自動登録後の未確定明細で修正できます。',
        ).pack(anchor='w', padx=17, pady=10)

        for title, command in (('新規', self.new), ('保存', self.save),
                               ('削除', self.delete), ('閉じる', self.destroy)):
            ttk.Button(self.footer, text=title, command=command).pack(side='left', padx=6)
        polish_form(self)
        self.refresh()
        self.new()
        self.show_ready(self.fields['name'])

    def _categories(self):
        return [name for name in DefaultColumns.ITEMS + self.app.data_manager.custom_columns
                if name != '日付']

    @staticmethod
    def _set(entry, value):
        if isinstance(entry, ttk.Combobox):
            entry.set(value)
        else:
            entry.delete(0, 'end')
            entry.insert(0, value)

    def _kind_changed(self, event=None):
        kind = self.fields['kind'].get()
        self.fields['category'].set(
            self.app.data_manager.RECURRING_DEFAULT_CATEGORIES.get(kind, 'サブスク'))

    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        forecast = self.app.data_manager.subscription_forecast()
        dates = {item['id']: date.isoformat() for item, date in forecast}
        year_total = self.app.data_manager.recurring_payment_year_total()
        self.forecast.configure(
            text=f'有効な定期支払い {len(forecast)}件 / 今後12か月の予定額 {year_total:,}円')
        for raw in self.app.data_manager.subscriptions:
            item = self.app.data_manager.recurring_payment_values(raw)
            count = f"{item['payment_count']}回" if item['payment_count'] else '継続'
            self.tree.insert('', 'end', iid=item['id'], values=(
                item['kind'], item['name'], item['category'], item['partner'], item['amount'],
                f"{item['interval_months']}か月", count, item['day'], item['method'],
                '有効' if item['enabled'] else '停止', dates.get(item['id'], '—')))

    def new(self):
        self.selected_id = None
        defaults = {
            'kind': 'サブスクリプション', 'name': '', 'partner': '', 'amount': '',
            'day': '1', 'method': '', 'start': datetime.date.today().isoformat(),
            'interval_months': '1', 'category': 'サブスク', 'payment_count': '',
        }
        self.fields['category']['values'] = self._categories()
        for key, entry in self.fields.items():
            self._set(entry, defaults[key])
        self.enabled.set(True)
        self.tree.selection_remove(self.tree.selection())

    def select(self, event=None):
        selected = self.tree.selection()
        if not selected:
            return
        self.selected_id = selected[0]
        raw = next(s for s in self.app.data_manager.subscriptions if s['id'] == self.selected_id)
        item = self.app.data_manager.recurring_payment_values(raw)
        values = {**item, 'start': item['start'],
                  'payment_count': item['payment_count'] or ''}
        self.fields['category']['values'] = self._categories()
        for key, entry in self.fields.items():
            self._set(entry, values[key])
        self.enabled.set(item['enabled'])

    def save(self):
        values = {key: entry.get().strip() for key, entry in self.fields.items()}
        values['start_date'] = values.pop('start')
        try:
            record = self.app.data_manager.save_subscription(
                **values, subscription_id=self.selected_id, enabled=self.enabled.get())
            # 定義保存後の自動生成だけが失敗しても、再試行は同じ定義を更新する。
            self.selected_id = record['id']
            self.app.data_manager.generate_due_subscriptions()
        except (ValueError, OSError) as error:
            messagebox.showerror('登録エラー', str(error), parent=self)
            return
        self.app._show_month(self.app.current_month)
        self.refresh()
        self.new()

    def delete(self):
        if self.selected_id and messagebox.askyesno(
                '削除', '定期支払いの設定を削除しますか？登録済みの明細は残ります。', parent=self):
            previous = self.app.data_manager.subscriptions[:]
            try:
                self.app.data_manager.subscriptions = [
                    item for item in previous if item['id'] != self.selected_id]
                self.app.data_manager.save_settings()
            except OSError as error:
                self.app.data_manager.subscriptions = previous
                messagebox.showerror('削除エラー', str(error), parent=self)
                return
            self.refresh()
            self.new()
