"""毎月のサブスクリプション登録。"""
import datetime
import tkinter as tk
from tkinter import ttk, messagebox
from ui.base_dialog import BaseDialog
from ui.form_widgets import scrollable_form, polish_form, PaymentMethodPicker


class SubscriptionDialog(BaseDialog):
    def __init__(self, parent, app):
        super().__init__(parent, 'サブスクリプション', 850, 520)
        self.app = app
        self.body, self.footer = scrollable_form(self)
        self.selected_id = None
        self.tree = ttk.Treeview(self.body, columns=('name', 'partner', 'amount', 'day', 'method', 'enabled'),
                                 show='headings', selectmode='browse', height=8)
        for key, title in zip(self.tree['columns'], ('名前', '支払先', '金額', '支払日', '支払方法', '有効')):
            self.tree.heading(key, text=title)
            self.tree.column(key, width=110)
        self.tree.pack(fill='both', expand=True, padx=12, pady=12)
        self.tree.bind('<<TreeviewSelect>>', self.select)
        form = ttk.Frame(self.body)
        form.pack(fill='x', padx=12)
        for column in range(3):
            form.columnconfigure(column, weight=1, uniform='fields')
        self.fields = {}
        for index, (key, title) in enumerate((('name', '名前'), ('partner', '支払先'), ('amount', '金額（円）'),
                                             ('day', '支払日（1〜31）'), ('method', '支払方法'), ('start', '開始日 YYYY-MM-DD'))):
            ttk.Label(form, text=title).grid(row=index // 3 * 2, column=index % 3, sticky='w')
            entry = PaymentMethodPicker(form, app.data_manager) if key == 'method' else ttk.Entry(form, width=28)
            entry.grid(row=index // 3 * 2 + 1, column=index % 3, padx=5, pady=5, sticky='ew')
            self.fields[key] = entry
        self.enabled = tk.BooleanVar(value=True)
        ttk.Checkbutton(form, text='自動登録を有効にする', variable=self.enabled).grid(row=4, column=0, sticky='w')
        ttk.Label(self.body, text='毎月払い。存在しない支払日は月末。未起動中の分は次回起動時に登録します。').pack(pady=8)
        buttons = self.footer
        for title, command in (('新規', self.new), ('保存', self.save), ('削除', self.delete), ('閉じる', self.destroy)):
            ttk.Button(buttons, text=title, command=command).pack(side='left', padx=6)
        polish_form(self)
        self.refresh()
        self.new()
        self.show_ready(self.fields['name'])

    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        for item in self.app.data_manager.subscriptions:
            self.tree.insert('', 'end', iid=item['id'], values=(item['name'], item['partner'], item['amount'],
                             item['day'], item['method'], '有効' if item['enabled'] else '停止'))

    def new(self):
        self.selected_id = None
        for key, entry in self.fields.items():
            entry.delete(0, 'end')
            entry.insert(0, datetime.date.today().isoformat() if key == 'start' else ('1' if key == 'day' else ''))
        self.enabled.set(True)

    def select(self, event=None):
        selected = self.tree.selection()
        if not selected:
            return
        self.selected_id = selected[0]
        item = next(s for s in self.app.data_manager.subscriptions if s['id'] == self.selected_id)
        for key, entry in self.fields.items():
            entry.delete(0, 'end')
            entry.insert(0, item[key])
        self.enabled.set(item['enabled'])

    def save(self):
        values = {key: entry.get().strip() for key, entry in self.fields.items()}
        values['start_date'] = values.pop('start')
        try:
            self.app.data_manager.save_subscription(**values, subscription_id=self.selected_id, enabled=self.enabled.get())
            self.app.data_manager.generate_due_subscriptions()
        except (ValueError, OSError) as error:
            messagebox.showerror('登録エラー', str(error), parent=self)
            return
        self.app._show_month(self.app.current_month)
        self.refresh()
        self.new()

    def delete(self):
        if self.selected_id and messagebox.askyesno('削除', 'サブスク定義を削除しますか？登録済みの明細は残ります。', parent=self):
            self.app.data_manager.subscriptions = [s for s in self.app.data_manager.subscriptions if s['id'] != self.selected_id]
            self.app.data_manager.save_settings()
            self.refresh()
            self.new()
