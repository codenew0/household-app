"""OCR結果を確認し、分類ごとの明細として登録する画面。"""
import copy
import datetime
import queue
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from ui.base_dialog import BaseDialog
from ui.form_widgets import scrollable_form, polish_form, PaymentMethodPicker
from utils.receipt_parser import recognize_products
from PIL import Image, ImageOps, ImageTk
from models.transactions import validate


class ReceiptDialog(BaseDialog):
    def __init__(self, parent, app):
        super().__init__(parent, 'レシートOCR', 980, 720)
        self.app = app
        self.body, self.footer = scrollable_form(self, split_image=True)
        self.geometry(f'{min(1400, self.winfo_screenwidth() - 80)}x{min(850, self.winfo_screenheight() - 100)}')
        self.staged = []
        self.results = queue.Queue()
        self.ocr_items = []
        self.source_image = None
        self.ocr_result = None
        ttk.Label(self.image_panel, text='元画像 / 選択した商品の行を拡大').pack(anchor='w', padx=8, pady=8)
        self.image_canvas = tk.Canvas(self.image_panel, highlightthickness=0, bg='#e2e8f0')
        image_scroll = ttk.Scrollbar(self.image_panel, orient='vertical', command=self.image_canvas.yview)
        image_scroll.pack(side='right', fill='y')
        self.image_canvas.pack(fill='both', expand=True, padx=4)
        self.image_canvas.configure(yscrollcommand=image_scroll.set)
        self.image_canvas.bind('<Configure>', lambda event: self.render_image())
        bar = ttk.Frame(self.body)
        bar.pack(fill='x', padx=12, pady=10)
        self.load_button = ttk.Button(bar, text='レシート画像を読み込む', command=self.load_image)
        self.load_button.pack(side='left')
        self.status = ttk.Label(bar, text='OCR結果を選択し、分類・金額を確認して明細に追加してください。')
        self.status.configure(wraplength=510)
        self.status.pack(side='left', padx=12)
        self.lines = ttk.Treeview(self.body, columns=('name', 'amount', 'note'), show='headings', height=8, selectmode='browse')
        for key, title, width in (('name', '商品名（認識候補）', 510), ('amount', '金額（円）', 100), ('note', '確認', 260)):
            self.lines.heading(key, text=title)
            self.lines.column(key, width=width)
        self.lines.pack(fill='both', expand=True, padx=12)
        self.lines.bind('<<TreeviewSelect>>', self.select_line)
        self.source_preview = ttk.Label(self.image_panel, text='商品を選択すると該当行を拡大します。')
        self.source_preview.pack(fill='x', padx=12, pady=5)
        ttk.Button(self.body, text='認識した全文・その他の行を表示', command=self.show_raw).pack(anchor='e', padx=12)
        form = ttk.Frame(self.body)
        form.pack(fill='x', padx=12, pady=10)
        for column in range(2):
            form.columnconfigure(column, weight=1, uniform='fields')
        self.fields = {}
        for index, (key, title) in enumerate((('date', '日付 YYYY-MM-DD'), ('partner', '支払先'),
                                             ('amount', 'ポイント利用前の金額（円）'), ('points', '利用ポイント（1pt＝1円）'),
                                             ('method', '支払方法'), ('memo', '詳細（メモ）'))):
            ttk.Label(form, text=title).grid(row=index // 2 * 2, column=index % 2, sticky='w')
            entry = PaymentMethodPicker(form, app.data_manager) if key == 'method' else ttk.Entry(form, width=28)
            entry.grid(row=index // 2 * 2 + 1, column=index % 2, sticky='ew', padx=4, pady=4)
            self.fields[key] = entry
        self.fields['date'].insert(0, datetime.date.today().isoformat())
        self.fields['points'].insert(0, '0')
        self.category = ttk.Combobox(form, values=app.get_all_columns()[1:], state='readonly', width=30)
        self.category.current(0)
        ttk.Label(form, text='分類').grid(row=6, column=0, sticky='w')
        self.category.grid(row=7, column=0, sticky='w')
        ttk.Button(form, text='確認した明細を追加 ↓', command=self.add_row).grid(row=7, column=1)
        ttk.Button(form, text='支払先から分類・支払方法を提案', command=self.suggest_partner).grid(row=8, column=0, columnspan=2, pady=8)
        self.preview = ttk.Treeview(self.body, columns=('date', 'category', 'partner', 'amount', 'points', 'memo'), show='headings', height=7)
        for name, title in zip(self.preview['columns'], ('日付', '分類', '支払先', '利用前金額', 'ポイント', 'メモ')):
            self.preview.heading(name, text=title)
            self.preview.column(name, width=145)
        self.preview.pack(fill='both', expand=True, padx=12, pady=8)
        buttons = self.footer
        ttk.Button(buttons, text='選択明細を削除', command=self.remove_rows).pack(side='left', padx=6)
        self.save_button = ttk.Button(buttons, text='明細を登録', command=self.save)
        self.save_button.pack(side='right', padx=6)
        ttk.Button(buttons, text='キャンセル', command=self.destroy).pack(side='right', padx=6)

        polish_form(self)
        self.show_ready()

    def load_image(self):
        path = filedialog.askopenfilename(parent=self, filetypes=[('画像', '*.png *.jpg *.jpeg *.bmp *.tif *.tiff')])
        if not path:
            return
        try:
            with Image.open(path) as image:
                self.source_image = ImageOps.exif_transpose(image).copy()
            self.render_image()
        except OSError as error:
            messagebox.showerror('画像エラー', str(error), parent=self)
            return
        self.load_button.configure(state='disabled')
        self.lines.delete(*self.lines.get_children())
        self.ocr_items = []
        self.ocr_result = None
        self.source_preview.configure(image='', text='文字認識中…')
        self.status.configure(text='文字認識中…')
        def work():
            try:
                self.results.put((True, recognize_products(path)))
            except Exception as error:
                self.results.put((False, str(error)))
        threading.Thread(target=work, daemon=True).start()
        self.after(100, self.poll)

    def poll(self):
        try:
            success, result = self.results.get_nowait()
        except queue.Empty:
            self.after(100, self.poll)
            return
        self.load_button.configure(state='normal')
        if not success:
            self.status.configure(text='文字認識に失敗しました。')
            messagebox.showerror('OCRエラー', result, parent=self)
            return
        self.ocr_result = result
        self.ocr_items = result['items']
        self.lines.delete(*self.lines.get_children())
        for index, item in enumerate(self.ocr_items):
            self.lines.insert('', 'end', iid=str(index), values=(item['name'], item['amount'] if item['amount'] is not None else '', item['note']))
        total = sum(item['amount'] or 0 for item in self.ocr_items)
        missing = sum(item['amount'] is None for item in self.ocr_items)
        self.status.configure(text=f'商品候補 {len(self.ocr_items)}件／金額計 {total:,}円／金額未読 {missing}件。原画像と確認してください。')

    def select_line(self, event=None):
        selected = self.lines.selection()
        if not selected:
            return
        item = self.ocr_items[int(selected[0])]
        self.fields['memo'].delete(0, 'end')
        self.fields['memo'].insert(0, item['name'])
        self.fields['amount'].delete(0, 'end')
        if item['amount'] is not None:
            self.fields['amount'].insert(0, str(item['amount']))
        # 補正前の候補も残し、商品名を比較して編集できるようにする。
        alternative = item.get('alternative_name')
        if alternative and alternative != item['name']:
            self.status.configure(text=f"別の認識候補: {alternative} ／ 元画像を見てメモ欄を修正してください。")
        self.fields['points'].delete(0, 'end')
        self.fields['points'].insert(0, '0')
        if self.source_image is not None:
            bounds = item['bounds']
            sy = self.source_image.height / self.ocr_result['height']
            crop = self.source_image.crop((0, int(bounds['y'] * sy), self.source_image.width,
                                           int((bounds['y'] + bounds['height']) * sy)))
            crop.thumbnail((max(250, self.image_panel.winfo_width() - 20), 130))
            self.preview_image = ImageTk.PhotoImage(crop)
            self.source_preview.configure(image=self.preview_image, text='')
            self.render_image(item)

    def render_image(self, selected=None):
        if self.source_image is None:
            return
        width = max(180, self.image_canvas.winfo_width() - 10)
        scale = width / self.source_image.width
        image = self.source_image.resize((width, max(1, round(self.source_image.height * scale))))
        self.full_preview_image = ImageTk.PhotoImage(image)
        self.image_canvas.delete('all')
        self.image_canvas.create_image(0, 0, anchor='nw', image=self.full_preview_image)
        self.image_canvas.configure(scrollregion=(0, 0, image.width, image.height))
        if selected and self.ocr_result:
            bounds = selected['bounds']
            top = bounds['y'] * self.source_image.height / self.ocr_result['height'] * scale
            height = bounds['height'] * self.source_image.height / self.ocr_result['height'] * scale
            self.image_canvas.create_rectangle(1, top, width - 1, top + height, outline='#2563eb', width=3)
            self.image_canvas.yview_moveto(max(0, (top - 50) / image.height))

    def suggest_partner(self):
        suggestion = self.app.data_manager.partner_suggestion(self.fields['partner'].get())
        if not suggestion:
            messagebox.showinfo('入力候補', 'この支払先の確定済み履歴がありません。', parent=self)
            return
        column, method = suggestion
        columns = self.app.get_all_columns()
        if 1 <= column < len(columns) and messagebox.askyesno('入力候補', f'分類: {columns[column]}\n支払方法: {method}\n反映しますか？', parent=self):
            self.category.current(column - 1)
            self.fields['method'].set(method)

    def show_raw(self):
        if not self.ocr_result:
            return
        window = BaseDialog(self, '認識全文（商品以外の行を含む）', 900, 500)
        text = tk.Text(window, width=90, height=25)
        text.pack(fill='both', expand=True)
        text.insert('end', '\n'.join(self.ocr_result['raw']) + '\n\nその他の行:\n' + '\n'.join(self.ocr_result['other']))
        text.configure(state='disabled')
        window.show_ready(modal=False)

    def add_row(self):
        values = {key: entry.get().strip() for key, entry in self.fields.items()}
        try:
            day = datetime.date.fromisoformat(values['date'])
            if not values['partner'] or not values['amount']:
                raise ValueError('支払先と金額を入力してください。')
            row = validate([values['partner'], values['amount'], values['memo'], values['points'], values['method']])
        except ValueError as error:
            messagebox.showwarning('入力エラー', str(error), parent=self)
            return
        column = self.category.current() + 1
        key = f'{day.year}-{day.month}-{day.day}-{column}'
        item = self.preview.insert('', 'end', values=(day.isoformat(), self.category.get(), row[0], row[1], row[3], row[2]))
        self.staged.append((item, key, row))

    def remove_rows(self):
        selected = self.preview.selection()
        self.staged = [item for item in self.staged if item[0] not in selected]
        for item in selected:
            self.preview.delete(item)

    def save(self):
        if not self.staged:
            messagebox.showwarning('明細なし', '確認した明細を一覧に追加してください。', parent=self)
            return
        manager = self.app.data_manager
        keys = {key for _, key, _ in self.staged}
        old = [(key, copy.deepcopy(manager.get_transaction_data(key)) or None) for key in keys]
        previous_partners = manager.transaction_partners.copy()
        for _, key, row in self.staged:
            manager.set_transaction_data(key, list(manager.get_transaction_data(key)) + [row])
        try:
            manager.save_transactions(keys)
            manager.transaction_partners.update(row[0] for _, _, row in self.staged)
            manager.save_settings()
        except OSError as error:
            for key, rows in old:
                manager.set_transaction_data(key, rows or [])
            manager.transaction_partners = previous_partners
            try:
                manager.save_transactions(keys, create_snapshot=False)
                manager.save_settings()
            except OSError:
                pass
            messagebox.showerror('保存エラー', f'保存できませんでした。再試行の前に保存先を確認してください。\n{error}', parent=self)
            return
        self.app._save_undo_state('paste', old)
        self.app._show_month(self.app.current_month)
        self.destroy()
