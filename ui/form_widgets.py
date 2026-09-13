"""共通の配色、固定フッター、支払方法の選択部品。"""
import tkinter as tk
from tkinter import ttk, messagebox
from ui.base_dialog import BaseDialog


def setup_form_styles(widget):
    style = ttk.Style(widget)
    style.configure('Form.TFrame', background='#f5f7fb')
    style.configure('Form.TLabel', background='#f5f7fb', foreground='#334155', font=('Yu Gothic UI', 10))
    style.configure('Form.TCheckbutton', background='#f5f7fb', foreground='#334155', font=('Yu Gothic UI', 10))
    for name, bg, fg, hover in (
            ('Primary', '#2563eb', '#ffffff', '#1d4ed8'),
            ('Secondary', '#e2e8f0', '#1e293b', '#cbd5e1'),
            ('Danger', '#fee2e2', '#b91c1c', '#fecaca'),
            ('Toolbar', '#334155', '#f8fafc', '#475569')):
        key = name + '.Form.TButton'
        style.configure(key, background=bg, foreground=fg, padding=(16, 9),
                        font=('Yu Gothic UI', 10, 'bold'), borderwidth=0, relief='flat')
        style.map(key, background=[('disabled', '#e2e8f0'), ('active', hover)],
                  foreground=[('disabled', '#94a3b8'), ('active', fg)])


def scrollable_form(dialog, split_image=False):
    """フッター領域を先に確保し、長いフォームだけをスクロールさせる。"""
    setup_form_styles(dialog)
    dialog.configure(bg='#f5f7fb')
    width = min(1000, dialog.winfo_screenwidth() - 60)
    height = min(800, dialog.winfo_screenheight() - 100)
    dialog.geometry(f'{width}x{height}')
    dialog.minsize(min(720, width), min(460, height))
    footer = ttk.Frame(dialog, style='Form.TFrame', padding=(16, 12))
    footer.pack(side='bottom', fill='x')
    ttk.Separator(dialog).pack(side='bottom', fill='x')
    viewport = ttk.Frame(dialog, style='Form.TFrame')
    viewport.pack(fill='both', expand=True)
    if split_image:
        panes = ttk.Panedwindow(viewport, orient='horizontal')
        panes.pack(fill='both', expand=True)
        dialog.image_panel = ttk.Frame(panes, style='Form.TFrame', width=320)
        right = ttk.Frame(panes, style='Form.TFrame', width=650)
        panes.add(dialog.image_panel, weight=1)
        panes.add(right, weight=2)
        viewport = right
    canvas = tk.Canvas(viewport, bg='#f5f7fb', highlightthickness=0)
    scrollbar = ttk.Scrollbar(viewport, orient='vertical', command=canvas.yview)
    scrollbar.pack(side='right', fill='y')
    canvas.pack(side='left', fill='both', expand=True)
    canvas.configure(yscrollcommand=scrollbar.set)
    body = ttk.Frame(canvas, style='Form.TFrame', padding=(4, 8))
    window = canvas.create_window((0, 0), window=body, anchor='nw')
    body.bind('<Configure>', lambda event: canvas.configure(scrollregion=canvas.bbox('all')))
    canvas.bind('<Configure>', lambda event: canvas.itemconfigure(window, width=event.width))
    def wheel(event):
        if isinstance(event.widget, (ttk.Treeview, ttk.Combobox)):
            return
        canvas.yview_scroll(-int(event.delta / 120), 'units')
    dialog.bind('<MouseWheel>', wheel, add=True)
    dialog.form_canvas = canvas
    return body, footer


def polish_form(widget):
    for child in widget.winfo_children():
        if getattr(child, 'preserve_styles', False):
            continue
        if isinstance(child, ttk.Button):
            text = child.cget('text')
            kind = 'Danger' if '削除' in text else 'Primary' if any(w in text for w in ('保存', '登録', '追加', '読み込む')) else 'Secondary'
            child.configure(style=kind + '.Form.TButton')
        elif isinstance(child, ttk.Frame):
            child.configure(style='Form.TFrame')
        elif isinstance(child, ttk.Label):
            child.configure(style='Form.TLabel')
        elif isinstance(child, ttk.Checkbutton):
            child.configure(style='Form.TCheckbutton')
        polish_form(child)


def manage_payment_methods(parent, manager):
    dialog = PaymentMethodsDialog(parent, manager)
    parent.wait_window(dialog)
    if parent.winfo_exists():
        parent.grab_set()


class PaymentMethodPicker(ttk.Frame):
    def __init__(self, parent, manager):
        super().__init__(parent, style='Form.TFrame')
        self.manager = manager
        self.combo = ttk.Combobox(self, state='readonly', width=19, postcommand=self.refresh)
        self.combo.pack(side='left', fill='x', expand=True)
        ttk.Button(self, text='管理', width=5, style='Secondary.Form.TButton',
                   command=self.manage).pack(side='left', padx=(6, 0))
        self.refresh()

    def refresh(self):
        self.combo['values'] = list(dict.fromkeys([''] + self.manager.payment_methods + [self.get()]))

    def manage(self):
        manage_payment_methods(self.winfo_toplevel(), self.manager)
        self.refresh()

    def get(self):
        return self.combo.get()

    def set(self, value):
        self.combo.set(value)
        self.refresh()

    def delete(self, first, last=None):
        self.set('')

    def insert(self, index, value):
        self.set(value)


class PaymentMethodsDialog(BaseDialog):
    def __init__(self, parent, manager):
        super().__init__(parent, '支払方法の管理', 450, 400)
        setup_form_styles(self)
        self.manager = manager
        self.configure(bg='#f5f7fb')
        ttk.Label(self, text='支払方法を登録', style='Form.TLabel').pack(anchor='w', padx=18, pady=(18, 8))
        self.name = ttk.Entry(self)
        self.name.pack(fill='x', padx=18, pady=6)
        ttk.Button(self, text='追加', style='Primary.Form.TButton', command=self.add).pack(anchor='e', padx=18)
        self.listbox = tk.Listbox(self, bd=0, highlightthickness=1, highlightbackground='#cbd5e1',
                                  font=('Yu Gothic UI', 11), selectbackground='#dbeafe', selectforeground='#1e3a8a')
        self.listbox.pack(fill='both', expand=True, padx=18, pady=12)
        buttons = ttk.Frame(self, style='Form.TFrame')
        buttons.pack(fill='x', padx=18, pady=(0, 14))
        ttk.Button(buttons, text='選択を削除', style='Danger.Form.TButton', command=self.remove).pack(side='left')
        ttk.Button(buttons, text='閉じる', style='Secondary.Form.TButton', command=self.destroy).pack(side='right')
        self.refresh()
        self.show_ready(self.name)

    def refresh(self):
        self.listbox.delete(0, 'end')
        for name in self.manager.payment_methods:
            self.listbox.insert('end', name)

    def add(self):
        try:
            self.manager.add_payment_method(self.name.get())
        except (ValueError, OSError) as error:
            messagebox.showerror('登録エラー', str(error), parent=self)
            return
        self.name.delete(0, 'end')
        self.refresh()

    def remove(self):
        selected = self.listbox.curselection()
        if selected:
            try:
                self.manager.remove_payment_method(self.listbox.get(selected[0]))
            except OSError as error:
                messagebox.showerror('保存エラー', str(error), parent=self)
                return
            self.refresh()
