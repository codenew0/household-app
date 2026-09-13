"""年月表示、月移動、年ごとの月グリッドをまとめた共通部品。"""
import datetime
import tkinter as tk
from tkinter import ttk
from config import ColorTheme


def year_label(year):
    """西暦と和暦を併記した年表示を返す。"""
    eras = ((2019, 5, '令和', 2018), (1989, 1, '平成', 1988),
            (1926, 12, '昭和', 1925), (1912, 7, '大正', 1911), (1868, 1, '明治', 1867))
    for first_year, first_month, name, offset in eras:
        if (year, 12) >= (first_year, first_month):
            number = year - offset
            return f'{year}年（{name}{"元" if number == 1 else number}年）'
    return f'{year}年'


class MonthPicker(ttk.Frame):
    def __init__(self, parent, year, month, command, anchor=None, theme='light'):
        super().__init__(parent)
        self.preserve_styles = True
        self.year, self.month = year, month
        self.command = command
        self.popup = None
        self.external_anchor = anchor is not None
        if theme == 'main':
            background, foreground, selected, hover, muted = ColorTheme.BG_TERTIARY, 'white', ColorTheme.BG_SECONDARY, '#585b70', '#a6adc8'
        else:
            background, foreground, selected, hover, muted = 'white', 'black', '#dbeafe', '#eef2f7', '#9ca3af'
        self.background = background
        self.styles = {kind: f'{theme}.Calendar.{kind}' for kind in ('TFrame', 'TLabel', 'TButton', 'Selected.TButton')}
        style = ttk.Style(self)
        style.configure(self.styles['TFrame'], background=background)
        style.configure(self.styles['TLabel'], background=background, foreground=foreground)
        for kind, color in (('TButton', background), ('Selected.TButton', selected)):
            name = self.styles[kind]
            style.configure(name, padding=(12, 10), font=('Yu Gothic UI', 11), background=color, foreground=foreground, borderwidth=0, relief='flat')
            style.map(name, background=[('pressed', selected), ('active', hover)], foreground=[('disabled', muted), ('active', foreground)])
        self.label = anchor if anchor is not None else ttk.Button(self, command=self.open, style='Month.TButton')
        if anchor is None:
            self.label.pack()
        self.set(year, month)

    def set(self, year, month):
        self.year, self.month = year, month
        if not self.external_anchor:
            self.label.configure(text=f'{year}年 {month:02d}月')

    def choose(self, year, month):
        self.close()
        self.set(year, month)
        self.command(year, month)

    def move(self, delta):
        value = self.year * 12 + self.month - 1 + delta
        year, index = divmod(value, 12)
        if 1900 <= year <= 9999:
            self.choose(year, index + 1)

    def today(self):
        now = datetime.date.today()
        self.choose(now.year, now.month)

    def open(self):
        if self.popup is not None:
            self.close(); return
        self.previous_grab = self.grab_current()
        popup = self.popup = tk.Toplevel(self)
        popup.withdraw()
        popup.overrideredirect(True)
        popup.transient(self.winfo_toplevel())
        popup.configure(background=self.background, padx=1, pady=1)
        panel = ttk.Frame(popup, style=self.styles['TFrame'])
        panel.pack(fill='both', expand=True)
        self.browsing_year = self.year
        year_bar = ttk.Frame(panel, style=self.styles['TFrame'])
        year_bar.pack(fill='x', padx=8, pady=8)
        self.year_previous = ttk.Button(year_bar, text='◀', width=3, style=self.styles['TButton'], command=lambda: self.change_year(-1))
        self.year_previous.pack(side='left')
        self.year_label = ttk.Label(year_bar, anchor='center', font=('Yu Gothic UI', 12, 'bold'), style=self.styles['TLabel'])
        self.year_label.pack(side='left', fill='x', expand=True)
        self.year_next = ttk.Button(year_bar, text='▶', width=3, style=self.styles['TButton'], command=lambda: self.change_year(1))
        self.year_next.pack(side='right')
        self.months_frame = ttk.Frame(panel, style=self.styles['TFrame'])
        self.months_frame.pack(fill='both', expand=True, padx=8, pady=(0, 8))
        popup.bind('<Escape>', lambda e: self.close() or 'break')
        popup.bind('<Button-1>', self.outside)
        self.render()
        popup.update_idletasks()
        x = self.label.winfo_rootx()
        y = self.label.winfo_rooty() + self.label.winfo_height()
        width, height = popup.winfo_reqwidth(), popup.winfo_reqheight()
        x = min(x, self.winfo_screenwidth() - width)
        if y + height > self.winfo_screenheight():
            y = self.label.winfo_rooty() - height
        popup.geometry(f'{width}x{height}+{x}+{max(0, y)}')
        popup.deiconify()
        popup.grab_set()
        popup.focus_set()

    def change_year(self, delta):
        year = self.browsing_year + delta
        if 1900 <= year <= 9999:
            self.browsing_year = year
            self.render()

    def render(self):
        year = self.browsing_year
        self.year_label.configure(text=year_label(year))
        self.year_previous.configure(state='disabled' if year == 1900 else 'normal')
        self.year_next.configure(state='disabled' if year == 9999 else 'normal')
        for child in self.months_frame.winfo_children():
            child.destroy()
        for column in range(4):
            self.months_frame.columnconfigure(column, weight=1)
        for month in range(1, 13):
            ttk.Button(self.months_frame, text=f'{month}月', width=6,
                       style=self.styles['Selected.TButton' if (year, month) == (self.year, self.month) else 'TButton'],
                       command=lambda y=year, m=month: self.choose(y, m)).grid(
                           row=(month - 1) // 4, column=(month - 1) % 4, sticky='ew', padx=2, pady=2)

    def outside(self, event):
        popup = self.popup
        if popup and not (popup.winfo_rootx() <= event.x_root < popup.winfo_rootx() + popup.winfo_width()
                          and popup.winfo_rooty() <= event.y_root < popup.winfo_rooty() + popup.winfo_height()):
            self.close()
            return 'break'

    def close(self):
        if self.popup is None: return
        popup, self.popup = self.popup, None
        if popup.grab_current() is popup:
            popup.grab_release()
        popup.destroy()
        if self.previous_grab is not None and self.previous_grab.winfo_exists():
            self.previous_grab.grab_set()
        owner = self.label.winfo_toplevel()
        def restore_focus():
            if not owner.winfo_exists() or not self.label.winfo_exists():
                return
            grab = owner.grab_current()
            if grab is not None and grab.winfo_toplevel() is not owner:
                return
            # Windowsではpopup破棄後に親のネイティブフォーカスも戻す。
            owner.focus_force()
            self.label.focus_set()
        owner.after_idle(restore_focus)
