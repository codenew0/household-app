# ui/base_dialog.py
"""
ダイアログの基底クラス
すべてのダイアログで共通の機能を提供
"""
import tkinter as tk
from config import DialogConfig


class BaseDialog(tk.Toplevel):
    """
    すべてのダイアログの基底クラス。
    
    共通の初期化処理(中央配置、モーダル設定など)を提供し、
    各ダイアログクラスはこのクラスを継承することで、
    一貫したUIと動作を実現する。
    """
    
    def __init__(self, parent, title, width=None, height=None):
        """
        ダイアログの基本設定を行う。
        
        Args:
            parent: 親ウィンドウ
            title: ダイアログのタイトル
            width: ダイアログの幅(デフォルト: DialogConfig.DEFAULT_WIDTH)
            height: ダイアログの高さ(デフォルト: DialogConfig.DEFAULT_HEIGHT)
        """
        parent = parent.winfo_toplevel()
        super().__init__(parent)
        # 子クラスによる部品生成・サイズ変更が終わるまで表示しない。
        self.withdraw()
        self.title(title)
        self.configure(bg='#f0f0f0')
        
        # デフォルトサイズの設定
        if width is None:
            width = DialogConfig.DEFAULT_WIDTH
        if height is None:
            height = DialogConfig.DEFAULT_HEIGHT
        
        # 位置は部品生成後に最終サイズで決める。
        self.geometry(f"{width}x{height}")
        self.minsize(int(width * DialogConfig.MIN_SIZE_RATIO), int(height * 0.67))
        
        # モーダルダイアログとして設定
        self.transient(parent)
        self.resizable(True, True)

    def show_ready(self, initial_focus=None, modal=True):
        """各ダイアログの初期化末尾で呼び、完成した画面だけを表示する。"""
        # after_idleだと構築途中のupdate_idletasksでも発火するため、
        # 表示タイミングは子クラスが明示する（図表の初期描画など）。
        self.update_idletasks()
        self._center_on_parent(self.winfo_width(), self.winfo_height())
        self.update_idletasks()
        self.deiconify()
        if modal:
            self.grab_set()
        self.lift()
        (initial_focus if initial_focus is not None else self).focus_set()
    
    def _center_on_parent(self, width, height):
        """
        ダイアログを親ウィンドウの中央に配置する。
        
        画面中央ではなく、直接の親ウィンドウの中央を基準にする。
        """
        # 親ウィンドウの位置とサイズを取得
        parent_x = self.master.winfo_x()
        parent_y = self.master.winfo_y()
        parent_w = self.master.winfo_width()
        parent_h = self.master.winfo_height()
        
        # 中央配置の計算
        x = parent_x + (parent_w - width) // 2
        y = parent_y + (parent_h - height) // 2
        
        # ダイアログの位置とサイズを設定
        self.geometry(f"{width}x{height}+{x}+{y}")
