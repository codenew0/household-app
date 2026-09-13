# config.py
"""
アプリケーション設定と定数定義
"""
import os
import sys
from datetime import datetime

# =====================================================
# ファイルパス設定
# =====================================================
def get_app_path():
    """
    実行環境に応じたアプリケーションのルートパスを取得する
    """
    # PyInstallerでEXE化されているかどうかを判定
    if getattr(sys, 'frozen', False):
        # EXEの場合：実行ファイル（.exe）のあるディレクトリを取得
        return os.path.dirname(sys.executable)
    else:
        # スクリプトの場合：このファイル（config.py）のあるディレクトリを取得
        return os.path.dirname(os.path.abspath(__file__))

# スクリプトのディレクトリを取得
SCRIPT_DIR = get_app_path()

# JSONファイルをスクリプトと同じディレクトリに保存
JSON_DIR = os.path.join(SCRIPT_DIR, "household_json")
SETTINGS_FILE = os.path.join(JSON_DIR, "settings.json")

# 新しいフォルダ構造のルートディレクトリ
DATA_ROOT_DIR = os.path.join(JSON_DIR, "data")

# =====================================================
# アプリケーションバージョン
# =====================================================
APP_VERSION = "3.0"
APP_TITLE = "💰 家計管理 2026"

# =====================================================
# ウィンドウ設定
# =====================================================
class WindowConfig:
    """メインウィンドウの設定"""
    WIDTH = 1400
    HEIGHT = 1020
    MIN_WIDTH = 1200
    MIN_HEIGHT = 800
    RESIZABLE = (True, True)

# =====================================================
# カラーテーマ設定
# =====================================================
class ColorTheme:
    """ダークテーマベースのカラー設定"""
    # 背景色
    BG_PRIMARY = '#1e1e2e'      # メイン背景色
    BG_SECONDARY = '#313244'    # セカンダリ背景色
    BG_TERTIARY = '#45475a'     # 第三背景色
    
    # アクセントカラー
    ACCENT = '#89b4fa'          # アクセントカラー(青)
    ACCENT_GREEN = '#a6e3a1'    # 緑のアクセント
    ACCENT_RED = '#f38ba8'      # 赤のアクセント
    
    # テキスト色
    TEXT_PRIMARY = '#cdd6f4'    # プライマリテキスト
    TEXT_SECONDARY = '#bac2de'  # セカンダリテキスト
    
    # その他
    BORDER = '#585b70'          # ボーダー色
    HOVER = '#74c0fc'           # ホバー時の色

# =====================================================
# Treeview設定
# =====================================================
class TreeviewConfig:
    """Treeviewの表示設定"""
    # タグ設定
    TAG_TOTAL = "TOTAL"           # 合計行タグ
    TAG_SUMMARY = "SUMMARY"       # まとめ行タグ
    TAG_NORMAL = "normal_row"     # 通常行(偶数)
    TAG_ODD = "odd_row"           # 通常行(奇数)
    TAG_DUPLICATE = "duplicate"   # 重複データタグ
    TAG_SAT = "saturday"          # 土曜日タグ
    TAG_SUN = "sunday"            # 日曜日タグ
    
    # 背景色
    BG_TOTAL = "#fff3cd"         # 合計行背景色
    BG_SUMMARY = "#d4edda"       # まとめ行背景色
    BG_NORMAL = "white"            # 通常行背景色
    BG_ODD = "#f8f9fa"           # 奇数行背景色
    BG_DUPLICATE = "#ffcccc"     # 重複データ背景色
    BG_SAT = "#fffef7"           # 土曜日背景色
    BG_SUN = "#f8f8f2"           # 日曜日背景色
    
    # 列幅設定
    COL_WIDTH_DATE = 60          # 日付列
    COL_WIDTH_DATA = 80          # データ列
    COL_WIDTH_BUTTON = 40        # ボタン列
    
    # その他
    ROW_HEIGHT = 27              # 行の高さ

# =====================================================
# デフォルトの支出項目
# =====================================================
class DefaultColumns:
    """デフォルトの支出項目リスト"""
    ITEMS = [
        "日付",
        "交通",
        "外食",
        "食料品",
        "日用品・雑貨",
        "通販",
        "ゲーム関連",
        "サブスク",
        "家賃・宿泊",
        "光熱・通信",
        "医療・美容",
        "イベント"
    ]
    
    # 特殊列のインデックス
    DATE_COLUMN_INDEX = 0        # 日付列のインデックス
    INCOME_COLUMN_INDEX = 3      # 収入列のインデックス(まとめ行用)
    EXPENSE_COLUMN_INDEX = 5     # 支出列のインデックス(まとめ行用)

# =====================================================
# フォント設定
# =====================================================
class FontConfig:
    """フォント設定"""
    # 利用可能な日本語フォントリスト(優先順位順)
    JAPANESE_FONTS = ['Yu Gothic', 'Meiryo', 'MS Gothic', 'MS PGothic', 'MS UI Gothic']
    
    # フォールバックフォント
    FALLBACK_FONT = 'DejaVu Sans'
    
    # アプリケーション内フォント
    DEFAULT = ('Arial', 9)
    HEADING = ('Arial', 10, 'bold')
    TITLE = ('Segoe UI', 16, 'bold')
    BUTTON = ('Segoe UI', 9)
    BUTTON_LARGE = ('Segoe UI', 12, 'bold')

# =====================================================
# ダイアログ設定
# =====================================================
class DialogConfig:
    """ダイアログのデフォルトサイズ"""
    DEFAULT_WIDTH = 800
    DEFAULT_HEIGHT = 600
    MIN_SIZE_RATIO = 0.75  # 最小サイズ比率
    
    # 個別ダイアログサイズ
    TRANSACTION_WIDTH = 700
    TRANSACTION_HEIGHT = 500
    
    CHART_WIDTH = 900
    CHART_HEIGHT = 600
    
    COLUMN_EDIT_WIDTH = 300
    COLUMN_EDIT_HEIGHT = 120

# =====================================================
# ユーティリティ関数
# =====================================================
def get_current_year():
    """現在の年を取得"""
    return datetime.now().year

def get_current_month():
    """現在の月を取得"""
    return datetime.now().month

def parse_amount(amount_str):
    """
    金額文字列を整数に変換する
    
    Args:
        amount_str: 金額を表す文字列
        
    Returns:
        int: パースされた金額(失敗時は0)
    """
    if not amount_str:
        return 0
    try:
        clean_amount = str(amount_str).replace(',', '').replace('¥', '').strip()
        return int(clean_amount) if clean_amount else 0
    except ValueError:
        return 0
