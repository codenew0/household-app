# main.py
# pyinstaller.exe .\main.py --onefile --noconsole --icon=money.ico --name 家計.exe
"""
家計管理アプリケーションのエントリーポイント
"""
import tkinter as tk
from tkinter import messagebox
import sys
from ui.main_window import MainWindow 
from utils.font_utils import setup_japanese_font
import os
import tempfile

LOCK_FILE = os.path.join(tempfile.gettempdir(), "kakeibo_app.lock")

def check_single_instance():
    try:
        # 既存のロックファイルをチェック
        if os.path.exists(LOCK_FILE):
            with open(LOCK_FILE, 'r') as f:
                pid = int(f.read().strip())
            # そのPIDのプロセスが実際に生きているか確認
            import psutil
            if psutil.pid_exists(pid):
                return None  # 既に起動中
            # プロセスが死んでいれば古いロックファイルを無視
        
        # ロックファイルを作成
        with open(LOCK_FILE, 'w') as f:
            f.write(str(os.getpid()))
        return LOCK_FILE
    except Exception:
        return None

def release_lock(lock):
    try:
        if lock and os.path.exists(lock):
            os.remove(lock)
    except Exception:
        pass

def main():
    """アプリケーションのメインエントリーポイント"""
    
    # 二重起動チェック
    instance_lock = check_single_instance()
    if not instance_lock:
        # ルートウィンドウを一時的に作成してメッセージを表示
        root = tk.Tk()
        root.withdraw()  # ウィンドウ本体は表示しない
        messagebox.showwarning("警告", "アプリケーションは既に起動しています。")
        root.destroy()
        sys.exit(0)

    # 日本語フォントを設定
    setup_japanese_font()
    
    try:
        # Tkinterのルートウィンドウを作成
        root = tk.Tk()
        
        # メインウィンドウを初期化
        app = MainWindow(root)
        
        # メインループを開始
        root.mainloop()
        
    except Exception as e:
        # エラーが発生した場合は詳細を表示
        print(f"アプリケーション実行エラー: {e}")
        import traceback
        traceback.print_exc()
    finally:
        # アプリケーション終了時にソケットを閉じる
        if instance_lock and hasattr(instance_lock, 'close'):
            instance_lock.close()

if __name__ == "__main__":
    main()