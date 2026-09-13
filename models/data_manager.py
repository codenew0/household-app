# models/data_manager.py
"""
データ管理クラス
家計データの読み込み、保存、検索などを担当
"""
import json
import os
import shutil
import csv
import datetime
import calendar
import uuid
import tempfile
import copy
from models.transactions import normalize, validate, PENDING, describe, cash_amount
from models.productivity import ProductivityMixin
from config import DefaultColumns


CSV_COLUMNS = (
    "年", "月", "日", "列番号", "項目", "支払先", "金額", "メモ",
    "ポイント", "支払方法", "状態", "自動登録ID",
)
CSV_REQUIRED_COLUMNS = {"年", "月", "日", "列番号", "支払先", "金額", "メモ"}


class DataManager(ProductivityMixin):
    """家計データの管理を担当するクラス"""

    @staticmethod
    def _write_json(path, data):
        os.makedirs(os.path.dirname(path), exist_ok=True)
        fd, temporary = tempfile.mkstemp(dir=os.path.dirname(path), suffix='.tmp')
        try:
            with os.fdopen(fd, 'w', encoding='utf-8') as file:
                json.dump(data, file, ensure_ascii=False, indent=2)
                file.flush()
                os.fsync(file.fileno())
            os.replace(temporary, path)
        finally:
            if os.path.exists(temporary):
                os.unlink(temporary)
    
    def __init__(self):
        """データマネージャーの初期化"""
        self.data = {}  # 詳細データを格納する辞書 {key: [[partner, amount, detail], ...]}
        self.custom_columns = []  # カスタム項目リスト
        self.transaction_partners = set()  # 支払先の履歴
        self.subscriptions = []
        self.payment_methods = ['現金', 'クレジットカード', 'デビットカード', 'PayPay', '電子マネー', '口座振替']
        self.subscription_runs = set()
        self.budgets = {}
        self.load_warnings = []
        
        # ファイルパスの設定
        from config import JSON_DIR, SETTINGS_FILE, DATA_ROOT_DIR, APP_VERSION
        
        self.JSON_DIR = JSON_DIR
        self.SETTINGS_FILE = SETTINGS_FILE
        self.DATA_ROOT_DIR = DATA_ROOT_DIR
        self.APP_VERSION = APP_VERSION
        
        # 初期化時にデータフォルダの存在を確認・作成する
        self._ensure_data_directory()
    
    def _ensure_data_directory(self):
        """データ保存先のディレクトリが存在しない場合、作成する"""
        # 設定・バックアップ用ディレクトリ
        if self.JSON_DIR and not os.path.exists(self.JSON_DIR):
            try:
                os.makedirs(self.JSON_DIR, exist_ok=True)
            except OSError as e:
                print(f"データフォルダ作成エラー: {e}")
        
        # 新データディレクトリ
        if not os.path.exists(self.DATA_ROOT_DIR):
            try:
                os.makedirs(self.DATA_ROOT_DIR, exist_ok=True)
            except OSError as e:
                print(f"新データフォルダ作成エラー: {e}")
    
    def _get_data_file_path(self, year, month):
        """
        指定された年月のデータファイルパスを取得
        
        新フォーマット: data/2025/2025_01.json
        """
        year_dir = os.path.join(self.DATA_ROOT_DIR, str(year))
        
        if not os.path.exists(year_dir):
            os.makedirs(year_dir, exist_ok=True)
        
        return os.path.join(year_dir, f"{year}_{month:02d}.json")
    
    def _parse_key(self, dict_key):
        """
        キー文字列を解析して年月日と列インデックスを取得
        
        Args:
            dict_key: "年-月-日-列" 形式のキー
            
        Returns:
            tuple: (year, month, day, col_index) または None
        """
        try:
            parts = dict_key.split("-")
            if len(parts) == 4:
                return int(parts[0]), int(parts[1]), int(parts[2]), int(parts[3])
        except (ValueError, IndexError):
            pass
        return None
    
    @staticmethod
    def _to_new_format_transaction(col_index, transaction):
        """内部形式の1取引を月別JSON形式に変換する。"""
        transaction = normalize(transaction)
        return {
            "列目": str(col_index),
            "支払先": str(transaction[0]) if transaction[0] else "",
            "金額": str(transaction[1]) if transaction[1] else "",
            "詳細": transaction[2],
            "ポイント": transaction[3],
            "支払方法": transaction[4],
            "状態": transaction[5],
            "自動登録ID": transaction[6]
        }

    def _convert_file_to_internal_format(self, year, month, month_data):
        """
        月別JSONのデータをメモリ内部用の辞書に変換する。
        
        Args:
            year: 年
            month: 月
            month_data: 月別JSONの取引データ
            
        Returns:
            dict: 内部用の取引データ
        """
        internal_data = {}
        
        for day_key, transactions in month_data.items():
            # 日付ごとに列ごとにグループ化
            col_groups = {}
            
            for transaction in transactions:
                col_index = transaction.get("列目", "")
                if not col_index:
                    continue
                
                if col_index not in col_groups:
                    col_groups[col_index] = []
                
                col_groups[col_index].append([
                    transaction.get("支払先", ""),
                    transaction.get("金額", ""),
                    transaction.get("詳細", ""),
                    transaction.get("ポイント", ""),
                    transaction.get("支払方法", ""),
                    transaction.get("状態", ""),
                    transaction.get("自動登録ID", "")
                ])
            
            # 日付と列を含む内部キーで格納
            for col_index, trans_list in col_groups.items():
                internal_key = f"{year}-{month}-{day_key}-{col_index}"
                internal_data[internal_key] = trans_list
        
        return internal_data
    
    def load_data(self):
        """
        データファイルから家計データを読み込む
        月別JSONファイルから家計データを読み込む。
        """
        self.data = {}
        self._load_new_format_data()
    
    def _load_new_format_data(self):
        """月別JSONのデータを読み込む。"""
        if not os.path.exists(self.DATA_ROOT_DIR):
            return
        
        for year_name in os.listdir(self.DATA_ROOT_DIR):
            year_path = os.path.join(self.DATA_ROOT_DIR, year_name)
            if not os.path.isdir(year_path):
                continue
            
            try:
                year = int(year_name)
            except ValueError:
                continue
            
            for entry_name in os.listdir(year_path):
                entry_path = os.path.join(year_path, entry_name)
                
                # 月別ファイル: 2025_01.json
                if os.path.isfile(entry_path) and entry_name.endswith(".json"):
                    try:
                        base_name = os.path.splitext(entry_name)[0]  # "2025_01"
                        parts = base_name.split("_")
                        if len(parts) == 2:
                            month = int(parts[1])
                        else:
                            continue
                    except ValueError:
                        continue
                    
                    self._load_month_file(entry_path, year, month)

    def _load_month_file(self, data_file, year, month):
        """月別データファイルを読み込んでメモリに展開する"""
        try:
            with open(data_file, "r", encoding="utf-8") as f:
                month_data = json.load(f)
            
            internal_data = self._convert_file_to_internal_format(
                year, month, month_data.get("data", {})
            )
            self.data.update(internal_data)
            
            for data_list in internal_data.values():
                for row in data_list:
                    if len(row) > 0 and row[0] and str(row[0]).strip():
                        self.transaction_partners.add(str(row[0]).strip())
                
        except Exception as e:
            print(f"データ読み込みエラー ({year}/{month}): {e}")
            self.load_warnings.append(f'{year}年{month}月の明細を読み込めませんでした: {e}')
    
    def _save_month_data(self, year, month, month_data, replace=False):
        """
        指定された年月のデータを保存
        
        Args:
            year: 年
            month: 月
            month_data: 保存するデータ {day: [transactions]}
        """
        data_file = self._get_data_file_path(year, month)
        
        # 通常は日単位でマージ。列削除時は月全体を完全置換する。
        existing_data = {}
        original_file_data = None
        if os.path.exists(data_file) and not replace:
            try:
                with open(data_file, "r", encoding="utf-8") as f:
                    file_content = json.load(f)
                    existing_data = file_content.get("data", {})
                    original_file_data = file_content.get("data", {}).copy()
            except (OSError, ValueError):
                raise

        # データをマージ
        if replace:
            existing_data = month_data.copy()
        else:
            existing_data.update(month_data)

        # 保存
        save_data = {
            "version": self.APP_VERSION,
            "year": year,
            "month": month,
            "data": existing_data
        }

        # マージ後のデータがファイルの内容と同じなら書き込みをスキップ
        if original_file_data is not None and original_file_data == existing_data:
            return
        
        self._write_json(data_file, save_data)
    
    def save_data(self):
        """全データを保存"""
        self._save_all_month_data()

    def save_backup(self):
        """
        全データを月別JSONと同じ取引形式でbackupsフォルダに保存する。
        フォルダ構造: backups/2026/12/11/data_143000.json
        30日以上前の日付フォルダを自動削除する。
        アプリ終了時に呼び出される。
        """
        if not self.data:
            return
        
        now = datetime.datetime.now()
        backup_root = os.path.join(self.JSON_DIR, "backups")
        
        # 今日の日付フォルダを作成: backups/2026/12/11/
        date_dir = os.path.join(
            backup_root,
            now.strftime("%Y"),
            now.strftime("%m"),
            now.strftime("%d")
        )
        try:
            os.makedirs(date_dir, exist_ok=True)
        except OSError as e:
            print(f"バックアップフォルダ作成エラー: {e}")
            return
        
        # 時刻付きファイル名で保存
        backup_file = os.path.join(date_dir, f"data_{now.strftime('%H%M%S')}.json")
        grouped_data = {
            f"{year}-{month}": month_data
            for (year, month), month_data in self._group_data_by_month().items()
        }

        backup_data = {
            "version": self.APP_VERSION,
            "data": grouped_data,
            "subscriptions": self.subscriptions,
            "subscription_runs": sorted(self.subscription_runs)
        }
        
        try:
            self._write_json(backup_file, backup_data)
            print(f"バックアップ保存完了: {backup_file}")
        except Exception as e:
            print(f"バックアップ保存エラー: {e}")
            return
        
        # 30日以上前の日付フォルダを自動削除
        cutoff = now - datetime.timedelta(days=30)
        self._remove_expired_backups(backup_root, cutoff)

    @staticmethod
    def _remove_expired_backups(backup_root, cutoff):
        """期限切れの日別バックアップと、空になった親フォルダを削除する。"""
        try:
            for year_name in os.listdir(backup_root):
                year_path = os.path.join(backup_root, year_name)
                if not os.path.isdir(year_path) or not year_name.isdigit():
                    continue
                for month_name in os.listdir(year_path):
                    month_path = os.path.join(year_path, month_name)
                    if not os.path.isdir(month_path) or not month_name.isdigit():
                        continue
                    for day_name in os.listdir(month_path):
                        day_path = os.path.join(month_path, day_name)
                        if not os.path.isdir(day_path) or not day_name.isdigit():
                            continue
                        try:
                            folder_date = datetime.datetime(
                                int(year_name), int(month_name), int(day_name))
                            if folder_date < cutoff:
                                shutil.rmtree(day_path)
                                print(f"古いバックアップを削除: {day_path}")
                        except (ValueError, OSError) as e:
                            print(f"バックアップ削除エラー: {day_path}: {e}")
                    # 月フォルダが空になったら削除
                    if not os.listdir(month_path):
                        os.rmdir(month_path)
                # 年フォルダが空になったら削除
                if not os.listdir(year_path):
                    os.rmdir(year_path)
        except OSError as e:
            print(f"古いバックアップ削除エラー: {e}")

    def _group_data_by_month(self):
        """メモリ上の明細を、保存形式の年月・日単位にまとめる。"""
        grouped = {}
        for key, transactions in self.data.items():
            parsed = self._parse_key(key)
            if not parsed:
                continue
            year, month, day, col_index = parsed
            rows = grouped.setdefault((year, month), {}).setdefault(str(day), [])
            for transaction in transactions:
                if len(transaction) >= 3:
                    rows.append(self._to_new_format_transaction(col_index, transaction))
        return grouped

    def _save_all_month_data(self):
        """全データを年月ごとにグループ化して保存する。"""
        for (year, month), month_data in self._group_data_by_month().items():
            self._save_month_data(year, month, month_data)
    
    def save_transaction(self, dict_key, create_snapshot=True):
        """
        指定キーに関連する日のデータを保存する。
        同じ日の全列のデータを収集してから保存する。

        Args:
            dict_key: データのキー(年-月-日-列インデックス)
        """
        parsed = self._parse_key(dict_key)
        if not parsed:
            return

        year, month, day, _ = parsed
        if create_snapshot:
            self.snapshot('明細変更の直前', from_disk=True)
        self._save_day_data(year, month, day)

    def save_transactions(self, dict_keys, create_snapshot=True):
        """
        複数キーに関連するデータをまとめて保存する。
        同じ年月日のデータは1回だけ保存する。

        Args:
            dict_keys: データのキーのリスト
        """
        dict_keys = list(dict_keys)
        if dict_keys and create_snapshot:
            self.snapshot('明細変更の直前', from_disk=True)
        # 年月日でグループ化して重複保存を防ぐ
        saved_days = set()
        for dict_key in dict_keys:
            parsed = self._parse_key(dict_key)
            if not parsed:
                continue
            year, month, day, _ = parsed
            day_tuple = (year, month, day)
            if day_tuple not in saved_days:
                saved_days.add(day_tuple)
                self._save_day_data(year, month, day)

    def _save_day_data(self, year, month, day):
        """
        指定された日の全列データを収集して月データファイルに保存する。

        Args:
            year: 年
            month: 月
            day: 日
        """
        day_key = str(day)
        month_data = {day_key: []}

        # 同じ年月日の全列のデータを収集
        for key, transactions in self.data.items():
            key_parsed = self._parse_key(key)
            if not key_parsed:
                continue
            k_year, k_month, k_day, k_col = key_parsed
            if k_year == year and k_month == month and k_day == day:
                for transaction in transactions:
                    if len(transaction) >= 3:
                        month_data[day_key].append(
                            self._to_new_format_transaction(k_col, transaction)
                        )

        self._save_month_data(year, month, month_data)
    
    def load_settings(self):
        """設定ファイルから設定を読み込む"""
        if os.path.exists(self.SETTINGS_FILE):
            try:
                with open(self.SETTINGS_FILE, "r", encoding="utf-8") as f:
                    settings = json.load(f)
                    self.custom_columns = settings.get("custom_columns", [])
                    self.transaction_partners = set(settings.get("transaction_partners", []))
                    self.subscriptions = settings.get("subscriptions", [])
                    self.payment_methods = settings.get('payment_methods', self.payment_methods)
                    self.subscription_runs = set(settings.get("subscription_runs", []))
                    self.budgets = settings.get('budgets', {})
            except Exception as e:
                print(f"設定読み込みエラー: {e}")
                timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
                backup = self.SETTINGS_FILE + f'.corrupt_{timestamp}'
                try:
                    shutil.copy2(self.SETTINGS_FILE, backup)
                    detail = f'破損した設定は {backup} に保護しました。'
                except OSError as backup_error:
                    detail = f'破損設定の保護にも失敗しました: {backup_error}'
                self.load_warnings.append(f'設定ファイルを読み込めませんでした。{detail}')
    
    def save_settings(self):
        """設定をファイルに保存する"""
        settings = {
            "custom_columns": self.custom_columns,
            "transaction_partners": list(self.transaction_partners),
            "subscriptions": self.subscriptions,
            "payment_methods": self.payment_methods,
            "subscription_runs": sorted(self.subscription_runs),
            "budgets": self.budgets
        }
        self._write_json(self.SETTINGS_FILE, settings)
    
    def get_transaction_data(self, dict_key):
        """指定されたキーの取引データを取得"""
        return self.data.get(dict_key, [])

    def add_payment_method(self, name):
        name = name.strip()
        if not name:
            raise ValueError('支払方法の名前を入力してください。')
        if name in self.payment_methods:
            raise ValueError('同じ支払方法が登録されています。')
        previous = self.payment_methods[:]
        self.payment_methods.append(name)
        try:
            self.save_settings()
        except OSError:
            self.payment_methods = previous
            raise

    def remove_payment_method(self, name):
        previous = self.payment_methods[:]
        self.payment_methods = [method for method in previous if method != name]
        try:
            self.save_settings()
        except OSError:
            self.payment_methods = previous
            raise
    
    def set_transaction_data(self, dict_key, data_list):
        """
        指定されたキーに取引データを設定する（メモリのみ）
        ファイル保存は呼び出し側で明示的に行う。

        Args:
            dict_key: データのキー(年-月-日-列インデックス)
            data_list: 設定する取引データのリスト
        """
        if data_list:
            self.data[dict_key] = data_list
        elif dict_key in self.data:
            del self.data[dict_key]

    def delete_transaction_data(self, dict_key):
        """
        指定されたキーの取引データを削除する（メモリのみ）
        ファイル保存は呼び出し側で明示的に行う。

        Args:
            dict_key: データのキー
        """
        if dict_key in self.data:
            del self.data[dict_key]
    
    def get_transaction_partners_list(self):
        """支払先の履歴をソート済みリストで取得"""
        return sorted(self.transaction_partners)
    
    def add_custom_column(self, column_name):
        """カスタム項目を追加"""
        if column_name and column_name not in self.custom_columns:
            previous = self.custom_columns[:]
            self.custom_columns.append(column_name)
            try:
                self.save_settings()
            except OSError:
                self.custom_columns = previous
                raise
            return True
        return False
    
    def edit_custom_column(self, old_name, new_name):
        """カスタム項目名を編集"""
        if old_name in self.custom_columns and new_name not in self.custom_columns:
            previous_columns = self.custom_columns[:]
            previous_subscriptions = copy.deepcopy(self.subscriptions)
            index = self.custom_columns.index(old_name)
            self.custom_columns[index] = new_name
            for item in self.subscriptions:
                if item.get('category') == old_name:
                    item['category'] = new_name
            try:
                self.save_settings()
            except OSError:
                self.custom_columns = previous_columns
                self.subscriptions = previous_subscriptions
                raise
            return True
        return False
    
    def remove_custom_column(self, column_name):
        """分類・関連データ・設定をまとめて削除し、失敗時は元へ戻す。"""
        if column_name not in self.custom_columns:
            return False
        column = len(DefaultColumns.ITEMS) + self.custom_columns.index(column_name)
        previous_data = copy.deepcopy(self.data)
        previous_columns = self.custom_columns[:]
        previous_budgets = copy.deepcopy(self.budgets)
        previous_subscriptions = copy.deepcopy(self.subscriptions)
        affected_months = {parsed[:2] for key in self.data if (parsed := self._parse_key(key)) and parsed[3] >= column}
        self.snapshot('分類削除の直前')
        try:
            self.budgets = {period: {str(int(key) - 1 if int(key) > column else int(key)): value
                            for key, value in values.items() if int(key) != column}
                            for period, values in self.budgets.items()}
            for item in self.subscriptions:
                if item.get('category') == column_name:
                    kind = item.get('kind', 'サブスクリプション')
                    item['category'] = self.RECURRING_DEFAULT_CATEGORIES.get(kind, 'サブスク')
            self.custom_columns.remove(column_name)
            new_data = {}
            for key, transactions in self.data.items():
                parsed = self._parse_key(key)
                if not parsed:
                    new_data[key] = transactions
                    continue
                year, month, day, key_col = parsed
                if key_col == column:
                    continue
                target = f'{year}-{month}-{day}-{key_col - 1}' if key_col > column else key
                new_data[target] = transactions
            self.data = new_data
            for year, month in affected_months:
                self._save_complete_month(year, month)
            self.save_settings()
            return True
        except Exception:
            self.data = previous_data
            self.custom_columns = previous_columns
            self.budgets = previous_budgets
            self.subscriptions = previous_subscriptions
            try:
                for year, month in affected_months:
                    self._save_complete_month(year, month)
                self.save_settings()
            except Exception:
                pass
            raise
    
    def search_transactions(self, search_text):
        """取引データを検索"""
        results = []
        search_text_lower = search_text.lower()
        
        for dict_key, data_list in self.data.items():
            try:
                parts = dict_key.split("-")
                if len(parts) == 4:
                    year, month, day, col_index = map(int, parts)
                    
                    for row in data_list:
                        if len(row) >= 3:
                            values = normalize(row)
                            partner = values[0].strip()
                            amount = values[1].strip()
                            detail = describe(values)
                            searchable = ' '.join((partner, amount, values[2], values[3],
                                                   values[4], values[5], detail))
                            
                            if search_text_lower in searchable.lower():
                                results.append({
                                    'year': year,
                                    'month': month,
                                    'day': day,
                                    'col_index': col_index,
                                    'partner': partner,
                                    'amount': str(cash_amount(row)),
                                    'gross_amount': values[1],
                                    'points': values[3],
                                    'payment_method': values[4],
                                    'status': values[5],
                                    'memo': values[2],
                                    'detail': detail
                                })
            except (ValueError, IndexError):
                continue
        
        return results

    def export_csv(self, file_path, start, end, all_columns):
        """指定年月範囲の取引をCSVへ出力する。"""
        rows = []
        for dict_key, data_list in self.data.items():
            parsed = self._parse_key(dict_key)
            if not parsed:
                continue
            year, month, day, col_index = parsed
            if not start <= (year, month) <= end:
                continue
            if day == 0:
                column_name = "収入"
            elif 0 <= col_index < len(all_columns):
                column_name = all_columns[col_index]
            else:
                column_name = f"列{col_index}"
            for transaction in data_list:
                if len(transaction) < 3:
                    continue
                rows.append([
                    year, month, day, col_index, column_name,
                    *normalize(transaction)
                ])

        rows.sort(key=lambda row: (int(row[0]), int(row[1]), int(row[2]), int(row[3])))
        with open(file_path, "w", encoding="utf-8-sig", newline="") as file:
            writer = csv.writer(file)
            writer.writerow(CSV_COLUMNS)
            writer.writerows(rows)
        return len(rows)

    RECURRING_PAYMENT_TYPES = ('サブスクリプション', '家賃', '公共料金', '分割払い')
    RECURRING_DEFAULT_CATEGORIES = {
        'サブスクリプション': 'サブスク',
        '家賃': '家賃・宿泊',
        '公共料金': '光熱・通信',
        '分割払い': '通販',
    }

    @staticmethod
    def _add_months(value, months):
        """月初の日付を指定月数だけ進める。"""
        zero_based = value.year * 12 + value.month - 1 + months
        return datetime.date(zero_based // 12, zero_based % 12 + 1, 1)

    def recurring_payment_values(self, item):
        """旧サブスク定義にも既定値を補い、共通形式で返す。"""
        kind = item.get('kind', 'サブスクリプション')
        return {
            **item,
            'kind': kind,
            'category': item.get('category') or self.RECURRING_DEFAULT_CATEGORIES.get(kind, 'サブスク'),
            'interval_months': max(1, int(item.get('interval_months', 1))),
            'payment_count': max(0, int(item.get('payment_count') or 0)),
            'enabled': item.get('enabled', True),
            'skip_through': item.get('skip_through', ''),
        }

    def iter_recurring_payment_dates(self, item, through=None):
        """開始日以降の支払日を (回数, 日付) で列挙する。"""
        item = self.recurring_payment_values(item)
        start = datetime.date.fromisoformat(item['start'])
        interval = item['interval_months']
        month = datetime.date(start.year, start.month, 1)
        due = datetime.date(month.year, month.month,
                            min(int(item['day']), calendar.monthrange(month.year, month.month)[1]))
        if due < start:
            month = self._add_months(month, interval)
        number = 1
        while not item['payment_count'] or number <= item['payment_count']:
            day = min(int(item['day']), calendar.monthrange(month.year, month.month)[1])
            due = datetime.date(month.year, month.month, day)
            if through is not None and due > through:
                break
            yield number, due
            number += 1
            month = self._add_months(month, interval)

    def save_subscription(self, name, partner, amount, day, method, start_date,
                          subscription_id=None, enabled=True, kind='サブスクリプション',
                          category=None, interval_months=1, payment_count=None):
        """定期支払いの定義を登録・更新する。既存明細には影響しない。"""
        if not name.strip() or not partner.strip():
            raise ValueError("名前と支払先を入力してください。")
        if kind not in self.RECURRING_PAYMENT_TYPES:
            raise ValueError("種類を選択してください。")
        day = int(day)
        if not 1 <= day <= 31:
            raise ValueError("支払日は1〜31で入力してください。")
        interval_months = int(interval_months)
        if not 1 <= interval_months <= 120:
            raise ValueError("支払間隔は1〜120か月で入力してください。")
        payment_count = int(payment_count or 0)
        if payment_count < 0 or payment_count > 1200:
            raise ValueError("支払回数は1〜1200回、または空欄で入力してください。")
        if kind == '分割払い' and payment_count < 1:
            raise ValueError("分割払いは支払回数を入力してください。")
        start = datetime.date.fromisoformat(start_date)
        amount = validate([partner, amount, ""])[1]
        if int(amount) < 0:
            raise ValueError("金額は0以上で入力してください。")
        category = (category or self.RECURRING_DEFAULT_CATEGORIES[kind]).strip()
        if category not in DefaultColumns.ITEMS + self.custom_columns or category == '日付':
            raise ValueError("登録先の分類を選択してください。")
        existing = next((item for item in self.subscriptions
                         if item.get('id') == subscription_id), {})
        record = dict(id=subscription_id or uuid.uuid4().hex, name=name.strip(),
                      partner=partner.strip(), amount=amount, day=day,
                      method=method.strip(), start=start.isoformat(), enabled=enabled,
                      kind=kind, category=category, interval_months=interval_months,
                      payment_count=payment_count,
                      skip_through=existing.get('skip_through', ''))
        previous = self.subscriptions[:]
        self.subscriptions = [s for s in previous if s['id'] != record['id']] + [record]
        try:
            self.save_settings()
        except Exception:
            self.subscriptions = previous
            raise
        return record

    def generate_due_subscriptions(self, today=None):
        """支払日到来分を1回だけ追加。削除した自動明細は再生成しない。"""
        today = today or datetime.date.today()
        previous_data = copy.deepcopy(self.data)
        previous_runs = self.subscription_runs.copy()
        previous_subscriptions = copy.deepcopy(self.subscriptions)
        schedule_changed = False
        seen = self.subscription_runs | {
            str(row[6]) for rows in self.data.values() for row in rows
            if len(row) > 6 and row[6]
        }
        affected = set()
        generated = []
        columns = DefaultColumns.ITEMS + self.custom_columns
        for subscription in self.subscriptions:
            raw_subscription = subscription
            subscription = self.recurring_payment_values(raw_subscription)
            if not subscription['enabled']:
                value = today.isoformat()
                if raw_subscription.get('skip_through', '') < value:
                    raw_subscription['skip_through'] = value
                    schedule_changed = True
                continue
            category = subscription['category']
            if category not in columns:
                category = self.RECURRING_DEFAULT_CATEGORIES.get(subscription['kind'], 'サブスク')
            column = columns.index(category)
            skip_through = (datetime.date.fromisoformat(subscription['skip_through'])
                            if subscription.get('skip_through') else None)
            for number, due in self.iter_recurring_payment_dates(subscription, today):
                if skip_through and due <= skip_through:
                    continue
                occurrence = f"{subscription['id']}:{due.year}-{due.month:02d}"
                if occurrence in seen:
                    continue
                key = f"{due.year}-{due.month}-{due.day}-{column}"
                memo = subscription['name']
                if subscription['payment_count']:
                    memo += f" ({number}/{subscription['payment_count']}回)"
                row = [subscription['partner'], subscription['amount'], memo,
                       "0", subscription['method'], PENDING, occurrence]
                self.data.setdefault(key, []).append(row)
                affected.add(key)
                generated.append(occurrence)
        # 月ファイルを先に保存。設定保存前の中断でも明細IDで重複を防げる。
        try:
            self.save_transactions(affected)
        except Exception:
            self.data = previous_data
            self.subscriptions = previous_subscriptions
            raise
        self.subscription_runs.update(seen)
        self.subscription_runs.update(generated)
        if self.subscription_runs != previous_runs or schedule_changed:
            try:
                self.save_settings()
            except Exception:
                self.subscription_runs = previous_runs
                self.subscriptions = previous_subscriptions
                raise
        return len(generated)

    def _read_csv(self, file_path, start, end, column_count):
        """CSVを検証し、指定期間内の明細を内部形式で返す。"""
        imported = {}
        row_count = 0
        with open(file_path, "r", encoding="utf-8-sig", newline="") as file:
            reader = csv.DictReader(file)
            if not reader.fieldnames or not CSV_REQUIRED_COLUMNS.issubset(reader.fieldnames):
                missing = "、".join(sorted(
                    CSV_REQUIRED_COLUMNS - set(reader.fieldnames or [])))
                raise ValueError(f"CSVの必須列がありません: {missing}")

            for line_number, row in enumerate(reader, start=2):
                try:
                    year = int(row["年"])
                    month = int(row["月"])
                    day = int(row["日"])
                    col_index = int(row["列番号"])
                except (TypeError, ValueError):
                    raise ValueError(f"{line_number}行目: 年・月・日・列番号は数値で指定してください。")

                if not start <= (year, month) <= end:
                    continue
                if not 1 <= month <= 12:
                    raise ValueError(f"{line_number}行目: 月は1～12で指定してください。")
                if day == 0:
                    if col_index != DefaultColumns.INCOME_COLUMN_INDEX:
                        raise ValueError(
                            f"{line_number}行目: 収入データの列番号は"
                            f"{DefaultColumns.INCOME_COLUMN_INDEX}です。")
                else:
                    try:
                        datetime.date(year, month, day)
                    except ValueError:
                        raise ValueError(f"{line_number}行目: 存在しない日付です。")
                    if not 1 <= col_index < column_count:
                        raise ValueError(f"{line_number}行目: 列番号{col_index}は現在の項目に存在しません。")

                key = f"{year}-{month}-{day}-{col_index}"
                imported.setdefault(key, []).append(validate([
                    row.get("支払先", ""), row.get("金額", ""), row.get("メモ", ""),
                    row.get("ポイント", ""), row.get("支払方法", ""),
                    row.get("状態", ""), row.get("自動登録ID", "")
                ]))
                row_count += 1
        return imported, row_count

    def import_csv(self, file_path, start, end, mode, column_count, preview_only=False):
        """
        CSVを指定期間に追加または上書きする。

        replaceは指定期間全体をCSVに含まれる取引で置き換える。
        """
        if mode not in ("append", "replace"):
            raise ValueError("インポート方式が不正です。")

        imported, row_count = self._read_csv(file_path, start, end, column_count)

        if preview_only:
            return imported
        self.snapshot('CSV取り込みの直前')
        previous_data = copy.deepcopy(self.data)
        previous_partners = self.transaction_partners.copy()
        affected_months = set(self._iter_months(start, end)) if mode == "replace" else {
            self._parse_key(key)[:2] for key in imported
        }
        if mode == "replace":
            self.data = {
                key: value for key, value in self.data.items()
                if not (self._parse_key(key) and start <= self._parse_key(key)[:2] <= end)
            }

        for key, rows in imported.items():
            if mode == "append" and key in self.data:
                self.data[key] = list(self.data[key]) + rows
            else:
                self.data[key] = rows
            for transaction in rows:
                partner = transaction[0]
                if partner and str(partner).strip():
                    self.transaction_partners.add(str(partner).strip())

        try:
            for year, month in affected_months:
                self._save_complete_month(year, month)
            self.save_settings()
        except Exception:
            self.data = previous_data
            self.transaction_partners = previous_partners
            try:
                for year, month in affected_months:
                    self._save_complete_month(year, month)
                self.save_settings()
            except Exception:
                pass
            raise
        return row_count

    @staticmethod
    def _iter_months(start, end):
        """開始年月から終了年月までを列挙する。"""
        year, month = start
        while (year, month) <= end:
            yield year, month
            month += 1
            if month == 13:
                year += 1
                month = 1

    def _save_complete_month(self, year, month):
        """メモリ上の月データで月別JSONを完全上書きする。"""
        month_data = self._group_data_by_month().get((year, month), {})
        self._save_month_data(year, month, month_data, replace=True)
