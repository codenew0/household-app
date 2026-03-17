# models/data_manager.py
"""
データ管理クラス
家計データの読み込み、保存、検索などを担当
"""
import json
import os
import shutil


class DataManager:
    """家計データの管理を担当するクラス"""
    
    def __init__(self):
        """データマネージャーの初期化"""
        self.data = {}  # 詳細データを格納する辞書 {key: [[partner, amount, detail], ...]}
        self.custom_columns = []  # カスタム項目リスト
        self.transaction_partners = set()  # 支払先の履歴
        
        # ファイルパスの設定
        from config import JSON_DIR, SETTINGS_FILE, DATA_FILE, APP_VERSION
        
        self.JSON_DIR = JSON_DIR
        self.SETTINGS_FILE = SETTINGS_FILE
        self.DATA_FILE = DATA_FILE  # 旧フォーマット
        self.DATA_FILE_OLD = os.path.join(JSON_DIR, "data_1.json")  # 旧フォーマットバックアップ
        self.DATA_ROOT_DIR = os.path.join(JSON_DIR, "data")  # 新フォーマットのルート
        self.APP_VERSION = APP_VERSION
        
        # 初期化時にデータフォルダの存在を確認・作成する
        self._ensure_data_directory()
    
    def _ensure_data_directory(self):
        """データ保存先のディレクトリが存在しない場合、作成する"""
        # 旧データディレクトリ
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
    
    def _convert_old_to_new_format(self, old_data):
        """
        旧フォーマットのデータを新フォーマットに変換
        
        旧: "2026-1-14-7": [["apple", "150", "iCloud"]]
        新: "2026-1-14": [{"列目": "7", "支払先": "apple", "金額": "150", "詳細": "iCloud"}]
        
        Args:
            old_data: 旧フォーマットのデータ辞書
            
        Returns:
            dict: 新フォーマットのデータ辞書 {year-month: {day: [transactions]}}
        """
        converted = {}
        
        for key, transactions in old_data.items():
            parsed = self._parse_key(key)
            if not parsed:
                continue
            
            year, month, day, col_index = parsed
            year_month_key = f"{year}-{month}"
            day_key = str(day)
            
            if year_month_key not in converted:
                converted[year_month_key] = {}
            
            if day_key not in converted[year_month_key]:
                converted[year_month_key][day_key] = []
            
            # 各取引を新フォーマットに変換
            for transaction in transactions:
                if len(transaction) >= 3:
                    new_transaction = {
                        "列目": str(col_index),
                        "支払先": str(transaction[0]) if transaction[0] else "",
                        "金額": str(transaction[1]) if transaction[1] else "",
                        "詳細": str(transaction[2]) if transaction[2] else ""
                    }
                    converted[year_month_key][day_key].append(new_transaction)
        
        return converted
    
    def _convert_new_to_old_format(self, year, month, new_data):
        """
        新フォーマットのデータを旧フォーマット(メモリ内部用)に変換
        
        Args:
            year: 年
            month: 月
            new_data: 新フォーマットのデータ
            
        Returns:
            dict: 旧フォーマットのデータ
        """
        old_format = {}
        
        for day_key, transactions in new_data.items():
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
                    transaction.get("詳細", "")
                ])
            
            # 旧フォーマットのキーで格納
            for col_index, trans_list in col_groups.items():
                old_key = f"{year}-{month}-{day_key}-{col_index}"
                old_format[old_key] = trans_list
        
        return old_format
    
    def load_data(self):
        """
        データファイルから家計データを読み込む
        1. 新フォーマット(data/年/月/data.json)から読み込み
        2. 旧フォーマット(data.json)があれば読み込んで変換・保存
        3. data_1.jsonがあれば読み込み
        """
        self.data = {}
        
        # 新フォーマットのデータを読み込み
        self._load_new_format_data()
        
        # 旧フォーマットのデータがあれば読み込み・変換
        if os.path.exists(self.DATA_FILE):
            self._migrate_old_format_data()
        
        # data_1.jsonからも読み込み
        if os.path.exists(self.DATA_FILE_OLD):
            self._load_old_backup_data()
    
    def _load_new_format_data(self):
        """新フォーマットのデータを読み込み"""
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
                
                # 新フォーマット: 2025_01.json のようなファイル
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
                
                # 旧フォルダ形式: 01/data.json (互換性のため)
                elif os.path.isdir(entry_path):
                    try:
                        month = int(entry_name)
                    except ValueError:
                        continue
                    
                    data_file = os.path.join(entry_path, "data.json")
                    if os.path.exists(data_file):
                        self._load_month_file(data_file, year, month)

    def _load_month_file(self, data_file, year, month):
        """月別データファイルを読み込んでメモリに展開する"""
        try:
            with open(data_file, "r", encoding="utf-8") as f:
                month_data = json.load(f)
            
            old_format = self._convert_new_to_old_format(year, month, month_data.get("data", {}))
            self.data.update(old_format)
            
            for data_list in old_format.values():
                for row in data_list:
                    if len(row) > 0 and row[0] and str(row[0]).strip():
                        self.transaction_partners.add(str(row[0]).strip())
                
        except Exception as e:
            print(f"データ読み込みエラー ({year}/{month}): {e}")
    
    def _migrate_old_format_data(self):
        """旧フォーマットのデータを新フォーマットに移行"""
        try:
            with open(self.DATA_FILE, "r", encoding="utf-8") as f:
                old_data = json.load(f)
            
            # バージョン情報がある場合
            if "version" in old_data:
                data_dict = old_data.get("data", {})
            else:
                data_dict = old_data.get("data", {})
            
            # 新フォーマットに変換
            converted = self._convert_old_to_new_format(data_dict)
            
            # 年月ごとに保存
            for year_month_key, month_data in converted.items():
                try:
                    year, month = map(int, year_month_key.split("-"))
                    self._save_month_data(year, month, month_data)
                except Exception as e:
                    print(f"移行エラー ({year_month_key}): {e}")
            
            # 旧データをバックアップとしてリネーム
            if not os.path.exists(self.DATA_FILE_OLD):
                shutil.copy2(self.DATA_FILE, self.DATA_FILE_OLD)
            
            # 旧データファイルを削除(オプション)
            # os.remove(self.DATA_FILE)
            
            print("旧フォーマットから新フォーマットへの移行が完了しました")
            
        except Exception as e:
            print(f"旧データ移行エラー: {e}")
    
    def _load_old_backup_data(self):
        """data_1.jsonからデータを読み込み"""
        try:
            with open(self.DATA_FILE_OLD, "r", encoding="utf-8") as f:
                old_data = json.load(f)
            
            if "version" in old_data:
                data_dict = old_data.get("data", {})
            else:
                data_dict = old_data.get("data", {})
            
            # 既存データとマージ(既存優先) + 支払先の抽出を同時に実行
            for key, value in data_dict.items():
                if key not in self.data:
                    self.data[key] = value
                    
                    # マージ時に支払先を抽出（効率化）
                    for row in value:
                        if len(row) > 0 and row[0] and str(row[0]).strip():
                            self.transaction_partners.add(str(row[0]).strip())
            
        except Exception as e:
            print(f"data_1.json読み込みエラー: {e}")
    
    def _save_month_data(self, year, month, month_data):
        """
        指定された年月のデータを保存
        
        Args:
            year: 年
            month: 月
            month_data: 保存するデータ {day: [transactions]}
        """
        data_file = self._get_data_file_path(year, month)
        
        # 既存データがあれば読み込んでマージ
        existing_data = {}
        existing_file_content = None
        if os.path.exists(data_file):
            try:
                with open(data_file, "r", encoding="utf-8") as f:
                    existing_file_content = json.load(f)
                    existing_data = existing_file_content.get("data", {})
            except:
                pass
        
        # データをマージ
        existing_data.update(month_data)
        
        # 保存
        save_data = {
            "version": self.APP_VERSION,
            "year": year,
            "month": month,
            "data": existing_data
        }
        
        # ファイルが既に存在し、内容が同じなら書き込みをスキップ
        if existing_file_content is not None and existing_file_content == save_data:
            return
        
        try:
            with open(data_file, "w", encoding="utf-8") as f:
                json.dump(save_data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"データ保存エラー ({year}/{month}): {e}")
    
    def save_data(self):
        """全データを保存"""
        self._save_all_month_data()

    def save_backup(self):
        """
        全データを旧フォーマットのJSONファイルとしてbackupsフォルダに保存する。
        フォルダ構造: backups/2026/12/11/data_143000.json
        30日以上前の日付フォルダを自動削除する。
        アプリ終了時に呼び出される。
        """
        if not self.data:
            return
        
        from datetime import datetime, timedelta
        
        now = datetime.now()
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
        backup_data = {
            "version": self.APP_VERSION,
            "data": self.data
        }
        
        try:
            with open(backup_file, "w", encoding="utf-8") as f:
                json.dump(backup_data, f, ensure_ascii=False, indent=2)
            print(f"バックアップ保存完了: {backup_file}")
        except Exception as e:
            print(f"バックアップ保存エラー: {e}")
            return
        
        # 30日以上前の日付フォルダを自動削除
        cutoff = now - timedelta(days=30)
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
                            folder_date = datetime(int(year_name), int(month_name), int(day_name))
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

    def _save_all_month_data(self):
        """全データを年月ごとにグループ化して保存"""
        # 年月ごとにグループ化
        year_month_data = {}
        
        for key, transactions in self.data.items():
            parsed = self._parse_key(key)
            if not parsed:
                continue
            
            year, month, day, col_index = parsed
            year_month_key = (year, month)
            
            if year_month_key not in year_month_data:
                year_month_data[year_month_key] = {}
            
            day_key = str(day)
            if day_key not in year_month_data[year_month_key]:
                year_month_data[year_month_key][day_key] = []
            
            # 新フォーマットに変換
            for transaction in transactions:
                if len(transaction) >= 3:
                    new_transaction = {
                        "列目": str(col_index),
                        "支払先": str(transaction[0]) if transaction[0] else "",
                        "金額": str(transaction[1]) if transaction[1] else "",
                        "詳細": str(transaction[2]) if transaction[2] else ""
                    }
                    year_month_data[year_month_key][day_key].append(new_transaction)
        
        # 年月ごとに保存
        for (year, month), month_data in year_month_data.items():
            self._save_month_data(year, month, month_data)
    
    def auto_save_transaction(self, dict_key):
        """
        特定の取引データを即座に保存
        
        Args:
            dict_key: データのキー(年-月-日-列インデックス)
        """
        parsed = self._parse_key(dict_key)
        if not parsed:
            return
        
        year, month, day, col_index = parsed
        
        # この年月の全データを収集
        month_data = {}
        day_key = str(day)
        
        # 現在の日のデータを新フォーマットに変換
        if dict_key in self.data:
            transactions = self.data[dict_key]
            
            if day_key not in month_data:
                month_data[day_key] = []
            
            for transaction in transactions:
                if len(transaction) >= 3:
                    new_transaction = {
                        "列目": str(col_index),
                        "支払先": str(transaction[0]) if transaction[0] else "",
                        "金額": str(transaction[1]) if transaction[1] else "",
                        "詳細": str(transaction[2]) if transaction[2] else ""
                    }
                    month_data[day_key].append(new_transaction)
        else:
            # データが削除された場合、空で保存(既存データから削除)
            month_data[day_key] = []
        
        # 保存
        self._save_month_data(year, month, month_data)
    
    def load_settings(self):
        """設定ファイルから設定を読み込む"""
        if os.path.exists(self.SETTINGS_FILE):
            try:
                with open(self.SETTINGS_FILE, "r", encoding="utf-8") as f:
                    settings = json.load(f)
                    self.custom_columns = settings.get("custom_columns", [])
                    self.transaction_partners = set(settings.get("transaction_partners", []))
            except Exception as e:
                print(f"設定読み込みエラー: {e}")
    
    def save_settings(self):
        """設定をファイルに保存する"""
        settings = {
            "custom_columns": self.custom_columns,
            "transaction_partners": list(self.transaction_partners)
        }
        try:
            with open(self.SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"設定保存エラー: {e}")
    
    def get_transaction_data(self, dict_key):
        """指定されたキーの取引データを取得"""
        return self.data.get(dict_key, [])
    
    def set_transaction_data(self, dict_key, data_list):
        """
        指定されたキーに取引データを設定し、即座に保存
        
        Args:
            dict_key: データのキー(年-月-日-列インデックス)
            data_list: 設定する取引データのリスト
        """
        if data_list:
            self.data[dict_key] = data_list
        elif dict_key in self.data:
            del self.data[dict_key]
        
        # 即座に保存
        self.auto_save_transaction(dict_key)
    
    def delete_transaction_data(self, dict_key):
        """
        指定されたキーの取引データを削除し、即座に保存
        
        Args:
            dict_key: データのキー
        """
        if dict_key in self.data:
            del self.data[dict_key]
        
        # 即座に保存
        self.auto_save_transaction(dict_key)
    
    def add_transaction_partner(self, partner):
        """支払先を履歴に追加"""
        if partner and partner.strip():
            self.transaction_partners.add(partner.strip())
            # 即座に設定を保存
            self.save_settings()
    
    def get_transaction_partners_list(self):
        """支払先の履歴をソート済みリストで取得"""
        return sorted(list(self.transaction_partners))
    
    def add_custom_column(self, column_name):
        """カスタム項目を追加"""
        if column_name and column_name not in self.custom_columns:
            self.custom_columns.append(column_name)
            self.save_settings()  # 即座に保存
            return True
        return False
    
    def edit_custom_column(self, old_name, new_name):
        """カスタム項目名を編集"""
        if old_name in self.custom_columns and new_name not in self.custom_columns:
            index = self.custom_columns.index(old_name)
            self.custom_columns[index] = new_name
            self.save_settings()  # 即座に保存
            return True
        return False
    
    def delete_custom_column(self, column_name):
        """カスタム項目を削除"""
        if column_name in self.custom_columns:
            self.custom_columns.remove(column_name)
            self.save_settings()  # 即座に保存
            return True
        return False
    
    def delete_column_data(self, col_index):
        """指定された列の全データを削除"""
        keys_to_delete = [key for key in self.data.keys()
                          if key.split("-")[3] == str(col_index)]
        
        # 削除対象を年月でグループ化
        year_month_set = set()
        for key in keys_to_delete:
            parsed = self._parse_key(key)
            if parsed:
                year, month, _, _ = parsed
                year_month_set.add((year, month))
            del self.data[key]
        
        # 影響を受けた年月のデータを保存
        for year, month in year_month_set:
            month_data = {}
            for key, transactions in self.data.items():
                parsed = self._parse_key(key)
                if parsed and parsed[0] == year and parsed[1] == month:
                    day, col_idx = parsed[2], parsed[3]
                    day_key = str(day)
                    
                    if day_key not in month_data:
                        month_data[day_key] = []
                    
                    for transaction in transactions:
                        if len(transaction) >= 3:
                            new_transaction = {
                                "列目": str(col_idx),
                                "支払先": str(transaction[0]) if transaction[0] else "",
                                "金額": str(transaction[1]) if transaction[1] else "",
                                "詳細": str(transaction[2]) if transaction[2] else ""
                            }
                            month_data[day_key].append(new_transaction)
            
            self._save_month_data(year, month, month_data)
    
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
                            partner = str(row[0]).strip() if row[0] else ""
                            amount = str(row[1]).strip() if row[1] else ""
                            detail = str(row[2]).strip() if row[2] else ""
                            
                            if (search_text_lower in partner.lower() or
                                search_text_lower in amount.lower() or
                                search_text_lower in detail.lower()):
                                results.append({
                                    'year': year,
                                    'month': month,
                                    'day': day,
                                    'col_index': col_index,
                                    'partner': partner,
                                    'amount': amount,
                                    'detail': detail
                                })
            except (ValueError, IndexError):
                continue
        
        return results