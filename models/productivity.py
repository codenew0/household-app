"""入力支援、予算、未確定処理、復元用スナップショット。"""
import calendar
import copy
import datetime
import json
import os
import uuid
from collections import Counter
from models.transactions import normalize, validate, cash_amount, is_pending


class ProductivityMixin:
    MAX_RESTORE_POINTS = 50
    SETTINGS_FIELDS = ('custom_columns', 'transaction_partners', 'subscriptions',
                       'payment_methods', 'subscription_runs', 'budgets')

    def snapshot(self, label, from_disk=False):
        source = self
        if from_disk:
            source = type(self)()
            source.load_settings()
            source.load_data()
        state = {'version': 1, 'label': label, 'created': datetime.datetime.now().isoformat(),
                 'data': source.data, 'settings': {name: sorted(getattr(source, name))
                    if isinstance(getattr(source, name), set) else getattr(source, name)
                    for name in self.SETTINGS_FIELDS}}
        path = os.path.join(self.JSON_DIR, 'restore_points',
                            datetime.datetime.now().strftime('%Y%m%d_%H%M%S_%f') + '_' + uuid.uuid4().hex[:6] + '.json')
        self._write_json(path, state)
        self._prune_restore_points()
        return path

    def _prune_restore_points(self):
        """復元ポイントが増え続けないよう、古いものから削除する。"""
        folder = os.path.join(self.JSON_DIR, 'restore_points')
        if not os.path.isdir(folder):
            return
        names = sorted(name for name in os.listdir(folder) if name.endswith('.json'))
        for name in names[:-self.MAX_RESTORE_POINTS]:
            try:
                os.remove(os.path.join(folder, name))
            except OSError:
                pass

    def restore_points(self):
        folder = os.path.join(self.JSON_DIR, 'restore_points')
        if not os.path.isdir(folder):
            return []
        result = []
        for name in sorted(os.listdir(folder), reverse=True):
            if not name.endswith('.json'):
                continue
            path = os.path.join(folder, name)
            try:
                with open(path, encoding='utf-8') as stream:
                    value = json.load(stream)
                result.append((path, value['created'], value['label'], sum(map(len, value['data'].values()))))
            except (OSError, ValueError, KeyError, TypeError):
                continue
        return result

    def restore_snapshot(self, path):
        with open(path, encoding='utf-8') as stream:
            state = json.load(stream)
        if state.get('version') != 1 or not isinstance(state.get('data'), dict):
            raise ValueError('復元データの形式が不正です。')
        for key, rows in state['data'].items():
            if not self._parse_key(key):
                raise ValueError('復元データの日付キーが不正です。')
            for row in rows:
                validate(row)
        settings = state['settings']
        if any(name not in settings for name in self.SETTINGS_FIELDS):
            raise ValueError('復元データの設定が不足しています。')
        self.snapshot('復元操作の直前')
        previous_data = copy.deepcopy(self.data)
        previous_settings = {name: copy.deepcopy(getattr(self, name)) for name in self.SETTINGS_FIELDS}
        months = {self._parse_key(key)[:2] for key in set(self.data) | set(state['data'])}
        try:
            self.data = copy.deepcopy(state['data'])
            for name in self.SETTINGS_FIELDS:
                setattr(self, name, set(settings[name]) if isinstance(previous_settings[name], set) else settings[name])
            for year, month in months:
                self._save_complete_month(year, month)
            self.save_settings()
        except Exception:
            self.data = previous_data
            for name, value in previous_settings.items():
                setattr(self, name, value)
            for year, month in months:
                self._save_complete_month(year, month)
            self.save_settings()
            raise

    def edit_entries(self, references, action):
        """一覧の参照は(key, index, 元の行)。古い画面からの更新を拒否する。"""
        for key, index, row in references:
            if index >= len(self.data.get(key, [])) or self.data[key][index] != row:
                raise ValueError('明細が更新されています。一覧を再表示してください。')
        previous = copy.deepcopy(self.data)
        self.snapshot('明細の' + ('一括確定' if action == 'confirm' else '削除') + '直前')
        try:
            for key, index, row in sorted(references, key=lambda value: (value[0], -value[1])):
                if action == 'confirm':
                    value = normalize(row)
                    value[5] = ''
                    self.data[key][index] = value
                elif action == 'delete':
                    del self.data[key][index]
                else:
                    raise ValueError('操作が不正です。')
            self.save_transactions({key for key, _, _ in references}, create_snapshot=False)
        except Exception:
            self.data = previous
            self.save_transactions({key for key, _, _ in references}, create_snapshot=False)
            raise

    def partner_suggestion(self, partner):
        choices = []
        for key, rows in self.data.items():
            parsed = self._parse_key(key)
            if not parsed or parsed[2] == 0:
                continue
            for row in rows:
                value = normalize(row)
                if value[0] == partner and not is_pending(value):
                    choices.append((parsed[3], value[4]))
        return Counter(choices).most_common(1)[0][0] if choices else None

    def monthly_spending(self, year, month):
        totals = Counter()
        for key, rows in self.data.items():
            parsed = self._parse_key(key)
            if parsed and parsed[:2] == (year, month) and parsed[2] > 0:
                totals[parsed[3]] += sum(cash_amount(row) for row in rows)
        return totals

    def subscription_forecast(self, today=None):
        """有効な定期支払いと、今日以降の次回支払日を返す。"""
        today = today or datetime.date.today()
        result = []
        for item in self.subscriptions:
            values = self.recurring_payment_values(item)
            if not values['enabled']:
                continue
            skip_through = datetime.date.fromisoformat(values['skip_through']) if values.get('skip_through') else None
            for _, due in self.iter_recurring_payment_dates(values):
                if skip_through and due <= skip_through:
                    continue
                if due >= today:
                    result.append((values, due))
                    break
        return result

    def recurring_payment_year_total(self, today=None):
        """今日から12か月間に予定されている定期支払いの合計。"""
        today = today or datetime.date.today()
        through_month = self._add_months(datetime.date(today.year, today.month, 1), 12)
        through = datetime.date(
            through_month.year, through_month.month,
            min(today.day, calendar.monthrange(through_month.year, through_month.month)[1]))
        total = 0
        for raw in self.subscriptions:
            item = self.recurring_payment_values(raw)
            if not item['enabled']:
                continue
            skip_through = datetime.date.fromisoformat(item['skip_through']) if item.get('skip_through') else None
            total += sum(int(item['amount']) for _, due in self.iter_recurring_payment_dates(item, through)
                         if today <= due < through and (not skip_through or due > skip_through))
        return total

    def duplicate_candidates(self, imported, include_existing=True):
        seen = set()
        for key, rows in (self.data.items() if include_existing else []):
            date = self._parse_key(key)[:3]
            for row in rows:
                value = normalize(row)
                seen.add((date, value[0], validate(value)[1]))
        matches = []
        for key, rows in imported.items():
            date = self._parse_key(key)[:3]
            for row in rows:
                value = normalize(row)
                signature = (date, value[0], validate(value)[1])
                if signature in seen:
                    matches.append((key, value))
                seen.add(signature)
        return matches
