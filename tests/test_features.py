import copy
import datetime
import os
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import config
from models.data_manager import DataManager
from models.clipboard import decode_clipboard, PLAIN_AMOUNT, DETAIL_ROWS, CELL_BLOCK
from models.transactions import cash_amount, normalize, PENDING, validate


class FeatureTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root_dir = self.directory.name
        self.patch = patch.multiple(
            config,
            JSON_DIR=self.root_dir,
            SETTINGS_FILE=os.path.join(self.root_dir, 'settings.json'),
            DATA_ROOT_DIR=os.path.join(self.root_dir, 'data'),
        )
        self.patch.start()
        self.addCleanup(self.patch.stop)
        self.manager = DataManager()

    def subscription(self, day=31, start='2026-01-01'):
        return self.manager.save_subscription('動画', '配信会社', '1000', day, 'カード', start)

    def test_month_end_catchup_pending_and_restart_idempotency(self):
        self.subscription()
        self.assertEqual(self.manager.generate_due_subscriptions(datetime.date(2026, 3, 1)), 2)
        self.assertIn('2026-2-28-7', self.manager.data)
        self.assertEqual(sum(cash_amount(row) for rows in self.manager.data.values() for row in rows), 0)
        other = DataManager()
        other.load_settings()
        other.load_data()
        self.assertEqual(other.generate_due_subscriptions(datetime.date(2026, 3, 1)), 0)
        self.assertEqual(len(other.data), 2)

    def test_deleted_occurrence_stays_deleted(self):
        self.subscription()
        self.manager.generate_due_subscriptions(datetime.date(2026, 1, 31))
        self.manager.delete_transaction_data('2026-1-31-7')
        self.manager.save_transaction('2026-1-31-7')
        other = DataManager()
        other.load_settings()
        other.load_data()
        self.assertEqual(other.generate_due_subscriptions(datetime.date(2026, 1, 31)), 0)
        self.assertEqual(other.data, {})

    def test_start_date_disabled_and_leap_year(self):
        s = self.subscription(start='2024-02-01')
        self.assertEqual(self.manager.generate_due_subscriptions(datetime.date(2024, 2, 28)), 0)
        self.assertEqual(self.manager.generate_due_subscriptions(datetime.date(2024, 2, 29)), 1)
        self.manager.subscriptions[0]['enabled'] = False
        self.assertEqual(self.manager.generate_due_subscriptions(datetime.date(2024, 5, 31)), 0)

    def test_reenabled_subscription_does_not_backfill_disabled_months(self):
        item = self.manager.save_subscription('動画', '店', '100', 1, '', '2026-01-01')
        self.manager.generate_due_subscriptions(datetime.date(2026, 1, 1))
        item['enabled'] = False
        self.manager.generate_due_subscriptions(datetime.date(2026, 3, 1))
        item['enabled'] = True
        self.assertEqual(self.manager.generate_due_subscriptions(datetime.date(2026, 4, 1)), 1)
        self.assertNotIn('2026-2-1-7', self.manager.data)
        self.assertNotIn('2026-3-1-7', self.manager.data)

    def test_points_pending_json_csv_roundtrip(self):
        rows = [['店', '1000', '商品', '200', 'カード', '', ''],
                ['動画', '1000', '契約', '0', 'カード', PENDING, 'id:2026-07']]
        self.manager.data = {'2026-7-1-7': rows}
        self.manager.save_data()
        self.manager.load_data()
        self.assertEqual(self.manager.data['2026-7-1-7'], rows)
        self.assertEqual(sum(map(cash_amount, rows)), 800)
        path = os.path.join(self.root_dir, 'roundtrip.csv')
        self.manager.export_csv(path, (2026, 7), (2026, 7), config.DefaultColumns.ITEMS)
        self.manager.import_csv(path, (2026, 7), (2026, 7), 'replace', 12)
        self.assertEqual(self.manager.data['2026-7-1-7'], rows)
        self.assertEqual(sum(int(item['amount']) for item in self.manager.search_transactions('')), 800)
        results = self.manager.search_transactions('カード')
        self.assertEqual(results[0]['gross_amount'], '1000')
        self.assertEqual(results[0]['points'], '200')
        self.assertEqual(results[0]['payment_method'], 'カード')
        self.assertEqual(results[0]['memo'], '商品')
        self.assertEqual(results[1]['status'], PENDING)

    def test_save_failure_does_not_mark_occurrence_generated(self):
        self.subscription()
        with patch.object(self.manager, 'save_transactions', side_effect=OSError('disk full')):
            with self.assertRaises(OSError):
                self.manager.generate_due_subscriptions(datetime.date(2026, 1, 31))
        self.assertEqual(self.manager.data, {})
        self.assertEqual(self.manager.subscription_runs, set())
        self.assertEqual(self.manager.generate_due_subscriptions(datetime.date(2026, 1, 31)), 1)

    def test_recurring_payment_interval_category_and_start_anchor(self):
        self.manager.save_subscription(
            '水道', '水道局', '3000', 10, '口座振替', '2026-01-20',
            kind='公共料金', category='光熱・通信', interval_months=3)
        self.assertEqual(self.manager.generate_due_subscriptions(datetime.date(2026, 7, 31)), 2)
        self.assertIn('2026-4-10-9', self.manager.data)
        self.assertIn('2026-7-10-9', self.manager.data)
        self.assertNotIn('2026-1-10-9', self.manager.data)

    def test_installment_stops_at_count_and_records_progress(self):
        self.manager.save_subscription(
            'パソコン', '家電店', '10000', 15, 'カード', '2026-01-01',
            kind='分割払い', category='通販', interval_months=1, payment_count=3)
        self.assertEqual(self.manager.generate_due_subscriptions(datetime.date(2026, 12, 31)), 3)
        memos = [row[2] for rows in self.manager.data.values() for row in rows]
        self.assertEqual(memos, ['パソコン (1/3回)', 'パソコン (2/3回)', 'パソコン (3/3回)'])

    def test_installment_requires_count_and_legacy_subscription_defaults(self):
        with self.assertRaises(ValueError):
            self.manager.save_subscription(
                '端末', '店', '1000', 1, 'カード', '2026-01-01', kind='分割払い')
        legacy = {'id': 'old', 'name': '動画', 'partner': '店', 'amount': '500',
                  'day': 1, 'method': 'カード', 'start': '2026-01-01', 'enabled': True}
        values = self.manager.recurring_payment_values(legacy)
        self.assertEqual((values['kind'], values['category'], values['interval_months']),
                         ('サブスクリプション', 'サブスク', 1))

    def test_old_three_field_rows_still_count_cash(self):
        self.assertEqual(cash_amount(['店', '1200', 'メモ']), 1200)
        self.assertEqual(normalize(['店', '1200', 'メモ'])[3:], ['', '', '', ''])

    def test_payment_methods_persist_and_delete_keeps_history(self):
        self.manager.add_payment_method('テストカード')
        self.manager.data = {'2026-7-1-1': [['店', '100', '', '0', 'テストカード']]}
        other = DataManager()
        other.load_settings()
        self.assertIn('テストカード', other.payment_methods)
        self.manager.remove_payment_method('テストカード')
        other.load_settings()
        self.assertNotIn('テストカード', other.payment_methods)
        self.assertEqual(self.manager.data['2026-7-1-1'][0][4], 'テストカード')
        with self.assertRaises(ValueError):
            self.manager.add_payment_method('  ')

    def test_points_are_deducted_and_full_points_payment_is_zero(self):
        self.assertEqual(cash_amount(validate(['店', '1,000', '商品', '200'])), 800)
        self.assertEqual(cash_amount(validate(['店', '1000', '商品', '1000'])), 0)
        self.assertEqual(cash_amount(validate(['店', '-100', '返金', '0'])), -100)
        with self.assertRaises(ValueError):
            validate(['店', '1000', '商品', '1001'])
        with self.assertRaises(ValueError):
            validate(['店', '1000', '商品', '-1'])

    def test_clipboard_formats_are_validated(self):
        kind, rows = decode_clipboard('1,200')
        self.assertEqual((kind, rows[0][1]), (PLAIN_AMOUNT, '1200'))
        kind, rows = decode_clipboard('[["店", "500", "メモ"]]')
        self.assertEqual((kind, rows[0][:3]), (DETAIL_ROWS, ['店', '500', 'メモ']))
        kind, cells = decode_clipboard(
            '[{"day": 1, "col_idx": 2, "data": [["店", "300", ""]]}]')
        self.assertEqual((kind, cells[0]['data'][0][1]), (CELL_BLOCK, '300'))
        with self.assertRaises(ValueError):
            decode_clipboard('[{"day": "1", "col_idx": 2, "data": []}]')


if __name__ == '__main__':
    unittest.main()
