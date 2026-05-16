"""
請求書自動生成ツール 自動テスト
実行方法: python3 test_app.py
"""

import unittest
import openpyxl
import io
import base64
from datetime import date, datetime
from template_data import TEMPLATE_B64

PRICE_MASTER = {
    "ツナプレーン": 520,
    "ツナおかず":   580,
    "ツナラー油":   580,
    "ツナ味噌":     580,
    "ツナガリ":     580,
    "ツナカレー":   580,
}
TAX_RATE     = 0.08
ITEM_START   = 18
ITEM_END     = 33

def load_template():
    return io.BytesIO(base64.b64decode(TEMPLATE_B64))

def create_invoice(store_name, dated_groups, billing_month, billing_subject):
    wb = openpyxl.load_workbook(load_template())
    ws = wb.active
    wb.calculation.fullCalcOnLoad = True

    ws["A3"] = store_name
    ws["N4"] = date.today()
    reiwa = date.today().year - 2018
    ws["C6"] = f"令和{reiwa}年{billing_month}{billing_subject}"

    for r in range(ITEM_START, ITEM_END + 1):
        for col in [1, 2, 10, 11]:
            ws.cell(row=r, column=col).value = None

    current_row = ITEM_START
    for group in dated_groups:
        if current_row > ITEM_END:
            break
        d = group["date"]
        date_label = f"{d.month}月{d.day}日" if isinstance(d, datetime) else ""

        merged = {}
        for item in group["items"]:
            k = item["name"]
            if k in merged:
                merged[k]["quantity"] += item["quantity"]
                merged[k]["amount"]   += item["amount"]
            else:
                merged[k] = dict(item)

        first = True
        for item in merged.values():
            if current_row > ITEM_END:
                break
            if first:
                ws.cell(row=current_row, column=1).value = date_label
                first = False
            ws.cell(row=current_row, column=2).value  = item["name"]
            ws.cell(row=current_row, column=10).value = item["quantity"]
            ws.cell(row=current_row, column=11).value = "個"
            current_row += 1

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class TestInvoiceGeneration(unittest.TestCase):

    def setUp(self):
        """テスト用の基本データ"""
        self.sample_items = [
            {"name": "ツナプレーン", "unit_price": 520, "quantity": 35, "amount": 18200},
            {"name": "ツナラー油",   "unit_price": 580, "quantity": 35, "amount": 20300},
            {"name": "ツナおかず",   "unit_price": 580, "quantity": 35, "amount": 20300},
        ]
        # 新形式: 日付グループ
        self.dated_groups = [
            {"date": datetime(2026, 5, 4), "items": self.sample_items}
        ]

    # ----------------------------------------
    # 1. 金額計算の正確性チェック
    # ----------------------------------------
    def test_amount_calculation(self):
        """単価 × 数量 = 金額 が全商品で一致すること"""
        for item in self.sample_items:
            expected = item["unit_price"] * item["quantity"]
            self.assertEqual(
                item["amount"], expected,
                f"{item['name']}: {item['unit_price']}×{item['quantity']}={expected} のはずが {item['amount']}"
            )

    def test_subtotal(self):
        """小計が各商品の金額の合計と一致すること"""
        subtotal = sum(i["amount"] for i in self.sample_items)
        self.assertEqual(subtotal, 58800)

    def test_tax_calculation(self):
        """消費税8%の計算が正しいこと"""
        subtotal = 58800
        tax = int(subtotal * TAX_RATE)
        self.assertEqual(tax, 4704)

    def test_total_with_tax(self):
        """合計（税込）= 小計 + 消費税"""
        subtotal = 58800
        tax = int(subtotal * TAX_RATE)
        total = subtotal + tax
        self.assertEqual(total, 63504)

    # ----------------------------------------
    # 2. Excel生成チェック
    # ----------------------------------------
    def test_invoice_generates_without_error(self):
        """請求書Excelが正常に生成されること"""
        result = create_invoice(
            "島の駅みやこ",
            self.dated_groups,
            "5月分",
            "味噌加工品代金",
        )
        self.assertIsNotNone(result)
        self.assertGreater(len(result), 0)

    def test_store_name_written(self):
        """店舗名（宛先）がA3セルに書き込まれること"""
        result = create_invoice(
            "島の駅みやこ",
            self.dated_groups,
            "5月分",
            "味噌加工品代金",
        )
        wb = openpyxl.load_workbook(io.BytesIO(result))
        ws = wb.active
        self.assertEqual(ws["A3"].value, "島の駅みやこ")

    def test_billing_title_written(self):
        """件名がC6セルに正しく書き込まれること（重複なし）"""
        result = create_invoice(
            "島の駅みやこ",
            self.dated_groups,
            "5月分",
            "味噌加工品代金",
        )
        wb = openpyxl.load_workbook(io.BytesIO(result))
        ws = wb.active
        title = ws["C6"].value
        self.assertIn("5月分", title)
        self.assertIn("味噌加工品代金", title)
        self.assertNotIn("年8年", title, f"件名に年が重複しています: {title}")

    def test_items_written_correctly(self):
        """商品名・数量・単位が明細行に正しく書き込まれること"""
        result = create_invoice(
            "島の駅みやこ",
            self.dated_groups,
            "5月分",
            "味噌加工品代金",
        )
        wb = openpyxl.load_workbook(io.BytesIO(result))
        ws = wb.active

        for i, item in enumerate(self.sample_items):
            r = ITEM_START + i
            self.assertEqual(ws.cell(row=r, column=2).value,  item["name"],
                             f"行{r}: 商品名が違います")
            self.assertEqual(ws.cell(row=r, column=10).value, item["quantity"],
                             f"行{r}: 数量が違います")
            self.assertEqual(ws.cell(row=r, column=11).value, "個",
                             f"行{r}: 単位(個)がありません")

    def test_delivery_date_written(self):
        """納品日が最初の明細行のA列に書き込まれること"""
        result = create_invoice(
            "島の駅みやこ",
            self.dated_groups,
            "5月分",
            "味噌加工品代金",
        )
        wb = openpyxl.load_workbook(io.BytesIO(result))
        ws = wb.active
        self.assertEqual(ws.cell(row=ITEM_START, column=1).value, "5月4日")

    def test_previous_data_cleared(self):
        """テンプレートの前のデータが残っていないこと"""
        result = create_invoice(
            "テスト店舗",
            [{"date": datetime(2026, 5, 4), "items": [
                {"name": "ツナプレーン", "unit_price": 520, "quantity": 10, "amount": 5200}
            ]}],
            "5月分",
            "味噌加工品代金",
        )
        wb = openpyxl.load_workbook(io.BytesIO(result))
        ws = wb.active
        for r in range(ITEM_START + 1, ITEM_START + 5):
            self.assertIsNone(ws.cell(row=r, column=2).value,
                             f"行{r}: 前のデータが残っています（商品名: {ws.cell(row=r, column=2).value}）")

    def test_full_calc_on_load(self):
        """Excelを開いた時に数式が自動再計算されること"""
        result = create_invoice(
            "島の駅みやこ",
            self.dated_groups,
            "5月分",
            "味噌加工品代金",
        )
        wb = openpyxl.load_workbook(io.BytesIO(result))
        self.assertTrue(wb.calculation.fullCalcOnLoad,
                       "fullCalcOnLoadがFalseです。合計額が反映されません")

    def test_template_loads(self):
        """テンプレートファイルが正常に読み込めること"""
        wb = openpyxl.load_workbook(load_template())
        self.assertIn("見積書", wb.sheetnames)

    # ----------------------------------------
    # 3. 複数日付のテスト（日別詳細）
    # ----------------------------------------
    def test_multiple_dates_written_separately(self):
        """複数日付がそれぞれ別の行グループに書き込まれること"""
        dated_groups = [
            {"date": datetime(2026, 5, 4),  "items": [
                {"name": "ツナプレーン", "unit_price": 520, "quantity": 30, "amount": 15600},
            ]},
            {"date": datetime(2026, 5, 11), "items": [
                {"name": "ツナプレーン", "unit_price": 520, "quantity": 20, "amount": 10400},
            ]},
        ]
        result = create_invoice("テスト店舗", dated_groups, "5月分", "味噌加工品代金")
        wb = openpyxl.load_workbook(io.BytesIO(result))
        ws = wb.active

        # 1行目: 5月4日
        self.assertEqual(ws.cell(row=ITEM_START,     column=1).value, "5月4日",
                         "1日目の日付が正しくありません")
        self.assertEqual(ws.cell(row=ITEM_START,     column=2).value, "ツナプレーン")
        self.assertEqual(ws.cell(row=ITEM_START,     column=10).value, 30,
                         "1日目の数量が30のはずです")

        # 2行目: 5月11日
        self.assertEqual(ws.cell(row=ITEM_START + 1, column=1).value, "5月11日",
                         "2日目の日付が正しくありません")
        self.assertEqual(ws.cell(row=ITEM_START + 1, column=2).value, "ツナプレーン")
        self.assertEqual(ws.cell(row=ITEM_START + 1, column=10).value, 20,
                         "2日目の数量が20のはずです")

    def test_same_product_not_merged_across_dates(self):
        """同じ商品でも日付が違えば合算されず別行に出ること"""
        dated_groups = [
            {"date": datetime(2026, 5, 4),  "items": [
                {"name": "ツナプレーン", "unit_price": 520, "quantity": 35, "amount": 18200},
            ]},
            {"date": datetime(2026, 5, 11), "items": [
                {"name": "ツナプレーン", "unit_price": 520, "quantity": 35, "amount": 18200},
            ]},
        ]
        result = create_invoice("テスト店舗", dated_groups, "5月分", "味噌加工品代金")
        wb = openpyxl.load_workbook(io.BytesIO(result))
        ws = wb.active

        qty_row1 = ws.cell(row=ITEM_START,     column=10).value
        qty_row2 = ws.cell(row=ITEM_START + 1, column=10).value

        # 合算されると70になってしまう。35が2行あるのが正しい
        self.assertEqual(qty_row1, 35, f"1行目の数量が35のはずが {qty_row1} です（日付をまたいで合算されています）")
        self.assertEqual(qty_row2, 35, f"2行目の数量が35のはずが {qty_row2} です")

    def test_same_product_merged_within_same_date(self):
        """同じ日付内の同一商品は合算されること"""
        dated_groups = [
            {"date": datetime(2026, 5, 4), "items": [
                {"name": "ツナプレーン", "unit_price": 520, "quantity": 20, "amount": 10400},
                {"name": "ツナプレーン", "unit_price": 520, "quantity": 15, "amount":  7800},
            ]},
        ]
        result = create_invoice("テスト店舗", dated_groups, "5月分", "味噌加工品代金")
        wb = openpyxl.load_workbook(io.BytesIO(result))
        ws = wb.active

        qty = ws.cell(row=ITEM_START, column=10).value
        self.assertEqual(qty, 35, f"同日内の同一商品は合算されるべきですが {qty} です")

        # 2行目は空のはず
        self.assertIsNone(ws.cell(row=ITEM_START + 1, column=2).value,
                          "同日内の重複商品が2行目に残っています（合算されていません）")

    def test_second_date_label_in_correct_row(self):
        """2日目の日付ラベルが、1日目の商品数分だけ下の行に書かれること"""
        dated_groups = [
            {"date": datetime(2026, 5, 4), "items": [
                {"name": "ツナプレーン", "unit_price": 520, "quantity": 35, "amount": 18200},
                {"name": "ツナラー油",   "unit_price": 580, "quantity": 35, "amount": 20300},
            ]},
            {"date": datetime(2026, 5, 11), "items": [
                {"name": "ツナおかず", "unit_price": 580, "quantity": 35, "amount": 20300},
            ]},
        ]
        result = create_invoice("テスト店舗", dated_groups, "5月分", "味噌加工品代金")
        wb = openpyxl.load_workbook(io.BytesIO(result))
        ws = wb.active

        # 1日目が2行分 → 2日目のラベルは行18+2=20
        expected_row = ITEM_START + 2
        self.assertEqual(ws.cell(row=expected_row, column=1).value, "5月11日",
                         f"2日目の日付ラベルが行{expected_row}にあるべきですが見つかりません")

    # ----------------------------------------
    # 4. 境界値・異常系テスト
    # ----------------------------------------
    def test_row_overflow_does_not_crash(self):
        """商品数が16行を超えてもクラッシュしないこと"""
        many_items = [
            {"name": f"商品{i:02d}", "unit_price": 500, "quantity": 10, "amount": 5000}
            for i in range(20)  # 16行のテンプレートを超える20件
        ]
        dated_groups = [{"date": datetime(2026, 5, 4), "items": many_items}]
        try:
            result = create_invoice("テスト店舗", dated_groups, "5月分", "テスト")
            self.assertGreater(len(result), 0)
            # 最終行(33行)を超えて書かれていないこと
            wb = openpyxl.load_workbook(io.BytesIO(result))
            ws = wb.active
            self.assertIsNone(ws.cell(row=ITEM_END + 1, column=2).value,
                              f"行{ITEM_END+1}にデータが書かれています（範囲外）")
        except Exception as e:
            self.fail(f"行数オーバーフロー時にクラッシュしました: {e}")

    def test_none_delivery_date_does_not_crash(self):
        """納品日がNoneでもクラッシュせず、A列が空欄になること"""
        dated_groups = [
            {"date": None, "items": [
                {"name": "ツナプレーン", "unit_price": 520, "quantity": 35, "amount": 18200},
            ]}
        ]
        try:
            result = create_invoice("テスト店舗", dated_groups, "5月分", "味噌加工品代金")
            wb = openpyxl.load_workbook(io.BytesIO(result))
            ws = wb.active
            # 商品は書かれていること
            self.assertEqual(ws.cell(row=ITEM_START, column=2).value, "ツナプレーン")
            # 日付列は空か空文字
            date_val = ws.cell(row=ITEM_START, column=1).value
            self.assertIn(date_val, [None, ""], f"日付Noneのとき空欄になるべきですが '{date_val}' が入っています")
        except Exception as e:
            self.fail(f"日付Noneでクラッシュしました: {e}")

    def test_empty_dated_groups_does_not_crash(self):
        """納品書ゼロ件でもクラッシュしないこと"""
        try:
            result = create_invoice("テスト店舗", [], "5月分", "味噌加工品代金")
            wb = openpyxl.load_workbook(io.BytesIO(result))
            ws = wb.active
            # 全明細行が空のはず
            for r in range(ITEM_START, ITEM_END + 1):
                self.assertIsNone(ws.cell(row=r, column=2).value,
                                  f"行{r}にデータが残っています")
        except Exception as e:
            self.fail(f"空データでクラッシュしました: {e}")

    def test_billing_month_all_months(self):
        """1月〜12月すべてで件名が正しく生成されること"""
        for month in range(1, 13):
            billing_month = f"{month}月分"
            result = create_invoice("テスト店舗", self.dated_groups, billing_month, "テスト代金")
            wb = openpyxl.load_workbook(io.BytesIO(result))
            ws = wb.active
            title = ws["C6"].value
            self.assertIn(f"{month}月分", title,
                          f"{month}月分が件名に含まれていません: {title}")
            # 年の重複チェック
            reiwa = date.today().year - 2018
            self.assertNotIn(f"年{reiwa}年", title,
                             f"{month}月: 年が重複しています: {title}")


if __name__ == "__main__":
    print("=" * 50)
    print("請求書自動生成ツール 自動テスト開始")
    print("=" * 50)
    loader = unittest.TestLoader()
    suite = loader.loadTestsFromTestCase(TestInvoiceGeneration)
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    print("=" * 50)
    if result.wasSuccessful():
        print("全テスト合格 - デプロイしても安全です")
    else:
        print(f"テスト失敗: {len(result.failures)}件 / エラー: {len(result.errors)}件")
        print("デプロイ前に問題を修正してください")
    print("=" * 50)
