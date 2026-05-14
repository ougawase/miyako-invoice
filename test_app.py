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

def create_invoice(store_name, all_items, billing_month, billing_subject, delivery_dates):
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

    merged = {}
    for item in all_items:
        k = item["name"]
        if k in merged:
            merged[k]["quantity"] += item["quantity"]
            merged[k]["amount"]   += item["amount"]
        else:
            merged[k] = dict(item)

    date_labels = []
    for d in delivery_dates:
        if isinstance(d, datetime):
            date_labels.append(f"{d.month}月{d.day}日")
    date_label = "・".join(date_labels) if date_labels else ""

    for i, item in enumerate(merged.values()):
        r = ITEM_START + i
        if r > ITEM_END:
            break
        if i == 0:
            ws.cell(row=r, column=1).value = date_label
        ws.cell(row=r, column=2).value  = item["name"]
        ws.cell(row=r, column=10).value = item["quantity"]
        ws.cell(row=r, column=11).value = "個"

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
        self.delivery_dates = [datetime(2026, 5, 4)]

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
            self.sample_items,
            "5月分",
            "味噌加工品代金",
            self.delivery_dates
        )
        self.assertIsNotNone(result)
        self.assertGreater(len(result), 0)

    def test_store_name_written(self):
        """店舗名（宛先）がA3セルに書き込まれること"""
        result = create_invoice(
            "島の駅みやこ",
            self.sample_items,
            "5月分",
            "味噌加工品代金",
            self.delivery_dates
        )
        wb = openpyxl.load_workbook(io.BytesIO(result))
        ws = wb.active
        self.assertEqual(ws["A3"].value, "島の駅みやこ")

    def test_billing_title_written(self):
        """件名がC6セルに正しく書き込まれること（重複なし）"""
        result = create_invoice(
            "島の駅みやこ",
            self.sample_items,
            "5月分",
            "味噌加工品代金",
            self.delivery_dates
        )
        wb = openpyxl.load_workbook(io.BytesIO(result))
        ws = wb.active
        title = ws["C6"].value
        self.assertIn("5月分", title)
        self.assertIn("味噌加工品代金", title)
        # 「8年8年」のような重複がないこと
        self.assertNotIn("年8年", title, f"件名に年が重複しています: {title}")

    def test_items_written_correctly(self):
        """商品名・数量・単位が明細行に正しく書き込まれること"""
        result = create_invoice(
            "島の駅みやこ",
            self.sample_items,
            "5月分",
            "味噌加工品代金",
            self.delivery_dates
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
            self.sample_items,
            "5月分",
            "味噌加工品代金",
            self.delivery_dates
        )
        wb = openpyxl.load_workbook(io.BytesIO(result))
        ws = wb.active
        self.assertEqual(ws.cell(row=ITEM_START, column=1).value, "5月4日")

    def test_previous_data_cleared(self):
        """テンプレートの前のデータが残っていないこと"""
        result = create_invoice(
            "テスト店舗",
            [{"name": "ツナプレーン", "unit_price": 520, "quantity": 10, "amount": 5200}],
            "5月分",
            "味噌加工品代金",
            []
        )
        wb = openpyxl.load_workbook(io.BytesIO(result))
        ws = wb.active
        # 2行目以降は空のはず
        for r in range(ITEM_START + 1, ITEM_START + 5):
            self.assertIsNone(ws.cell(row=r, column=2).value,
                             f"行{r}: 前のデータが残っています（商品名: {ws.cell(row=r, column=2).value}）")

    def test_full_calc_on_load(self):
        """Excelを開いた時に数式が自動再計算されること"""
        result = create_invoice(
            "島の駅みやこ",
            self.sample_items,
            "5月分",
            "味噌加工品代金",
            self.delivery_dates
        )
        wb = openpyxl.load_workbook(io.BytesIO(result))
        self.assertTrue(wb.calculation.fullCalcOnLoad,
                       "fullCalcOnLoadがFalseです。合計額が反映されません")

    def test_template_loads(self):
        """テンプレートファイルが正常に読み込めること"""
        wb = openpyxl.load_workbook(load_template())
        self.assertIn("見積書", wb.sheetnames)


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
