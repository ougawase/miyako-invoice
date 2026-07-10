# 請求書自動生成ツール（miyako-invoice）

**株式会社 新家の本番業務システム**です。納品書のExcelをアップロードすると、
店舗ごとの請求書Excelが自動で出来上がります（毎月の請求業務で実際に使われています）。

> まず [CONTRIBUTING.md](CONTRIBUTING.md) を読んでください。
> 一番大事なルール：**mainに直接pushしない**（mainへのマージ＝即・本番反映です）。

## セットアップ（10分）

```bash
git clone https://github.com/ougawase/miyako-invoice.git
cd miyako-invoice
pip3 install -r requirements.txt
```

## 動かしてみる

```bash
python3 -m streamlit run app.py
```

ブラウザが開いたら、リポジトリ内の `sample_nouhin_1_居酒屋_海風.xlsx` など
サンプル納品書3つをアップロード →「請求書を生成する」。
店舗ごとの請求書がZIPでダウンロードできれば環境構築は成功です。

## テスト（変更したら必ず・全部グリーンが必須）

```bash
python3 -m pytest test_app.py -q     # 35件
```

## 仕組みの全体像

```
納品書Excel（複数）
   │  read_nouhinshо()      … A3=店舗名 / N4=納品日 / B18:B33×J列=明細 / V6:W11=価格表
   ▼
store_data（店舗ごとにまとめる）
   │  価格表にない商品 → 画面で単価をその場入力（¥0の請求書は作れない設計）
   ▼
create_invoice()             … テンプレートに数値を直書き（数式キャッシュに頼らない）
   │  verify_invoice_output() … 書いた値を読み直して検証（壊れたファイルは出さない）
   ▼
請求書Excel（店舗別・ZIP）
```

- 画面まわり: `app.py` ／ 変換ロジック: `invoice.py`（テストは本番と同じこのコードを直接叩く）
- 消費税はテンプレート内の数式（**8%・軽減税率＝食品卸の正解**）。**変更禁止**
- 金額・税率・テンプレートのセル位置に関わる変更は、**実装前にオーナー（平戸）へ相談**

## 変更の出し方

```bash
git checkout -b feat/変更の内容
# 実装 → テスト追加 → python3 -m pytest test_app.py -q が全グリーン
git push origin feat/変更の内容   # → GitHubでPull Requestを作成
```

オーナーがレビュー・承認したらマージ（＝自動で本番デプロイ）されます。
