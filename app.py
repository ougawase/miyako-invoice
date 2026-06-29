import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import date, datetime
from invoice import (
    read_nouhinshо, create_invoice, verify_invoice_output,
    load_nouhin_template_bytes,
    ITEM_START, ITEM_END,
)

st.set_page_config(page_title="Invoice Generator | 株式会社 新家", layout="centered")

st.markdown("""
<style>
    /* ベース */
    html, body, [class*="css"], .stApp {
        font-family: 'Hiragino Kaku Gothic Pro', 'Noto Sans JP', sans-serif;
        background-color: #0a0a0a !important;
        color: rgba(255,255,255,0.9);
        line-height: 1.65;
    }

    /* Streamlit内部の白背景を全て上書き */
    .main, .block-container {
        background-color: #0a0a0a !important;
    }
    section[data-testid="stSidebar"] {
        background-color: #0f0f0f !important;
    }

    /* ヘッダー */
    .app-header {
        padding: 3rem 0 1.8rem 0;
        border-bottom: 1px solid rgba(255,255,255,0.08);
        margin-bottom: 2.5rem;
    }
    .app-header h1 {
        font-size: 1.3rem;
        font-weight: 700;
        letter-spacing: 0.08em;
        color: rgba(255,255,255,0.95);
        margin: 0;
    }
    .app-header p {
        font-size: 0.78rem;
        color: rgba(255,255,255,0.4);
        margin: 0.4rem 0 0 0;
        letter-spacing: 0.04em;
    }

    /* セクションラベル */
    .section-label {
        font-size: 0.65rem;
        font-weight: 700;
        letter-spacing: 0.15em;
        text-transform: uppercase;
        color: rgba(255,255,255,0.35);
        margin-bottom: 0.5rem;
    }

    /* ボタン */
    .stButton > button {
        background-color: #ffffff !important;
        color: #0a0a0a !important;
        border: none !important;
        border-radius: 6px;
        font-size: 0.85rem;
        font-weight: 600;
        letter-spacing: 0.06em;
        padding: 0.7rem 1.5rem;
        transition: opacity 0.15s ease;
    }
    .stButton > button:hover {
        background-color: rgba(255,255,255,0.88) !important;
        color: #0a0a0a !important;
    }

    /* プライマリボタン */
    .stButton > button[kind="primary"] {
        background-color: #ffffff !important;
        color: #0a0a0a !important;
    }

    /* 入力フィールド */
    input, textarea, [data-testid="stTextInput"] input {
        background-color: #141414 !important;
        color: rgba(255,255,255,0.9) !important;
        border: 1px solid rgba(255,255,255,0.1) !important;
        border-radius: 6px !important;
    }
    input:focus, textarea:focus {
        border-color: rgba(255,255,255,0.3) !important;
        box-shadow: none !important;
    }

    /* ラベルテキスト */
    label, .stTextInput label {
        color: rgba(255,255,255,0.5) !important;
        font-size: 0.78rem !important;
    }

    /* ファイルアップローダー */
    [data-testid="stFileUploader"] {
        background-color: #141414 !important;
        border: 1px dashed rgba(255,255,255,0.15) !important;
        border-radius: 8px !important;
    }
    [data-testid="stFileUploader"] *,
    [data-testid="stFileUploader"] p,
    [data-testid="stFileUploader"] span,
    [data-testid="stFileUploader"] small,
    [data-testid="stFileUploader"] div,
    [data-testid="stFileUploaderDropzone"] *,
    [data-testid="stFileUploaderDropzone"] {
        background-color: #141414 !important;
        color: rgba(255,255,255,0.65) !important;
    }
    [data-testid="stFileUploaderDropzone"] {
        border: none !important;
    }
    /* アップロード済みファイル表示 */
    [data-testid="stFileUploaderFile"] {
        background-color: #1e1e1e !important;
        border: 1px solid rgba(255,255,255,0.08) !important;
        border-radius: 6px !important;
    }
    [data-testid="stFileUploaderFile"] * {
        color: rgba(255,255,255,0.75) !important;
    }
    /* Browse files ボタン → ファイルを選択 */
    [data-testid="stFileUploader"] button {
        background-color: #2a2a2a !important;
        color: rgba(255,255,255,0.8) !important;
        border: 1px solid rgba(255,255,255,0.15) !important;
        border-radius: 6px !important;
    }
    [data-testid="stFileUploader"] button::after {
        content: "　📂 ファイルを選択";
    }
    [data-testid="stFileUploader"] button > span {
        display: none;
    }

    /* ダウンロードボタン */
    [data-testid="stDownloadButton"] button {
        background-color: #1a1a1a !important;
        color: rgba(255,255,255,0.85) !important;
        border: 1px solid rgba(255,255,255,0.2) !important;
        border-radius: 6px !important;
        font-size: 0.85rem !important;
        font-weight: 600 !important;
        letter-spacing: 0.04em !important;
        padding: 0.7rem 1.5rem !important;
        width: 100% !important;
        transition: background-color 0.15s ease !important;
    }
    [data-testid="stDownloadButton"] button:hover {
        background-color: #252525 !important;
    }

    /* ステップカード */
    .step-card {
        background-color: #111111;
        border: 1px solid rgba(255,255,255,0.08);
        border-radius: 10px;
        padding: 1.4rem 1.6rem 1.2rem;
        margin-bottom: 1rem;
    }
    .step-header {
        display: flex;
        align-items: center;
        gap: 0.7rem;
        margin-bottom: 0.8rem;
    }
    .step-num {
        background-color: rgba(255,255,255,0.08);
        color: rgba(255,255,255,0.5);
        font-size: 0.7rem;
        font-weight: 700;
        letter-spacing: 0.1em;
        border-radius: 4px;
        padding: 0.15rem 0.5rem;
    }
    .step-title {
        font-size: 0.88rem;
        font-weight: 600;
        color: rgba(255,255,255,0.9);
        letter-spacing: 0.04em;
    }
    .step-desc {
        font-size: 0.75rem;
        color: rgba(255,255,255,0.4);
        margin-bottom: 0.8rem;
        line-height: 1.6;
    }
    .format-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 0.78rem;
        margin-top: 0.3rem;
    }
    .format-table th {
        background-color: rgba(255,255,255,0.05);
        color: rgba(255,255,255,0.45);
        font-weight: 600;
        letter-spacing: 0.06em;
        padding: 0.35rem 0.6rem;
        text-align: left;
        border-bottom: 1px solid rgba(255,255,255,0.08);
    }
    .format-table td {
        padding: 0.4rem 0.6rem;
        color: rgba(255,255,255,0.75);
        border-bottom: 1px solid rgba(255,255,255,0.04);
        vertical-align: top;
    }
    .format-table td.cell-ref {
        font-family: monospace;
        color: rgba(255,255,255,0.5);
        white-space: nowrap;
    }
    .required-badge {
        display: inline-block;
        background-color: rgba(239,83,80,0.15);
        color: #ef9a9a;
        font-size: 0.65rem;
        font-weight: 700;
        letter-spacing: 0.06em;
        border-radius: 3px;
        padding: 0.05rem 0.35rem;
        margin-left: 0.3rem;
    }

    /* expander */
    [data-testid="stExpander"] {
        background-color: #141414 !important;
        border: 1px solid rgba(255,255,255,0.07) !important;
        border-radius: 8px !important;
    }
    [data-testid="stExpander"] summary {
        color: rgba(255,255,255,0.85) !important;
    }

    /* dataframe / テーブル */
    [data-testid="stDataFrame"], .stDataFrame {
        background-color: #141414 !important;
        border: 1px solid rgba(255,255,255,0.07) !important;
        border-radius: 6px !important;
    }

    /* メトリクス */
    [data-testid="stMetric"] {
        background-color: #141414;
        border: 1px solid rgba(255,255,255,0.07);
        border-radius: 8px;
        padding: 1rem;
    }
    [data-testid="stMetricLabel"] {
        color: rgba(255,255,255,0.45) !important;
        font-size: 0.72rem !important;
        letter-spacing: 0.06em;
    }
    [data-testid="stMetricValue"] {
        color: rgba(255,255,255,0.95) !important;
        font-size: 1.1rem !important;
        font-weight: 600;
    }

    /* caption / info テキスト */
    .stCaption, [data-testid="stCaptionContainer"] {
        color: rgba(255,255,255,0.35) !important;
    }

    /* warning / error / success */
    [data-testid="stAlert"] {
        background-color: #1a1a1a !important;
        border: 1px solid rgba(255,255,255,0.1) !important;
        border-radius: 6px !important;
        color: rgba(255,255,255,0.8) !important;
    }

    /* progress bar */
    [data-testid="stProgressBar"] > div {
        background-color: rgba(255,255,255,0.15) !important;
        border-radius: 4px;
    }
    [data-testid="stProgressBar"] > div > div {
        background-color: #ffffff !important;
    }

    /* divider */
    hr {
        border-color: rgba(255,255,255,0.06) !important;
    }

    /* フッター非表示 */
    footer { visibility: hidden; }
    #MainMenu { visibility: hidden; }
</style>

<div class="app-header">
    <h1>Invoice Generator</h1>
    <p>株式会社 新家 / 納品書から請求書を自動生成します</p>
</div>
""", unsafe_allow_html=True)


# ---- UI ----

reiwa_now = date.today().year - 2018

# ── 請求書設定 ──
st.markdown("""
<div class="step-card">
  <div class="step-header">
    <span class="step-num">SETTINGS</span>
    <span class="step-title">請求書の設定</span>
  </div>
</div>
""", unsafe_allow_html=True)

col1, col2, col3 = st.columns([2, 2, 1])
with col1:
    st.markdown('<p class="section-label">請求月</p>', unsafe_allow_html=True)
    billing_month = st.text_input(
        "請求月",
        value=f"{date.today().month}月分",
        label_visibility="collapsed"
    )
with col2:
    st.markdown('<p class="section-label">件名（商品・サービス名）</p>', unsafe_allow_html=True)
    billing_subject = st.text_input(
        "件名",
        value="味噌加工品代金",
        label_visibility="collapsed"
    )
with col3:
    st.markdown('<p class="section-label">消費税率</p>', unsafe_allow_html=True)
    tax_rate_label = st.selectbox(
        "消費税率",
        options=["10%", "8%"],
        label_visibility="collapsed"
    )
    tax_rate = 0.10 if tax_rate_label == "10%" else 0.08

st.caption(f"請求書の件名: 令和{reiwa_now}年{billing_month}{billing_subject}　消費税: {tax_rate_label}")

with st.expander("任意項目（納期・支払条件・有効期限）"):
    oc1, oc2, oc3 = st.columns(3)
    with oc1:
        payment_due = st.text_input("納期", value="", placeholder="例：納品後即日")
    with oc2:
        payment_terms = st.text_input("支払条件", value="", placeholder="例：月末締め翌月末払い")
    with oc3:
        valid_until = st.text_input("有効期限", value="", placeholder="例：2025/12/31")

st.write("")

# ── STEP 1: テンプレートをダウンロード ──
st.markdown("""
<div class="step-card">
  <div class="step-header">
    <span class="step-num">STEP 1</span>
    <span class="step-title">テンプレートをダウンロード</span>
  </div>
  <p class="step-desc">
    はじめて利用する方は、まずテンプレートをダウンロードしてください。<br>
    黄色のセルに入力して保存したファイルを STEP 2 でアップロードします。
  </p>
</div>
""", unsafe_allow_html=True)

st.download_button(
    "📥 納品書テンプレートをダウンロード",
    data=load_nouhin_template_bytes(),
    file_name="納品書テンプレート.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

# 記入ルール
with st.expander("📋 テンプレートの記入ルールを確認する"):
    st.markdown("""
<table class="format-table">
  <thead>
    <tr>
      <th>セル</th>
      <th>入力内容</th>
      <th>備考</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td class="cell-ref">B3</td>
      <td>店舗名 <span class="required-badge">必須</span></td>
      <td>例：○○商店</td>
    </tr>
    <tr>
      <td class="cell-ref">B4</td>
      <td>納品日 <span class="required-badge">必須</span></td>
      <td>例：2025/06/15</td>
    </tr>
    <tr>
      <td class="cell-ref">B10以降</td>
      <td>商品名 <span class="required-badge">必須</span></td>
      <td>1行に1商品。商品名が空の行で読み込み終了（最大16行）</td>
    </tr>
    <tr>
      <td class="cell-ref">C10以降</td>
      <td>数量 <span class="required-badge">必須</span></td>
      <td>商品名と同じ行に半角数字で入力</td>
    </tr>
    <tr>
      <td class="cell-ref">D10以降</td>
      <td>単価 <span class="required-badge">必須</span></td>
      <td>商品名と同じ行に半角数字で入力（税抜）</td>
    </tr>
    <tr>
      <td class="cell-ref">E10以降</td>
      <td>金額（自動）</td>
      <td>数量×単価で自動計算されます。入力不要</td>
    </tr>
  </tbody>
</table>
""", unsafe_allow_html=True)
    st.caption("※ 単価は利用者が手入力してください。商品名が空になった行で読み込みを終了します。")

st.write("")

# ── STEP 2: 納品書をアップロード ──
st.markdown("""
<div class="step-card">
  <div class="step-header">
    <span class="step-num">STEP 2</span>
    <span class="step-title">納品書をアップロード</span>
  </div>
  <p class="step-desc">
    テンプレートに入力・保存した Excel ファイルを選択してください。<br>
    複数の納品書を同時にアップロードできます（店舗ごとに請求書を生成します）。
  </p>
</div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "納品書ファイル（.xlsx）を選択",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

if uploaded_files:
    st.caption(f"✓ {len(uploaded_files)} 件のファイルを選択中")

st.write("")

# ── STEP 3: 請求書を生成 ──
st.markdown("""
<div class="step-card">
  <div class="step-header">
    <span class="step-num">STEP 3</span>
    <span class="step-title">請求書を生成してダウンロード</span>
  </div>
  <p class="step-desc">
    アップロードした納品書を読み取り、店舗ごとに請求書Excelを生成します。
  </p>
</div>
""", unsafe_allow_html=True)

if st.button("請求書を生成する", type="primary", use_container_width=True):
    if not uploaded_files:
        st.warning("納品書ファイルをアップロードしてください")
        st.stop()

    store_data    = {}
    errors        = []
    date_warnings = []  # 納品日未入力

    progress = st.progress(0, text="読み取り中...")
    for i, f in enumerate(uploaded_files):
        progress.progress((i + 1) / len(uploaded_files), text=f"読み取り: {f.name}")
        result = read_nouhinshо(f.read())
        if result is None or "error" in result:
            msg = result["error"] if isinstance(result, dict) else "読み取れませんでした"
            errors.append(f"**{f.name}**: {msg}")
            continue

        for w in result.get("warnings", []):
            date_warnings.append(f"{f.name}: {w}")

        store = result["store_name"]
        if store not in store_data:
            store_data[store] = {"dated_groups": []}
        store_data[store]["dated_groups"].append({
            "date":  result["delivery_date"],
            "items": result["items"],
        })
    progress.empty()

    if date_warnings:
        for w in date_warnings:
            st.warning(f"⚠️ {w}")

    if errors:
        for e in errors:
            st.error(f"❌ {e}")
        st.info("💡 STEP 1 からテンプレートをダウンロードして、黄色のセルに入力してからアップロードしてください。")
    if not store_data:
        st.error("データを抽出できませんでした")
        st.stop()

    st.write("")
    st.markdown('<p class="section-label">内容確認</p>', unsafe_allow_html=True)
    st.caption("金額に誤りがないか確認してからダウンロードしてください")

    all_ok = True
    for store_name, data in store_data.items():
        # 全アイテムを合算（金額確認・合計表示用）
        all_items_flat = [item for g in data["dated_groups"] for item in g["items"]]
        merged_all = {}
        for item in all_items_flat:
            k = item["name"]
            if k in merged_all:
                merged_all[k]["quantity"] += item["quantity"]
                merged_all[k]["amount"]   += item["amount"]
            else:
                merged_all[k] = dict(item)
        merged_items = list(merged_all.values())

        subtotal = sum(i["amount"] for i in merged_items)
        tax      = int(subtotal * tax_rate)
        total    = subtotal + tax

        # 単価×数量の検証
        check_errors = [
            f"{i['name']}: 計算値¥{i['unit_price']*i['quantity']:,} ≠ ¥{i['amount']:,}"
            for i in merged_items if i["unit_price"] * i["quantity"] != i["amount"]
        ]

        status = "-- 確認済み" if not check_errors else "-- 要確認"
        with st.expander(f"{store_name}　{status}　合計 ¥{total:,}（税込）", expanded=True):
            if check_errors:
                st.error("金額の不一致が検出されました: " + " / ".join(check_errors))
                all_ok = False

            # 日別明細を表示
            for group in data["dated_groups"]:
                d = group["date"]
                label = f"{d.month}月{d.day}日" if isinstance(d, date) else "日付不明"
                st.caption(f"**{label}**")
                df = pd.DataFrame(group["items"])[["name", "unit_price", "quantity", "amount"]]
                df.columns = ["商品名", "単価", "数量", "金額"]
                df["単価"] = df["単価"].map(lambda x: f"¥{x:,}")
                df["金額"] = df["金額"].map(lambda x: f"¥{x:,}")
                st.dataframe(df, use_container_width=True, hide_index=True)

            c1, c2, c3 = st.columns(3)
            c1.metric("小計（税抜）", f"¥{subtotal:,}")
            c2.metric(f"消費税（{tax_rate_label}）", f"¥{tax:,}")
            c3.metric("合計（税込）", f"¥{total:,}")

    st.write("")

    if not all_ok:
        st.error("金額の不一致があります。納品書ファイルを確認してください。")
    else:
        zip_buf = io.BytesIO()
        verify_failed = []
        with zipfile.ZipFile(zip_buf, "w") as zf:
            for store_name, data in store_data.items():
                # 納品日: 複数ファイルある場合は最初のものを使用
                first_date = data["dated_groups"][0]["date"] if data["dated_groups"] else None
                excel_bytes = create_invoice(
                    store_name,
                    data["dated_groups"],
                    billing_month,
                    billing_subject,
                    tax_rate=tax_rate,
                    payment_due=payment_due,
                    payment_terms=payment_terms,
                    valid_until=valid_until,
                    delivery_date=first_date,
                )
                # 生成後に書き込み値を再検証
                check = verify_invoice_output(excel_bytes, data["dated_groups"])
                if not check["ok"]:
                    verify_failed.extend([f"{store_name}: {e}" for e in check["errors"]])
                    continue  # 壊れたファイルはZIPに含めない

                safe = store_name.replace("/", "_").replace(" ", "_").replace("　", "_")
                zf.writestr(f"令和{date.today().year-2018}年{billing_month}請求書（{safe}）.xlsx", excel_bytes)

        if verify_failed:
            st.error("Excel書き込み検証でエラーが見つかりました。ダウンロードを中止しました。")
            for e in verify_failed:
                st.error(f"　{e}")
            st.stop()

        zip_buf.seek(0)
        st.download_button(
            label=f"請求書をダウンロード（{len(store_data)}店舗分 / ZIP）",
            data=zip_buf,
            file_name=f"請求書一括_{billing_month}.zip",
            mime="application/zip",
            use_container_width=True,
            type="primary",
        )
