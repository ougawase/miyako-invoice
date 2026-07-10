import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import date, datetime
from invoice import (
    read_nouhinshо, create_invoice, verify_invoice_output,
    apply_price_overrides,
    PRICE_MASTER, TAX_RATE, ITEM_START, ITEM_END,
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
    /* Browse files ボタン */
    [data-testid="stFileUploader"] button {
        background-color: #2a2a2a !important;
        color: rgba(255,255,255,0.8) !important;
        border: 1px solid rgba(255,255,255,0.15) !important;
        border-radius: 6px !important;
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

col1, col2 = st.columns(2)
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

st.caption(f"請求書の件名: 令和{reiwa_now}年{billing_month}{billing_subject}")

st.write("")
st.markdown('<p class="section-label">納品書ファイル</p>', unsafe_allow_html=True)
uploaded_files = st.file_uploader(
    "納品書ファイル",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    label_visibility="collapsed"
)

if uploaded_files:
    st.caption(f"{len(uploaded_files)} 件のファイルを読み込みました")

st.write("")

if st.button("請求書を生成する", type="primary", use_container_width=True):
    if not uploaded_files:
        st.warning("納品書ファイルをアップロードしてください")
        st.stop()

    store_data = {}
    errors     = []

    progress = st.progress(0, text="読み取り中...")
    for i, f in enumerate(uploaded_files):
        progress.progress((i + 1) / len(uploaded_files), text=f"読み取り: {f.name}")
        result = read_nouhinshо(f.read())
        if result is None:
            errors.append(f.name)
            continue
        store = result["store_name"]
        if store not in store_data:
            store_data[store] = {"dated_groups": []}
        store_data[store]["dated_groups"].append({
            "date":  result["delivery_date"],
            "items": result["items"],
        })
    progress.empty()
    # 結果はセッションに保存し、ボタンの外で描画する。単価の入力・確認・
    # ダウンロードが再実行（入力のたびに起きる）をまたいで消えないようにするため。
    st.session_state["gen"] = {"store_data": store_data, "errors": errors}

gen = st.session_state.get("gen")
if gen:
    if gen["errors"]:
        st.warning(f"読み取れなかったファイル: {', '.join(gen['errors'])}")
    if not gen["store_data"]:
        st.error("データを抽出できませんでした")
        st.stop()

    # 価格表にない商品は単価0で読まれている。以前は1件でも全体を止めていたが、
    # 「その場で単価を入れて続行」に変更（2026-07-09。"ただ入れれば請求書になる"に戻す）。
    # 単価0のまま請求書を作らないことは、下の入力必須チェックで引き続き担保する。
    unknown_names = []
    for data in gen["store_data"].values():
        for g in data["dated_groups"]:
            for it in g["items"]:
                if it["unit_price"] == 0 and it["name"] not in unknown_names:
                    unknown_names.append(it["name"])

    overrides = {}
    if unknown_names:
        st.write("")
        st.markdown('<p class="section-label">単価の入力</p>', unsafe_allow_html=True)
        st.warning("価格表にない商品がありました。単価（税抜・円）を入れると、そのまま請求書を作成できます。")
        _pcols = st.columns(2)
        for _i, _nm in enumerate(unknown_names):
            with _pcols[_i % 2]:
                overrides[_nm] = int(st.number_input(
                    f"「{_nm}」の単価", min_value=0, step=10, key=f"pfx_{_nm}"))
        _missing = [nm for nm in unknown_names if not overrides.get(nm)]
        if _missing:
            st.info("単価が未入力：" + "、".join(_missing)
                    + "　（すべて入れると、内容確認とダウンロードが表示されます）")
            st.stop()

    store_data = (apply_price_overrides(gen["store_data"], overrides)
                  if unknown_names else gen["store_data"])

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
        tax      = int(subtotal * TAX_RATE)
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
            c2.metric("消費税（8%）", f"¥{tax:,}")
            c3.metric("合計（税込）", f"¥{total:,}")

    st.write("")

    if not all_ok:
        st.error("金額の不一致があります。納品書ファイルを確認してください。")
    else:
        zip_buf = io.BytesIO()
        verify_failed = []
        with zipfile.ZipFile(zip_buf, "w") as zf:
            for store_name, data in store_data.items():
                excel_bytes = create_invoice(
                    store_name,
                    data["dated_groups"],
                    billing_month,
                    billing_subject,
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
