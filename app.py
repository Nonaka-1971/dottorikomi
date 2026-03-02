# ✅ 必要なライブラリをインストール（最初の1回のみ）
# !pip install streamlit pandas openpyxl

import streamlit as st
import pandas as pd
from io import BytesIO
import re

# ===============================
# ✅ ヘッダー
# ===============================
st.set_page_config(page_title="📊 要素内訳一覧抽出ツール（仕訳日記帳から）", layout="centered")

st.title("📊 要素内訳一覧抽出ツール（仕訳日記帳から）")
st.markdown("仕訳日記帳（正規化・集計済）とテンプレートファイルをアップロードしてください。")

# ===============================
# ✅ ファイルアップロード
# ===============================
shiwake_file = st.file_uploader("📥 仕訳日記帳ファイル（正規化・集計済）をアップロード", type=["xlsx"])
template_file = st.file_uploader("📥 要素内訳一覧テンプレートファイル（空）をアップロード", type=["xlsx"])

def _safe_str_series(s: pd.Series) -> pd.Series:
    return s.astype("string").fillna("")

def _require_columns(df: pd.DataFrame, required: list[str]) -> list[str]:
    missing = [c for c in required if c not in df.columns]
    return missing

if shiwake_file and template_file:
    # ===============================
    # ✅ ファイル読み込み
    # ===============================
    df_all = pd.read_excel(shiwake_file)
    df_all.columns = df_all.columns.str.strip()

    df_template = pd.read_excel(template_file)
    df_template.columns = df_template.columns.str.strip()

    # ===============================
    # ✅ 必須列チェック（仕訳）
    # ===============================
    required_cols = [
        "要素内訳借方勘定科目名称",
        "要素内訳貸方勘定科目名称",
        "要素内訳借方勘定科目コード",
        "要素内訳貸方勘定科目コード",
    ]
    missing = _require_columns(df_all, required_cols)
    if missing:
        st.error(f"⚠️ 仕訳ファイルに必要な列がありません：{missing}")
        st.stop()

    # ===============================
    # ✅ モード判定（テンプレ名）
    # ===============================
    template_name = template_file.name
    if "原価" in template_name:
        mode = "原価"
    elif "支払" in template_name:
        mode = "支払"
    elif "売上" in template_name:
        mode = "売上"
    elif "入金" in template_name:
        mode = "入金"
    else:
        mode = "不明"

    st.success(f"🔍 判定結果 → 処理モード：{mode}")

    # 不明の場合は手動選択できるように（動作を止めない）
    if mode == "不明":
        mode = st.selectbox("⚙️ 処理モードを選択してください", ["原価", "支払", "売上", "入金"])
        st.info(f"✅ 手動選択 → 処理モード：{mode}")

    # ===============================
    # ✅ 抽出条件（ここが修正ポイント）
    # ===============================
    df_target = pd.DataFrame()

    # 原価：貸方が 工事未払金/買掛金/未払金 など（元のまま）
    if mode == "原価":
        cost_credit_accounts = ["工事未払金", "買掛金", "未払金"]
        df_target = df_all[df_all["要素内訳貸方勘定科目名称"].isin(cost_credit_accounts)].copy()
        st.info("✅ 原価モード：貸方が工事未払金/買掛金/未払金のデータを抽出します。")

    # 支払：借方が 工事未払金/買掛金/未払金 など（元のまま）
    elif mode == "支払":
        payment_debit_accounts = ["工事未払金", "買掛金", "未払金"]
        df_target = df_all[df_all["要素内訳借方勘定科目名称"].isin(payment_debit_accounts)].copy()
        st.info("✅ 支払モード：借方が工事未払金/買掛金/未払金のデータを抽出します。")

    # ✅ 売上：貸方が売上系（添付データでは「鉄骨工事売上」がここに入る）
    elif mode == "売上":
        # 売上系キーワード（必要なら追加）
        sales_pattern = r"(売上|完成工事高|役務収益|請負収益)"
        credit_name = _safe_str_series(df_all["要素内訳貸方勘定科目名称"])
        df_target = df_all[credit_name.str.contains(sales_pattern, regex=True, na=False)].copy()
        st.info("✅ 売上モード：貸方が売上系科目のデータを抽出します。")

    # ✅ 入金：借方が現金・預金系 AND 貸方が売掛・未収系（添付に「売掛金」「完成工事未収入金」あり）
    elif mode == "入金":
        debit_name = _safe_str_series(df_all["要素内訳借方勘定科目名称"])
        credit_name = _safe_str_series(df_all["要素内訳貸方勘定科目名称"])

        receipt_debit_pattern = r"(現金|預金)"  # 「普通預金」「定期預金」「個人預金」なども拾う
        ar_credit_candidates = ["売掛金", "完成工事未収入金", "未収入金", "工事未収金"]

        df_target = df_all[
            debit_name.str.contains(receipt_debit_pattern, regex=True, na=False)
            & credit_name.isin(ar_credit_candidates)
        ].copy()

        st.info("✅ 入金モード：借方が現金/預金、かつ貸方が売掛・未収系のデータを抽出します。")

    else:
        st.error("⚠️ 処理モードが不正です。")
        st.stop()

    # ===============================
    # ✅ 列構造に合わせて再構成し、ダウンロード可能に
    # ===============================
    if not df_target.empty:
        # テンプレ列に合わせて不足列を追加
        for col in df_template.columns:
            if col not in df_target.columns:
                df_target[col] = ""

        # テンプレ列順に並べ替え
        df_target = df_target[df_template.columns]

        output = BytesIO()
        df_target.to_excel(output, index=False)
        output.seek(0)

        st.success(f"✅ {mode} モードの抽出結果が {len(df_target)} 件あります。")
        st.download_button(
            label="📤 抽出結果をダウンロード",
            data=output,
            file_name=f"{mode}_要素内訳一覧_抽出結果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ 抽出対象のデータがありませんでした。抽出条件（科目名）を見直してください。")

