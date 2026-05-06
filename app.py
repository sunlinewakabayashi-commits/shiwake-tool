"""
仕訳入力ツール - Streamlit ウェブアプリ版
"""
import streamlit as st
import os
import sys
import tempfile
import importlib.util

st.set_page_config(page_title="仕訳入力ツール", page_icon="📊", layout="wide")
st.title("📊 仕訳入力ツール")
st.caption("領収書PDFを読み取り、仕訳データExcelを作成します。")

# ========== APIキーの設定 ==========
if "ANTHROPIC_API_KEY" in st.secrets:
    os.environ["ANTHROPIC_API_KEY"] = st.secrets["ANTHROPIC_API_KEY"]
if not os.environ.get("ANTHROPIC_API_KEY"):
    st.error("⚠ ANTHROPIC_API_KEY が設定されていません。Streamlit の Secrets に登録してください。")
    st.stop()

# ========== 仕訳入力.py の関数を読み込む ==========
@st.cache_resource
def load_module():
    spec = importlib.util.spec_from_file_location("shiwake", "仕訳入力.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod

mod = load_module()

# ========== UI ==========
col1, col2 = st.columns(2)

with col1:
    st.subheader("① 会計ファイル（Excel）")
    accounting_upload = st.file_uploader(
        "あかしまい_損益確認 R8年度分.xlsx をアップロード",
        type=["xlsx"],
        help="最終残高・管理番号の取得に使用します"
    )

with col2:
    st.subheader("② 領収書ファイル")
    receipt_uploads = st.file_uploader(
        "領収書PDF・画像をアップロード（複数可）",
        type=["pdf", "jpg", "jpeg", "png"],
        accept_multiple_files=True,
        help="ファイル名は「会社名_月.pdf」形式にしてください（例: あかしまい_04.pdf）"
    )

# ========== 実行ボタン ==========
if st.button("🚀 仕訳データを作成", type="primary", disabled=not (accounting_upload and receipt_uploads)):

    with tempfile.TemporaryDirectory() as tmpdir:
        # 会計ファイルを保存
        accounting_path = os.path.join(tmpdir, accounting_upload.name)
        with open(accounting_path, "wb") as f:
            f.write(accounting_upload.getvalue())

        # 領収書ファイルを保存
        receipt_paths = []
        for ru in receipt_uploads:
            path = os.path.join(tmpdir, ru.name)
            with open(path, "wb") as f:
                f.write(ru.getvalue())
            receipt_paths.append(path)

        all_rows = []
        total_added = 0
        log_lines = []

        def ui_log(msg):
            log_lines.append(msg)

        for filepath in receipt_paths:
            ui_log(f"{'='*40}")
            try:
                company, month = mod.parse_filename(filepath)
            except ValueError as e:
                ui_log(f"スキップ: {e}")
                continue

            sheet_name = f"{company}{mod.SHEET_SUFFIX}_{month}"
            ui_log(f"会社: {company}  月: {month}月  → シート: {sheet_name}")

            last_balance, last_number = mod.get_last_state_readonly(accounting_path, sheet_name)
            ui_log(f"  最後の管理番号: {last_number}  残高: {last_balance:,}円")

            with st.spinner(f"{os.path.basename(filepath)} を読み取り中..."):
                try:
                    receipts = mod.read_receipts_from_file(filepath)
                except Exception as e:
                    ui_log(f"エラー: {e}")
                    continue

            # 税率分割・店舗カテゴリ適用
            expanded = []
            for r in receipts:
                for entry in mod.split_by_tax_rate(r):
                    expanded.append(mod.apply_store_category(entry))
            receipts = expanded

            # 重複除外
            seen = set()
            deduped = []
            for r in receipts:
                key = (r.get("日付",""), r.get("店名",""), int(r.get("合計金額",0)))
                if key not in seen:
                    seen.add(key)
                    deduped.append(r)
            if len(deduped) < len(receipts):
                ui_log(f"  重複 {len(receipts)-len(deduped)}件 を自動除外")
            receipts = deduped

            prev_receipt = None
            current_kanri = None
            for receipt in receipts:
                is_split = (prev_receipt is not None
                            and receipt.get("日付") == prev_receipt.get("日付")
                            and receipt.get("店名") == prev_receipt.get("店名"))
                if is_split:
                    kanri = current_kanri
                else:
                    kanri = mod.next_kanri_number(last_number, month)
                    current_kanri = kanri
                    last_number = kanri

                date_val = mod.date_str_to_excel(receipt.get("日付",""))
                total = int(receipt.get("合計金額", 0))
                tax = int(receipt.get("消費税", 0))
                subtotal = total - tax
                new_balance = last_balance - total
                billing = round(subtotal * mod.KAKERATE, 1)

                row = [kanri, None, company, date_val, receipt.get("カテゴリ",""),
                       receipt.get("店名",""), total, subtotal, None,
                       new_balance, mod.KAKERATE, billing, receipt.get("メモ","")]
                all_rows.append(row)

                ui_log(f"  → [{kanri}] {receipt.get('店名','')} {total:,}円 ({receipt.get('カテゴリ','')})")
                last_balance = new_balance
                prev_receipt = receipt
                total_added += 1

        # ========== 結果表示 ==========
        st.divider()
        if total_added == 0:
            st.warning("処理するデータがありませんでした。")
        else:
            st.success(f"✅ {total_added}件を処理しました")

            # テーブル表示
            import pandas as pd
            headers = ["管理番号","店舗","会社","日付","内容","取引先","税込価格","税抜価格","預入","残高","掛け率","請求金額","備考"]
            df = pd.DataFrame(all_rows, columns=headers)
            df["日付"] = pd.to_datetime(df["日付"] - 25569, unit="D", origin="unix").dt.strftime("%Y/%m/%d")
            st.dataframe(df, use_container_width=True)

            # Excel出力・ダウンロード
            output_path = os.path.join(tmpdir, f"仕訳データ_{os.path.splitext(receipt_uploads[0].name)[0]}.xlsx")
            mod.create_output_excel(all_rows, output_path)
            with open(output_path, "rb") as f:
                st.download_button(
                    label="📥 Excelをダウンロード",
                    data=f.read(),
                    file_name=os.path.basename(output_path),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        # ログ表示
        with st.expander("📋 処理ログ"):
            st.text("\n".join(log_lines))
