import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import json
from datetime import date, datetime

st.set_page_config(page_title="レジメン登録", page_icon="💊")
st.title("💊 レジメン登録")
st.caption("NotebookLMで抽出したJSONを貼り付けて登録してください")

# ===== 認証 =====
@st.cache_resource
def get_spreadsheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scopes
    )
    gc = gspread.authorize(creds)
    sh = gc.open_by_url(st.secrets["spreadsheet"]["url"])
    return sh

# ===== JSON貼り付けエリア =====
json_text = st.text_area(
    "JSONをここに貼り付けてください",
    height=300,
    placeholder='{ "basic_info": { ... }, "drugs": [ ... ] }'
)

if st.button("👁 内容を確認", type="primary"):
    if not json_text.strip():
        st.error("JSONを貼り付けてください")
    else:
        try:
            data = json.loads(json_text)
            info = data["basic_info"]

            # セッションに保存
            st.session_state["parsed_data"] = data
            st.session_state["json_text"] = json_text

            # 基本情報表示
            st.subheader("📋 基本情報")
            col1, col2 = st.columns(2)
            with col1:
                st.write(f"**プロトコールNo：** {info['protocol_no']}")
                st.write(f"**レジメン名：** {info['regimen_name']}")
            with col2:
                st.write(f"**対象疾患：** {info['disease']}")
                st.write(f"**1コース日数：** {info['course_days'] or '要確認'}")

            # 薬剤情報表示
            st.subheader("💊 薬剤情報")
            import pandas as pd
            drug_list = []
            for drug in data["drugs"]:
                drug_list.append({
                    "順序": drug.get("order", ""),
                    "商品名": drug.get("brand_name", ""),
                    "投与量": f"{drug.get('dose_value', '')}{drug.get('dose_unit', '')}",
                    "用量根拠": drug.get("dose_basis", ""),
                    "Day": drug.get("day_text", ""),
                    "管理コード": drug.get("management_code", ""),
                })
            st.dataframe(pd.DataFrame(drug_list), use_container_width=True)

            # 重複チェック
            sh = get_spreadsheet()
            ws_basic = sh.worksheet("基本情報")
            existing = ws_basic.col_values(1)
            protocol_no = info["protocol_no"]

            if protocol_no in existing:
                st.warning(f"⚠️ {protocol_no} はすでに登録されています。上書きしますか？")
                st.session_state["is_duplicate"] = True
            else:
                st.success(f"✅ {protocol_no} は新規登録です。")
                st.session_state["is_duplicate"] = False

        except json.JSONDecodeError as e:
            st.error(f"❌ JSONの形式が正しくありません: {e}")
        except KeyError as e:
            st.error(f"❌ 必要な項目が見つかりません: {e}")

# ===== 登録ボタン =====
if "parsed_data" in st.session_state:
    is_duplicate = st.session_state.get("is_duplicate", False)

    if is_duplicate:
        col1, col2 = st.columns(2)
        with col1:
            btn_cancel = st.button("❌ キャンセル")
        with col2:
            btn_overwrite = st.button("⚠️ 上書き登録", type="primary")
        if btn_cancel:
            st.session_state.clear()
            st.rerun()
        if btn_overwrite:
            do_register = True
            overwrite = True
        else:
            do_register = False
            overwrite = False
    else:
        btn_new = st.button("✅ 新規登録", type="primary")
        do_register = btn_new
        overwrite = False

    if do_register:
        try:
            data = st.session_state["parsed_data"]
            sh = get_spreadsheet()
            today = date.today().strftime("%Y/%-m/%-d")
            now = datetime.now().strftime("%Y/%-m/%-d %H:%M")
            protocol_no = data["basic_info"]["protocol_no"]
            ws_basic = sh.worksheet("基本情報")
            ws_drug = sh.worksheet("薬剤情報")
            ws_log = sh.worksheet("抽出ログ")

            with st.spinner("登録中..."):
                if overwrite:
                    # 基本情報削除
                    existing = ws_basic.col_values(1)
                    if protocol_no in existing:
                        row_idx = existing.index(protocol_no) + 1
                        ws_basic.delete_rows(row_idx)
                    # 薬剤情報削除
                    all_values = ws_drug.get_all_values()
                    rows_to_delete = [
                        i + 1 for i, row in enumerate(all_values)
                        if row and row[0] == protocol_no
                    ]
                    for row_idx in reversed(rows_to_delete):
                        ws_drug.delete_rows(row_idx)

                # 基本情報書き込み
                basic_row = [
                    protocol_no,
                    data["basic_info"]["regimen_name"],
                    data["basic_info"]["disease"],
                    "",
                    data["basic_info"]["course_days"] or "要確認",
                    "", "", "", "",
                    today,
                    "",
                ]
                ws_basic.append_row(basic_row, value_input_option="USER_ENTERED")

                # 薬剤情報書き込み
                for drug in data["drugs"]:
                    drug_row = [
                        protocol_no,
                        drug.get("order", ""),
                        drug.get("management_code", ""),
                        drug.get("brand_name", ""),
                        drug.get("dose_value", ""),
                        drug.get("dose_unit", ""),
                        drug.get("dose_basis", ""),
                        drug.get("day_text", ""),
                        drug.get("day_numbers", ""),
                        drug.get("timing", ""),
                        drug.get("diluent_ml", ""),
                        drug.get("infusion_time_text", ""),
                        drug.get("infusion_time_hours", ""),
                        drug.get("flag_O_chemo", ""),
                        drug.get("flag_O_support", ""),
                        drug.get("flag_seal", ""),
                        drug.get("flag_chart", ""),
                        drug.get("flag_leaflet", ""),
                        drug.get("note", ""),
                    ]
                    ws_drug.append_row(drug_row, value_input_option="USER_ENTERED")

                # ログ記録
                operation = "上書き登録" if overwrite else "新規登録"
                log_row = [
                    now,
                    protocol_no,
                    data["basic_info"]["regimen_name"],
                    operation,
                    len(data["drugs"]),
                ]
                ws_log.append_row(log_row, value_input_option="USER_ENTERED")

            st.success(f"✅ {protocol_no} を{operation}しました！")
            st.balloons()
            st.session_state.clear()

        except Exception as e:
            st.error(f"❌ エラーが発生しました: {e}")
