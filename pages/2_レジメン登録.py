import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import json
from datetime import date, datetime

st.title("💊 レジメン登録")
st.caption("抽出したJSONを スプレッドシートに 登録")
st.divider()

st.markdown("#### 📋 JSONの取得はこちら　→　[レジメン情報抽出ページへ](./レジメン情報抽出)")
st.divider()

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

# ===== キー名ぶれを吸収するヘルパー関数 =====
def get_val(d, *keys, default=""):
    """複数のキー名候補から値を取得する"""
    for key in keys:
        if key in d and d[key] is not None:
            return d[key]
    return default

def get_drugs(data):
    """drug_info / drugs どちらでも取得できる"""
    return data.get("drug_info") or data.get("drugs") or []

def get_basic(data):
    """basic_info を取得"""
    return data.get("basic_info") or {}

# ===== JSON貼り付けエリア =====
json_text = st.text_area(
    "JSONをここに貼り付けてください",
    height=300,
    placeholder='{ "basic_info": { ... }, "drug_info": [ ... ] }'
)

if st.button("👁 内容を確認", type="primary"):
    if not json_text.strip():
        st.error("JSONを貼り付けてください")
    else:
        try:
            data = json.loads(json_text)
            info = get_basic(data)
            drugs = get_drugs(data)

            # 必須チェック
            protocol_no = get_val(info, "protocol_no")
            regimen_name = get_val(info, "regimen_name")
            disease = get_val(
                info,
                "disease",          # アプリのキー名
                "target_disease",   # LMがぶれたとき
                "対象疾患",
                default=""
            )
            course_days = get_val(info, "course_days", default="要確認")

            if not protocol_no:
                st.error("❌ basic_info に 'protocol_no' がありません。JSONを確認してください。")
                st.stop()
            if not regimen_name:
                st.error("❌ basic_info に 'regimen_name' がありません。JSONを確認してください。")
                st.stop()
            if not drugs:
                st.error("❌ 薬剤情報（drug_info）が見つかりません。JSONを確認してください。")
                st.stop()

            # セッションに保存（正規化済みデータ）
            st.session_state["parsed_data"] = data
            st.session_state["json_text"] = json_text
            st.session_state["protocol_no"] = protocol_no
            st.session_state["regimen_name"] = regimen_name
            st.session_state["disease"] = disease
            st.session_state["course_days"] = course_days

            # 基本情報表示
            st.subheader("📋 基本情報")
            col1, col2 = st.columns(2)
            with col1:
                st.write(f"**プロトコールNo：** {protocol_no}")
                st.write(f"**レジメン名：** {regimen_name}")
            with col2:
                st.write(f"**対象疾患：** {disease or '（未取得）'}")
                st.write(f"**1コース日数：** {course_days}")

            # 薬剤情報表示
            st.subheader("💊 薬剤情報")
            import pandas as pd
            drug_list = []
            for drug in drugs:
                drug_list.append({
                    "順序": get_val(drug, "order"),
                    "商品名": get_val(drug, "product_name", "brand_name"),
                    "投与量": f"{get_val(drug, 'dosage_value', 'dose_value')}"
                              f"{get_val(drug, 'dosage_unit', 'dose_unit')}",
                    "用量根拠": get_val(drug, "dosage_basis", "dose_basis"),
                    "Day": get_val(drug, "admin_day_text", "day_text"),
                    "管理コード": get_val(drug, "management_code"),
                })
            st.dataframe(pd.DataFrame(drug_list), use_container_width=True)

            # 重複チェック
            sh = get_spreadsheet()
            ws_basic = sh.worksheet("基本情報")
            existing = ws_basic.col_values(1)

            if protocol_no in existing:
                st.warning(f"⚠️ {protocol_no} はすでに登録されています。上書きしますか？")
                st.session_state["is_duplicate"] = True
            else:
                st.success(f"✅ {protocol_no} は新規登録です。")
                st.session_state["is_duplicate"] = False

        except json.JSONDecodeError as e:
            st.error(f"❌ JSONの形式が正しくありません: {e}")
        except Exception as e:
            st.error(f"❌ エラーが発生しました: {e}")

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
            info = get_basic(data)
            drugs = get_drugs(data)
            sh = get_spreadsheet()
            today = date.today().strftime("%Y/%-m/%-d")
            now = datetime.now().strftime("%Y/%-m/%-d %H:%M")

            protocol_no   = st.session_state["protocol_no"]
            regimen_name  = st.session_state["regimen_name"]
            disease       = st.session_state["disease"]
            course_days   = st.session_state["course_days"]

            ws_basic = sh.worksheet("基本情報")
            ws_drug  = sh.worksheet("薬剤情報")
            ws_log   = sh.worksheet("抽出ログ")

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
                    regimen_name,
                    disease,
                    "",
                    course_days if course_days else "要確認",
                    "", "", "", "",
                    today,
                    "",
                ]
                ws_basic.append_row(basic_row, value_input_option="USER_ENTERED")

                # 薬剤情報書き込み（LMキー名・旧キー名どちらも対応）
                for drug in drugs:
                    drug_row = [
                        protocol_no,
                        get_val(drug, "order"),
                        get_val(drug, "management_code"),
                        get_val(drug, "product_name",        "brand_name"),
                        get_val(drug, "dosage_value",        "dose_value"),
                        get_val(drug, "dosage_unit",         "dose_unit"),
                        get_val(drug, "dosage_basis",        "dose_basis"),
                        get_val(drug, "admin_day_text",      "day_text"),
                        get_val(drug, "admin_day_numeric",   "day_numbers"),
                        get_val(drug, "admin_timing",        "timing"),
                        get_val(drug, "diluent_volume",      "diluent_ml"),
                        get_val(drug, "admin_time_text",     "infusion_time_text"),
                        get_val(drug, "admin_time_numeric",  "infusion_time_hours"),
                        get_val(drug, "anticancer_flag",     "flag_O_chemo"),
                        get_val(drug, "support_flag",        "flag_O_support"),
                        get_val(drug, "seal_flag",           "flag_seal"),
                        get_val(drug, "figure_flag",         "flag_chart"),
                        get_val(drug, "manual_flag",         "flag_leaflet"),
                        get_val(drug, "remarks",             "note"),
                    ]
                    ws_drug.append_row(drug_row, value_input_option="USER_ENTERED")

                # ===== Pdカテゴリ自動設定 =====
                try:
                    ws_ae     = sh.worksheet("抗がん剤副作用マスタ")
                    ws_pd_sh  = sh.worksheet("Pd")
                    ae_data   = ws_ae.get_all_records()
                    pd_data   = ws_pd_sh.get_all_records()

                    # トリガー対応表作成
                    trigger_to_pdid = {}
                    code_to_pdid    = {}
                    priority_dict   = {}
                    for p in pd_data:
                        trigger = str(p.get('トリガーキーワード', '')).strip()
                        cat_id  = str(p.get('カテゴリID', '')).strip()
                        try:
                            priority_dict[cat_id] = int(p.get('優先順位', 99))
                        except:
                            priority_dict[cat_id] = 99
                        if not trigger or trigger == '手動設定':
                            continue
                        if trigger.startswith('AC'):
                            for c in trigger.split('|'):
                                code_to_pdid[c.strip()] = cat_id
                        else:
                            trigger_to_pdid[trigger] = cat_id

                    ae_dict    = {str(r['管理コード']).strip(): r for r in ae_data}
                    ae_headers = ws_ae.row_values(1)
                    ae_columns = ae_headers[2:]

                    # 薬剤コードから副作用→PdカテゴリIDを収集
                    collected = set()
                    for drug in drugs:
                        drug_code = str(get_val(drug, 'management_code')).strip()
                        ae_row = ae_dict.get(drug_code)
                        if ae_row:
                            for col in ae_columns:
                                if str(ae_row.get(col, '')).strip() == '○':
                                    pd_id = trigger_to_pdid.get(col)
                                    if pd_id:
                                        collected.add(pd_id)
                        if drug_code in code_to_pdid:
                            collected.add(code_to_pdid[drug_code])

                    # 優先順位でソート
                    sorted_ids = sorted(
                        collected,
                        key=lambda x: priority_dict.get(x, 99)
                    )
                    pd_category = '|'.join(sorted_ids)

                    # 基本情報L列（12列目）に書き込み
                    if pd_category:
                        existing_basic = ws_basic.col_values(1)
                        if protocol_no in existing_basic:
                            basic_row_idx = existing_basic.index(protocol_no) + 1
                            ws_basic.update_cell(basic_row_idx, 12, pd_category)

                except Exception as pd_e:
                    st.warning(f"⚠️ Pdカテゴリ自動設定でエラー: {pd_e}")

                # ログ記録
                operation = "上書き登録" if overwrite else "新規登録"
                log_row = [
                    now,
                    protocol_no,
                    regimen_name,
                    operation,
                    len(drugs),
                ]
                ws_log.append_row(log_row, value_input_option="USER_ENTERED")

            st.success(f"✅ {protocol_no} を{operation}しました！")
            st.balloons()
            # ===== テンプレート生成ページへのボタン =====
            st.divider()
            col1, col2 = st.columns(2)
            with col1:
                st.info("続けてテンプレートを生成しますか？")
            with col2:
                if st.button("📊 テンプレート生成ページへ", type="primary"):
                    st.session_state.clear()
                    st.switch_page("pages/3_テンプレート生成.py")
            if st.button("🔄 続けて別のレジメンを登録する"):
                st.session_state.clear()
                st.rerun()
            

        except Exception as e:
            st.error(f"❌ エラーが発生しました: {e}")
