import streamlit as st
from google import genai
from google.genai import types
import gspread
from google.oauth2.service_account import Credentials
import json, re

st.title("📋 レジメン情報抽出")
st.caption("レジメンPDFをアップロードするとAIが自動でJSONを抽出します")
st.divider()

@st.cache_resource
def get_gemini_client():
    api_key = st.secrets["gemini"]["api_key"]
    return genai.Client(api_key=api_key)

@st.cache_resource
def get_gspread_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=scopes
    )
    return gspread.authorize(creds)

@st.cache_data(ttl=600)
def load_definition():
    gc = get_gspread_client()
    SPREADSHEET_ID = "1dLEUYSZlrIK1uHqEtEAfS1jSAPpXCIiAiAk_iaRuY-8"
    sh = gc.open_by_key(SPREADSHEET_ID)
    ws = sh.worksheet("抽出定義書")
    return ws.acell("A1").value

st.subheader("STEP 1　レジメンPDFをアップロード")

uploaded = st.file_uploader(
    "PDFファイルをここにドロップ",
    type="pdf",
    help="レジメン情報PDFを1件アップロードしてください"
)

if uploaded:
    st.success(f"✅ {uploaded.name} を読み込みました")

st.divider()

st.subheader("STEP 2　AIが自動解析")

if uploaded:
    if st.button("🤖 自動解析スタート", type="primary", use_container_width=True):
        with st.spinner("AIが解析中です...少々お待ちください⏳"):
            try:
                definition = load_definition()
                pdf_bytes = uploaded.read()
                client = get_gemini_client()
                response = client.models.generate_content(
                    model="gemini-2.5-flash",
                    contents=[
                        definition,
                        types.Part.from_bytes(
                            mime_type="application/pdf",
                            data=pdf_bytes
                        )
                    ]
                )
                raw = response.text
                match = re.search(r"\{.*\}", raw, re.DOTALL)
                if match:
                    json_str = match.group()
                    parsed = json.loads(json_str)
                    st.session_state["extracted_json"] = json_str
                    st.session_state["extracted_parsed"] = parsed
                    st.success("✅ 解析完了！")
                else:
                    st.error("JSONの抽出に失敗しました。もう一度試してください。")
                    st.text(raw)
            except Exception as e:
                st.error(f"エラーが発生しました: {e}")
else:
    st.info("👆 まずPDFをアップロードしてください")

st.divider()

st.subheader("STEP 3　結果を確認して登録へ")

if "extracted_parsed" in st.session_state:
    parsed = st.session_state["extracted_parsed"]

    if "pattern_determination" in parsed:
        st.info(parsed["pattern_determination"])

    st.markdown("#### 📋 基本情報")
    basic = parsed.get("basic_info", {})
    col1, col2 = st.columns(2)
    with col1:
        st.write(f"**プロトコールNo：** {basic.get('protocol_no','')}")
        st.write(f"**レジメン名：** {basic.get('regimen_name','')}")
        st.write(f"**対象疾患：** {basic.get('disease','')}")
        st.write(f"**疾患分類：** {basic.get('disease_category','')}")
    with col2:
        st.write(f"**1コース日数：** {basic.get('course_days','')}")
        st.write(f"**催吐性リスク：** {basic.get('emetic_risk','')}")
        st.write(f"**備考：** {basic.get('remarks','')}")
        st.write(f"**登録日：** {basic.get('registration_date','')}")

    st.markdown("#### 💊 薬剤情報")
    drug_list = parsed.get("drug_info", [])
    if drug_list:
        import pandas as pd
        df = pd.DataFrame(drug_list)
        st.dataframe(df, use_container_width=True)

    with st.expander("🔧 JSONを直接編集する"):
        edited = st.text_area(
            "JSON",
            value=st.session_state["extracted_json"],
            height=400
        )
        if st.button("✅ 編集内容を反映"):
            try:
                st.session_state["extracted_json"] = edited
                st.session_state["extracted_parsed"] = json.loads(edited)
                st.success("反映しました！")
                st.rerun()
            except json.JSONDecodeError as e:
                st.error(f"JSONの形式が正しくありません: {e}")

    st.divider()
    st.success("内容を確認したら、レジメン登録ページへ進んでください")
    st.page_link("pages/2_レジメン登録.py", label="💊 レジメン登録ページへ", icon="💊")

else:
    st.info("👆 STEP2で解析すると結果がここに表示されます")
