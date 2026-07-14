import streamlit as st

st.set_page_config(
    page_title="ケモテンプレートシステム",
    page_icon="🏠",
    layout="centered",
    initial_sidebar_state="expanded"
)

st.title("💊 ケモテンプレートシステム")
st.caption("東北大学病院 薬剤部")
st.divider()

st.markdown("""
### 使い方

**💊 レジメン登録**
- NotebookLMでレジメンPDFからJSONを抽出
- 抽出したJSONを貼り付けてスプレッドシートに登録

**📊 テンプレート生成**
- 登録済みレジメンからExcel・パワポを生成
""")

st.divider()
st.caption("⚠️ 生成されたファイルは必ず内容を確認してから使用してください")
