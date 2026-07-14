import streamlit as st

st.title("📋 レジメン情報抽出")
st.caption("NotebookLMを使ってレジメンPDFからJSONを抽出します")
st.divider()

# ===== STEP1 =====
st.subheader("STEP 1　NotebookLMにソースを追加")
st.markdown("""
[📖 NotebookLMを開く](https://notebooklm.google.com/)

以下の**3つ**をNotebookLMのソースに追加してください。
""")

col1, col2, col3 = st.columns(3)
with col1:
    st.info("""
**① レジメン情報PDF**

登録したいレジメンの
PDFファイルを追加
""")
with col2:
    st.info("""
**② 薬品マスタ**

以下のリンクからPDFで
ダウンロードして追加

[薬品マスタを開く](https://docs.google.com/spreadsheets/d/1dLEUYSZlrIK1uHqEtEAfS1jSAPpXCIiAiAk_iaRuY-8/edit?gid=575053318#gid=575053318)
""")
with col3:
    st.info("""
**③ 抽出定義書**

下のボタンからダウンロード
して追加
""")
    with open("NotebookLM 抽出定義書260712.txt", "rb") as f:
        st.download_button(
            label="📥 抽出定義書をダウンロード",
            data=f,
            file_name="ケモテンプレート抽出定義書.txt",
            mime="text/plain"
        )
        
st.divider()

# ===== STEP2 =====
st.subheader("STEP 2　指示文をコピーしてNotebookLMに貼り付け")
st.warning("⚠️ 「プロトコールNo」の部分を実際のプロトコールNoに書き換えてください")

prompt = """「プロトコールNo」のPDFを定義書のルールに従って抽出してください。
パターン判定から始め、JSON形式で出力してください。
薬品マスタと照合して管理コードも記入してください。"""

st.code(prompt, language=None)

st.divider()

# ===== STEP3 =====
st.subheader("STEP 3　JSONをコピーしてレジメン登録へ")
st.markdown("""
1. NotebookLMの出力から `{` ～ `}` を**全部コピー**
2. **💊 レジメン登録**ページに貼り付けて登録
""")

st.page_link("pages/2_レジメン登録.py", label="💊 レジメン登録ページへ", icon="💊")
