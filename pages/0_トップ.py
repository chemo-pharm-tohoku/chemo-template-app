import streamlit as st

st.title("💊 ケモテンプレートシステム")
st.caption("東北大学病院 薬剤部")
st.divider()

st.markdown("### このシステムでできること")

col1, col2, col3 = st.columns(3)

with col1:
    st.success("""
**📋 レジメン情報抽出**

NotebookLMを使って
レジメンPDFから
データを自動抽出
""")
    st.page_link(
        "pages/1_レジメン情報抽出.py",
        label="📋 レジメン情報抽出へ",
        icon="📋"
    )

with col2:
    st.info("""
**💊 レジメン登録**

抽出したJSONを
スプレッドシートに
登録
""")
    st.page_link(
        "pages/2_レジメン登録.py",
        label="💊 レジメン登録へ",
        icon="💊"
    )

with col3:
    st.warning("""
**📊 テンプレート生成**

登録済みレジメンから
Excel・パワポを
生成
""")
    st.page_link(
        "pages/3_テンプレート生成.py",
        label="📊 テンプレート生成へ",
        icon="📊"
    )

st.divider()

st.markdown("### その他")

col4, col5 = st.columns(2)

with col4:
    st.info("""
**📝 Pd説明文管理**

化学療法指導記録（Pd欄）に
使用する説明文テンプレートを
管理・登録
""")
    st.page_link(
        "pages/4_Pd説明文管理.py",
        label="📝 Pd説明文管理へ",
        icon="📝"
    )

with col5:
    st.empty()

st.divider()
st.caption("⚠️ 生成されたファイルは必ず内容を確認してから使用してください")
