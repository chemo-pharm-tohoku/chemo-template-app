import streamlit as st

st.title("💊 ケモテンプレートシステム")
st.caption("東北大学病院 薬剤部")
st.divider()

st.markdown("### 新規レジメン登録")

col1, col2 = st.columns(2)

with col1:
    st.success("""
**📋 レジメン情報抽出・登録**

レジメンPDFをアップロードすると
AIが自動解析→スプレッドシート登録
→スケジュールシール生成まで一気通貫
""")
    st.page_link(
        "pages/1_レジメン情報抽出・登録.py",
        label="📋 レジメン情報抽出・登録へ",
        icon="📋"
    )

with col2:
    st.warning("""
**📊 テンプレート生成**

登録済みレジメンから
Excel・パワポを生成
Pdカテゴリ確認・副作用登録もこちら
""")
    st.page_link(
        "pages/3_テンプレート生成.py",
        label="📊 テンプレート生成へ",
        icon="📊"
    )

st.divider()
st.markdown("### 管理・設定")

col3, col4 = st.columns(2)

with col3:
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

with col4:
    st.empty()

st.divider()

st.markdown("### 使い方")
st.markdown("""
1. **新規レジメン登録**は「📋 レジメン情報抽出・登録」から
2. **Excel・Pd欄生成**は「📊 テンプレート生成」から
3. **副作用マスタ登録**は「📊 テンプレート生成」の中で行えます
""")

st.divider()
st.caption("⚠️ 生成されたファイルは必ず内容を確認してから使用してください")
