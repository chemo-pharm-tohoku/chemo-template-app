import streamlit as st

st.set_page_config(
    page_title="ケモテンプレートシステム",
    page_icon="💊",
    layout="centered",
    initial_sidebar_state="expanded",
    menu_items={}
)

st.sidebar.title("メニュー")

# ===== タイトル =====
st.title("💊 ケモテンプレートシステム")
st.subheader("東北大学病院 薬剤部")
st.divider()

# ===== 新規レジメン登録 =====
st.subheader("🆕 新規レジメン登録")

with st.container(border=True):
    st.markdown("#### 📋 レジメン情報抽出・登録")
    st.write(
        "レジメンPDFをアップロードすると "
        "AIが自動解析 → スプレッドシート登録 "
        "→ スケジュールシール生成まで一気通貫"
    )
    if st.button(
        "📋 レジメン情報抽出・登録へ",
        type="primary",
        use_container_width=True,
        key="btn_top_1"
    ):
        st.switch_page("pages/1_レジメン情報抽出・登録.py")

st.divider()

# ===== テンプレート生成 =====
st.subheader("📄 テンプレート生成")

col_l, col_r = st.columns(2)

with col_l:
    with st.container(border=True):
        st.markdown("#### 📊 テンプレート Excel 生成")
        st.write(
            "登録済みレジメンから Excel、またはExcelに貼り付けるテキスト を生成。\n"
            "Pd（説明した内容）カテゴリ確認・副作用登録もこちら。"
        )
        if st.button(
            "📊 Excel 生成へ",
            type="primary",
            use_container_width=True,
            key="btn_top_3"
        ):
            st.switch_page("pages/3_テンプレート生成.py")

with col_r:
    with st.container(border=True):
        st.markdown("#### 📋 個人テキスト生成")
        st.write(
            "登録済みレジメンからパラメータを入力し "
            "O欄・Pd欄のテキストを生成。\n"
            "コピーしてファーマロードに直接貼り付けできる。"
        )
        if st.button(
            "📋 テキスト生成へ",
            type="primary",
            use_container_width=True,
            key="btn_top_4"
        ):
            st.switch_page("pages/4_O欄Pd欄生成.py")

    st.link_button(
        "📋 登録済みレジメンを確認する（マスタスプレッドシート）",
        "https://docs.google.com/spreadsheets/d/1dLEUYSZlrIK1uHqEtEAfS1jSAPpXCIiAiAk_iaRuY-8/edit?gid=0#gid=0",
        use_container_width=True
    )

st.divider()

# ===== 管理・設定 =====
st.subheader("⚙️ 管理・設定")

with st.container(border=True):
    st.markdown("#### 🗂️ マスタ スプレッドシートを直接編集する")
    st.caption("※ 追加後はページを再読み込みすると反映されます")
    st.markdown("")

    # 基本情報
    col1, col2 = st.columns([1, 3])
    with col1:
        st.link_button(
            "基本情報",
            "https://docs.google.com/spreadsheets/d/1dLEUYSZlrIK1uHqEtEAfS1jSAPpXCIiAiAk_iaRuY-8/edit?gid=0#gid=0",
            use_container_width=True
        )
    with col2:
        st.caption(
            "Pd カテゴリは、Pd欄に記載する患者さんに説明した文章を設定します。"
            "（PdカテゴリIDは、"
            "[Pd シート](https://docs.google.com/spreadsheets/d/1dLEUYSZlrIK1uHqEtEAfS1jSAPpXCIiAiAk_iaRuY-8/edit?gid=224247887#gid=224247887)"
            " 参照）"
        )

    st.markdown("")

    # 薬剤情報
    col1, col2 = st.columns([1, 3])
    with col1:
        st.link_button(
            "薬剤情報",
            "https://docs.google.com/spreadsheets/d/1dLEUYSZlrIK1uHqEtEAfS1jSAPpXCIiAiAk_iaRuY-8/edit?gid=1881167826#gid=1881167826",
            use_container_width=True
        )
    with col2:
        st.caption(
            "①O欄_抗がん剤　①O欄_支持療法　②シール　③図　④説明書 の項目に "
            "○ をつけるとテンプレートで表現されます。"
        )

    st.markdown("")

    # Pd
    col1, col2 = st.columns([1, 3])
    with col1:
        st.link_button(
            "Pd",
            "https://docs.google.com/spreadsheets/d/1dLEUYSZlrIK1uHqEtEAfS1jSAPpXCIiAiAk_iaRuY-8/edit?gid=224247887#gid=224247887",
            use_container_width=True
        )
    with col2:
        st.caption(
            "患者さんに説明した副作用内容を編集できます。"
        )

    st.markdown("")

    # 抗がん剤副作用マスタ
    col1, col2 = st.columns([1, 3])
    with col1:
        st.link_button(
            "抗がん剤副作用マスタ",
            "https://docs.google.com/spreadsheets/d/1dLEUYSZlrIK1uHqEtEAfS1jSAPpXCIiAiAk_iaRuY-8/edit?gid=2028693062#gid=2028693062",
            use_container_width=True
        )
    with col2:
        st.caption(
            "テンプレート生成ページでレジメンの副作用を登録する際の "
            "チェックボックスが表示されます。"
        )

    st.markdown("")
    st.warning("⚠️ 本マスタを編集するとシステム全体に影響します。")

st.divider()

# ===== Pd説明文管理 =====
st.subheader("📝 Pd 説明文管理")

with st.container(border=True):
    st.markdown("#### 📝 Pd 説明文管理")
    st.write(
        "化学療法指導記録（Pd欄）に使用する "
        "説明文テンプレートを管理・登録。"
    )
    if st.button(
        "📝 Pd 説明文管理へ",
        use_container_width=True,
        key="btn_top_pd"
    ):
        st.switch_page("pages/5_Pd説明文管理.py")
