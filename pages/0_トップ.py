import streamlit as st
import pathlib

MANUAL_DIR = pathlib.Path(__file__).resolve().parent.parent / "manuals"

st.set_page_config(
    page_title="ケモテンプレートシステム",
    page_icon="💊",
    layout="centered",
    initial_sidebar_state="expanded",
    menu_items={}
)

st.sidebar.title("メニュー")

# ===== タイトル =====
st.title("💊 ケモテンプレート生成システム")
st.subheader("東北大学病院 薬剤部")
st.divider()
        

with st.container(border=True):
    st.markdown("#### 📋 レジメン情報抽出・登録")
    st.write(
        "確認票のテキストを貼り付けると "
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

col_l, col_r = st.columns(2)

with col_l:
    with st.container(border=True):
        st.markdown("#### 📊 テンプレート 生成")
        st.write(
            "登録済みレジメンから Excel、またはExcelに貼り付けるテキスト を生成。"
        )
        if st.button(
            "📊 テンプレート生成へ",
            type="primary",
            use_container_width=True,
            key="btn_top_3"
        ):
            st.switch_page("pages/3_テンプレート生成.py")

with col_r:
    with st.container(border=True):
        st.markdown("#### 📋 個人パラメーター入力生成")
        st.write(
            "登録済みレジメンからパラメータを入力し "
            "O欄・Pd欄のテキストを生成。\n"
            "コピーしてファーマロードに直接貼り付けできる。"
        )
        if st.button(
            "📋 パラメーター入力生成へ",
            type="primary",
            use_container_width=True,
            key="btn_top_4"
        ):
            st.switch_page("pages/4_O欄Pd欄生成.py")

col_l2, col_r2 = st.columns(2)

with col_l2:
    with st.container(border=True):
        st.markdown("#### 📄 説明書生成")
        st.write(
            "患者さん説明用のスケジュール表を画面に表示。\n"
            "ダウンロード不要、コピーしてWordに貼り付け編集・印刷できます。"
        )
        if st.button(
            "📄 説明書生成へ",
            type="primary",
            use_container_width=True,
            key="btn_top_6"
        ):
            st.switch_page("pages/6_説明書生成.py")

with col_r2:
    with st.container(border=True):
        st.markdown("#### 🧪 新薬メンテナンス")
        st.write(
            "新薬の薬品マスタ登録・抗がん剤副作用マスタ整備・"
            "Pd整合性チェックはこちらから。"
        )
        if st.button(
            "🧪 新薬メンテナンスへ",
            type="primary",
            use_container_width=True,
            key="btn_top_5"
        ):
            st.switch_page("pages/5_新薬メンテナンス.py")

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
            "プロトコールNo・レジメン名・1コース日数・備考等を管理します。"
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
            "①O欄_抗がん剤　①O欄_支持療法　①O欄_内服抗がん薬　②シール　③図　④説明書 の項目に "
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
            "患者さんに指導した副作用内容を編集できます。"
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
            "O欄（モニタリング項目条件表示）・Pd欄（説明事項）のON/OFFを制御します。"
        )
        with st.expander("※O欄の副作用項目について"):
            st.caption(
                "【常時表示】と【条件表示】に分けています。\n\n"
                "**常時表示**：嘔吐・悪心・食欲不振・便秘・倦怠感・骨髄抑制・"
                "肝機能障害・腎機能障害・電解質異常・その他（全レジメン共通）\n\n"
                "**条件表示**：下痢・口腔粘膜炎・脱毛・末梢神経障害・味覚異常・"
                "IRR・手足症候群・皮膚障害・間質性肺炎・心障害"
            )

    st.markdown("")
    st.warning("⚠️ 本マスタを編集するとシステム全体に影響します。")

# ===== マニュアル =====
st.subheader("📖 マニュアル")

with st.container(border=True):
    st.caption("使い方に迷ったら、まずはこちらをご確認ください")

    import urllib.parse

    GITHUB_OWNER = "chemo-pharm-tohoku"
    GITHUB_REPO  = "chemo-template-app"
    GITHUB_BRANCH = "main"

    manual_prefixes = [
        ("① ケモテンプレート生成アプリの使用方法", "ケモテンプレート生成アプリの使用方法"),
        ("② Pd欄・抗がん剤副作用マスタ 運用マニュアル", "Pd説明文メンテナンス方法"),
    ]

    for label, prefix in manual_prefixes:
        matched = sorted(MANUAL_DIR.glob(f"{prefix}*.pdf"), reverse=True)
        if matched:
            file_path = matched[0]  # ファイル名の末尾日付が新しいものを優先採用
            encoded_name = urllib.parse.quote(file_path.name)
            blob_url = (
                f"https://github.com/{GITHUB_OWNER}/{GITHUB_REPO}"
                f"/blob/{GITHUB_BRANCH}/manuals/{encoded_name}"
            )
            st.link_button(
                f"📄 {label}（{file_path.stem}）を開く",
                blob_url,
                use_container_width=True,
            )
        else:
            st.caption(f"⚠️ {label}（未配置：manuals/{prefix}*.pdf）")
st.divider()

# ===== システム改善要望 =====
st.subheader("📝 システム改善要望")

with st.container(border=True):
    st.write(
        "アプリの不具合・使いにくい点・追加してほしい機能などがあれば、"
        "以下のフォームからご連絡ください。"
    )
    st.link_button(
        "📝 システム改善要望フォームはこちら",
        "https://forms.gle/DNzaqpwMz9E1qBGz7",
        use_container_width=True,
    )
