import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

st.title("📝 Pd説明文管理")
st.caption("化学療法指導記録（Pd欄）に使用する説明文テンプレートを管理します。")
st.caption("種別A（標準）・種別B（副作用発現時）・種別C（薬剤固有注意）の3種類で管理しています。")
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

@st.cache_data(ttl=60)
def load_pd_data():
    try:
        sh = get_spreadsheet()
        ws = sh.worksheet("Pd")
        return ws.get_all_records()
    except Exception as e:
        st.error(f"データの取得に失敗しました: {e}")
        return []

# ===== 種別をIDから取得 =====
def get_kind(item):
    cat_id = str(item.get("カテゴリID", ""))
    if len(cat_id) >= 3:
        return cat_id[2]  # PDA→A / PDB→B / PDC→C
    return str(item.get("種別", ""))

# ===== 次のIDを自動採番 =====
def get_next_id(pd_data, kind):
    prefix = f"PD{kind}"
    existing = [
        str(d.get("カテゴリID", ""))
        for d in pd_data
        if str(d.get("カテゴリID", "")).startswith(prefix)
    ]
    nums = []
    for id_str in existing:
        num_part = id_str[len(prefix):]
        if num_part.isdigit():
            nums.append(int(num_part))
    next_num = max(nums) + 1 if nums else 1
    return f"{prefix}{next_num:03d}"

# ===== 種別カラーの定義 =====
KIND_COLOR = {
    "A": "🟢",
    "B": "🟡",
    "C": "🔵",
}
KIND_LABEL = {
    "A": "A：標準説明文",
    "B": "B：副作用発現時",
    "C": "C：薬剤固有注意",
}
KIND_OPTIONS = {
    "A：標準説明文":   "A",
    "B：副作用発現時": "B",
    "C：薬剤固有注意": "C",
}

# ===== メイン =====
st.markdown("""
| 種別 | 意味 |
|------|------|
| 🟢 A（PDA） | 毎回説明する標準説明文 |
| 🟡 B（PDB） | 副作用が出た患者だけに説明 |
| 🔵 C（PDC） | 特定薬剤に必ず伝える固有注意 |
""")
st.divider()

# ===== データ読み込み =====
with st.spinner("データを読み込み中..."):
    pd_data = load_pd_data()

if not pd_data:
    st.warning("Pdシートにデータがありません。")
    st.stop()

# ===== 新規追加ボタン =====
if st.button("➕ 新規説明文を追加する", type="primary"):
    st.session_state["show_add_form"] = True

# ===== 新規追加フォーム =====
if st.session_state.get("show_add_form", False):
    st.divider()
    st.subheader("➕ 新規説明文の追加")

    with st.form("add_pd_form"):

        # 種別選択（先に選ぶとIDが自動で決まる）
        kind_label = st.selectbox(
            "種別 *",
            options=list(KIND_OPTIONS.keys()),
        )
        kind_key = KIND_OPTIONS[kind_label]

        # 自動採番したIDを初期値として表示
        auto_id = get_next_id(pd_data, kind_key)
        cat_id = st.text_input(
            "カテゴリID *",
            value=auto_id,
            help="自動採番されます。変更も可能です。"
        )

        cat_name = st.text_input(
            "カテゴリ名 *",
            placeholder="例：irAE全般、下痢発現時"
        )

        trigger = st.text_input(
            "トリガーキーワード",
            placeholder="例：抗PD-1|抗PD-L1　（複数はパイプ区切り）",
            help="自動判定に使用するキーワード。手動設定の場合は空欄でも可。"
        )

        description = st.text_area(
            "説明文 *",
            height=150,
            placeholder="例：irAEは発現時期や重症度、発現部位などに個人差があります…"
        )

        priority = st.number_input(
            "優先順位",
            min_value=1,
            max_value=999,
            value=99,
            help="数値が小さいほど上に表示されます"
        )

        note = st.text_input(
            "備考",
            placeholder="任意"
        )

        col1, col2 = st.columns(2)
        with col1:
            submitted = st.form_submit_button(
                "✅ 登録する",
                type="primary",
                use_container_width=True
            )
        with col2:
            cancelled = st.form_submit_button(
                "❌ キャンセル",
                use_container_width=True
            )

    # ===== キャンセル =====
    if cancelled:
        st.session_state["show_add_form"] = False
        st.rerun()

    # ===== 登録処理 =====
    if submitted:
        # バリデーション
        errors = []
        if not cat_id.strip():
            errors.append("カテゴリIDを入力してください")
        if not cat_name.strip():
            errors.append("カテゴリ名を入力してください")
        if not description.strip():
            errors.append("説明文を入力してください")

        # ID重複チェック
        existing_ids = [str(d.get("カテゴリID", "")) for d in pd_data]
        if cat_id.strip() in existing_ids:
            errors.append(f"カテゴリID「{cat_id}」はすでに存在します")

        if errors:
            for e in errors:
                st.error(f"❌ {e}")
        else:
            try:
                sh = get_spreadsheet()
                ws = sh.worksheet("Pd")
                new_row = [
                    cat_id.strip(),
                    cat_name.strip(),
                    kind_key,
                    trigger.strip(),
                    description.strip(),
                    priority,
                    note.strip(),
                ]
                ws.append_row(new_row, value_input_option="USER_ENTERED")

                st.success(f"✅ 「{cat_name}」を登録しました！")
                st.session_state["show_add_form"] = False
                # キャッシュクリアして再読み込み
                load_pd_data.clear()
                st.rerun()

            except Exception as e:
                st.error(f"❌ 登録に失敗しました: {e}")

    st.divider()

# ===== 絞り込みフィルター =====
st.subheader("🔍 絞り込み")
col1, col2 = st.columns(2)
with col1:
    kind_options = ["すべて", "A：標準説明文", "B：副作用発現時", "C：薬剤固有注意"]
    selected_kind = st.selectbox("種別で絞り込み", kind_options)
with col2:
    search_word = st.text_input(
        "キーワード検索",
        placeholder="カテゴリ名・説明文で検索"
    )

# ===== フィルタリング =====
filtered = pd_data
if selected_kind != "すべて":
    kind_key = selected_kind[0]
    filtered = [d for d in filtered if get_kind(d) == kind_key]
if search_word:
    filtered = [
        d for d in filtered
        if search_word in str(d.get("カテゴリ名", ""))
        or search_word in str(d.get("説明文", ""))
    ]

st.divider()

# ===== 件数表示 =====
total = len(pd_data)
shown = len(filtered)
st.caption(f"全{total}件中　{shown}件表示")

# ===== Expander表示の共通関数 =====
def show_item(item):
    kind = get_kind(item)
    with st.expander(
        f"{KIND_COLOR.get(kind,'⚪')} "
        f"**{item.get('カテゴリID','')}**　"
        f"{item.get('カテゴリ名','')}"
    ):
        col1, col2 = st.columns([1, 3])
        with col1:
            st.markdown(f"**種別：** {KIND_LABEL.get(kind,'')}")
            st.markdown(f"**優先順位：** {item.get('優先順位','')}")
            st.markdown("**トリガー：**")
            triggers = str(item.get("トリガーキーワード", "")).split("|")
            for t in triggers:
                if t.strip():
                    st.markdown(f"　・{t.strip()}")
            if item.get("備考"):
                st.markdown(f"**備考：** {item.get('備考','')}")
        with col2:
            st.markdown("**説明文：**")
            st.info(item.get("説明文", "（説明文未登録）"))
            if st.button(
                "📋 コピー用テキスト表示",
                key=f"copy_{item.get('カテゴリID','')}"
            ):
                st.code(item.get("説明文", ""), language=None)

def sort_items(items):
    return sorted(
        items,
        key=lambda x: int(x.get("優先順位", 99))
        if str(x.get("優先順位", "")).isdigit() else 99
    )

# ===== 種別ごとにタブ表示 =====
if selected_kind == "すべて" and not search_word:
    tab_a, tab_b, tab_c = st.tabs([
        "🟢 A：標準説明文",
        "🟡 B：副作用発現時",
        "🔵 C：薬剤固有注意",
    ])
    tabs = {"A": tab_a, "B": tab_b, "C": tab_c}
    for kind, tab in tabs.items():
        kind_data = [d for d in pd_data if get_kind(d) == kind]
        with tab:
            if not kind_data:
                st.info("登録データがありません")
                continue
            for item in sort_items(kind_data):
                show_item(item)
else:
    if not filtered:
        st.warning("該当するデータがありません")
    else:
        for item in sort_items(filtered):
            show_item(item)

st.divider()
PD_SHEET_URL = f"{st.secrets['spreadsheet']['url']}#gid=224247887"
st.markdown(f"[📊 スプレッドシートを直接編集する（Pdシート）]({PD_SHEET_URL})")
st.caption("※ 追加後はページを再読み込みすると反映されます")
