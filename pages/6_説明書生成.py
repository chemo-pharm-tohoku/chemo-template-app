import streamlit as st
import re
from collections import defaultdict
import gspread
from google.oauth2 import service_account

st.set_page_config(
    page_title="説明書生成",
    page_icon="📄",
    layout="centered",
    initial_sidebar_state="expanded",
    menu_items={}
)

st.sidebar.title("メニュー")

SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1dLEUYSZlrIK1uHqEtEAfS1jSAPpXCIiAiAk_iaRuY-8/edit"

@st.cache_resource
def get_gspread_client():
    creds = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=[
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
        ]
    )
    return gspread.authorize(creds)

@st.cache_data(ttl=60)
def fetch_sheet_realtime(sheet_name):
    try:
        gc = get_gspread_client()
        ss = gc.open_by_url(SPREADSHEET_URL)
        ws = ss.worksheet(sheet_name)
        return ws.get_all_records()
    except Exception as e:
        st.error(f"シート「{sheet_name}」の取得に失敗: {e}")
        return []

@st.cache_data(ttl=300)
def fetch_sheet(sheet_name):
    try:
        gc = get_gspread_client()
        ss = gc.open_by_url(SPREADSHEET_URL)
        ws = ss.worksheet(sheet_name)
        return ws.get_all_records()
    except Exception as e:
        st.error(f"シート「{sheet_name}」の取得に失敗: {e}")
        return []

@st.cache_data(ttl=60)
def load_all_data():
    basic_data  = fetch_sheet_realtime("基本情報")
    drug_data   = fetch_sheet_realtime("薬剤情報")
    master_data = fetch_sheet_realtime("薬品マスタ")
    notes_data  = fetch_sheet("注意事項")
    return basic_data, drug_data, master_data, notes_data


def parse_days_num(day_str):
    days = []
    for part in str(day_str).split('|'):
        part = part.strip()
        if '-' in part:
            try:
                s, e = part.split('-')
                days.extend(range(int(s), int(e)+1))
            except:
                pass
        elif part.isdigit():
            days.append(int(part))
    return sorted(set(days))


def shorten_regimen_name(regimen_name):
    name = regimen_name
    name = re.sub(r'[（(][^）)]*[）)]', '', name)
    for word in [
        '術前','術後','周術期','切除不能','再発','進行',
        '維持','補助','一次治療','二次治療','初回','難治性',
        '肺癌','胃癌','大腸癌','乳癌','膵癌','肝癌',
        '食道癌','子宮癌','卵巣癌','前立腺癌','膀胱癌',
        '腎癌','甲状腺癌','悪性リンパ腫','白血病',
        '骨髄腫','中皮腫','胸腺癌','胸腺腫','神経内分泌腫瘍',
    ]:
        name = name.replace(word, '')
    name = re.sub(r'\s+', ' ', name).strip('　 ')
    return name


INJECTION_ORDER = {
    'NK1':1,'5HT3':2,'ステロイド':3,'G-CSF':4,
    '利尿薬':5,'解毒薬':6,'抗アレルギー':7,
    'H2ブロッカー':8,'電解質補正':9,'その他注射':10
}


def get_regimen(protocol_no, basic_data, drug_data, master_data):
    master_dict = {m['管理コード']: m for m in master_data}
    basic = [b for b in basic_data if b['プロトコールNo'] == protocol_no]
    if not basic:
        return None
    basic = basic[0]
    drugs_raw = sorted(
        [d for d in drug_data if d['プロトコールNo'] == protocol_no],
        key=lambda x: (
            0 if str(x['投与順序']).isdigit() else 1,
            int(x['投与順序']) if str(x['投与順序']).isdigit() else 99
        )
    )
    drugs = []
    for drug in drugs_raw:
        code   = str(drug['管理コード'])
        master = master_dict.get(code, {})
        merged = dict(drug)
        merged.update({
            '一般名（全角）'    : master.get('一般名（全角）',''),
            '採用商品名（全角）': master.get('採用商品名（全角）',''),
            '薬剤区分'         : master.get('薬剤区分',''),
            '支持療法分類'     : master.get('支持療法分類',''),
            '患者向け説明'     : master.get('患者向け説明',''),
        })
        drugs.append(merged)
    return {'basic': basic, 'drugs': drugs, 'master_dict': master_dict}


def get_rp_info(rp_drugs):
    rp_sorted = sorted(
        rp_drugs,
        key=lambda d: INJECTION_ORDER.get(str(d.get('支持療法分類','')), 99)
    )
    names = [str(d.get('商品名','') or d.get('採用商品名（全角）','')) for d in rp_sorted]
    name_text = '＋'.join([n for n in names if n])
    classes = set(); kubuns = set()
    for d in rp_sorted:
        kubuns.add(str(d.get('薬剤区分','')))
        cls = str(d.get('支持療法分類',''))
        if cls:
            classes.add(cls)
    if '抗がん剤' in kubuns:
        desc = '抗がん剤です。'
    else:
        has_nausea  = bool(classes & {'NK1','5HT3'})
        has_allergy = bool(classes & {'ステロイド','抗アレルギー'})
        if has_nausea and has_allergy: desc = '吐き気やアレルギーを抑えるお薬です。'
        elif has_nausea:               desc = '吐き気を抑えるお薬です。'
        elif has_allergy:              desc = 'アレルギーを抑えるお薬です。'
        elif 'G-CSF' in classes:      desc = '白血球を増やすお薬です。'
        elif '利尿薬' in classes:     desc = '尿の量を増やすお薬です。'
        elif '解毒薬' in classes:     desc = '副作用を和らげるお薬です。'
        else:                         desc = '点滴のお薬です。'
    time_text  = str(rp_sorted[0].get('投与時間文字',''))
    kubun_main = '抗がん剤' if '抗がん剤' in kubuns else '支持療法'
    return name_text, desc, time_text, kubun_main


def build_schedule_html(protocol_no, basic_data, drug_data, master_data, notes_data):
    result = get_regimen(protocol_no, basic_data, drug_data, master_data)
    if result is None:
        return None
    basic = result['basic']
    drugs = result['drugs']
    cycle = int(basic['1コース日数'])
    regimen_short = shorten_regimen_name(basic['レジメン名'])

    schedule_drugs = [d for d in drugs if str(d.get('④説明書','')) == '○']
    inj_schedule   = [d for d in schedule_drugs if str(d.get('投与順序','')) != '内服']
    oral_schedule  = [d for d in schedule_drugs if str(d.get('投与順序','')) == '内服']
    oral_cancer    = [d for d in drugs if str(d.get('①O欄_内服抗がん薬','')) == '○']

    rp_groups = defaultdict(list)
    for drug in inj_schedule:
        rp_groups[str(drug.get('投与順序',''))].append(drug)
    sorted_rps = sorted(rp_groups.keys(), key=lambda x: int(x) if x.isdigit() else 99)

    all_invest_days = set()
    for rp in sorted_rps:
        all_invest_days.update(parse_days_num(rp_groups[rp][0].get('投与Day数値','')))
    invest_days = sorted(all_invest_days)

    columns = []
    prev = 0
    for d in invest_days:
        if d > prev + 1:
            rs, re_ = prev + 1, d - 1
            columns.append({
                'type': '休薬中',
                'label': f'{rs}日目' if rs == re_ else f'{rs}〜{re_}日目',
                'days': list(range(rs, re_ + 1)),
            })
        columns.append({'type': '投与', 'label': f'{d}日目', 'days': [d]})
        prev = d
    if prev < cycle:
        rs, re_ = prev + 1, cycle
        columns.append({
            'type': '休薬',
            'label': f'{rs}日目' if rs == re_ else f'{rs}〜{re_}日目',
            'days': list(range(rs, re_ + 1)),
        })

    # ---- HTML組み立て ----
    html = []
    html.append(
        f"<div style='font-family:\"BIZ UDゴシック\",sans-serif; "
        f"width:190mm; border:2px solid #333; padding:8px;'>"
    )
    html.append(
        f"<div style='background:#2E4057;color:#fff;font-weight:bold;"
        f"font-size:16px;padding:6px 10px;'>【{protocol_no}】{regimen_short}</div>"
    )
    html.append(
        f"<div style='background:#4472C4;color:#fff;font-weight:bold;"
        f"font-size:12px;padding:4px 10px;'>1コース{cycle}日</div>"
    )
    html.append("<div style='font-weight:bold;font-size:13px;margin:8px 0 4px;'>◎治療スケジュール</div>")
    html.append("<div style='font-weight:bold;font-size:11px;margin-bottom:2px;'>＜注射＞</div>")

    html.append(
        "<table style='border-collapse:collapse;width:100%;font-size:10px;'>"
    )
    html.append("<tr>")
    for h in ['順序', '薬品名', '説明', '時間']:
        html.append(
            f"<th style='background:#2E4057;color:#fff;border:1px solid #999;"
            f"padding:3px;'>{h}</th>"
        )
    for col in columns:
        bg = '#2E4057' if col['type'] == '投与' else '#9E9E9E'
        html.append(
            f"<th style='background:{bg};color:#fff;border:1px solid #999;"
            f"padding:3px;'>{col['label']}</th>"
        )
    html.append("</tr>")

    for rp in sorted_rps:
        rp_drugs = rp_groups[rp]
        name_text, desc, time_text, kubun_main = get_rp_info(rp_drugs)
        row_bg = '#FFFDE7' if kubun_main == '抗がん剤' else '#D6EAF8'
        drug_days = parse_days_num(rp_drugs[0].get('投与Day数値', ''))

        html.append("<tr>")
        html.append(
            f"<td style='border:1px solid #999;padding:3px;text-align:center;"
            f"font-weight:bold;'>{rp}</td>"
        )
        html.append(
            f"<td style='background:{row_bg};border:1px solid #999;padding:3px;'>"
            f"{name_text}</td>"
        )
        html.append(
            f"<td style='background:{row_bg};border:1px solid #999;padding:3px;'>"
            f"{desc}</td>"
        )
        html.append(
            f"<td style='background:{row_bg};border:1px solid #999;padding:3px;"
            f"text-align:center;'>{time_text}</td>"
        )
        for col in columns:
            if col['type'] == '投与':
                hit = any(d in drug_days for d in col['days'])
                cell_bg = row_bg if hit else '#FAFAFA'
                mark = '●' if hit else '－'
                html.append(
                    f"<td style='background:{cell_bg};border:1px solid #999;"
                    f"padding:3px;text-align:center;'>{mark}</td>"
                )
            else:
                html.append(
                    f"<td style='background:#BFBFBF;border:1px solid #999;"
                    f"padding:3px;text-align:center;'></td>"
                )
        html.append("</tr>")
    html.append("</table>")

    if oral_schedule:
        html.append("<div style='font-weight:bold;font-size:11px;margin:8px 0 4px;'>＜内服＞</div>")
        for drug in oral_schedule:
            name = str(drug.get('商品名','') or drug.get('採用商品名（全角）',''))
            desc = str(drug.get('患者向け説明',''))
            timing = str(drug.get('投与タイミング',''))
            html.append(
                f"<div style='font-size:10px;margin-bottom:4px;'>"
                f"・<b>{name}</b>：{desc} {timing}</div>"
            )

    if oral_cancer:
        html.append("<div style='font-weight:bold;font-size:11px;margin:8px 0 4px;'>＜内服抗がん薬＞</div>")
        for drug in oral_cancer:
            code = str(drug.get('管理コード',''))
            master = result['master_dict'].get(code, {})
            name = str(master.get('一般名（全角）','') or drug.get('商品名',''))
            text_val = str(drug.get('投与量数値','') or '').strip()
            html.append(
                f"<div style='font-size:10px;margin-bottom:4px;'>"
                f"・<b>{name}</b>：{text_val}</div>"
            )

    common_notes = sorted(
        [n for n in notes_data if str(n.get('プロトコールNo','')) == '共通'],
        key=lambda x: int(x['順序']) if str(x['順序']).isdigit() else 99
    )
    if common_notes:
        html.append("<div style='font-weight:bold;font-size:13px;margin:10px 0 4px;'>◎注意事項</div>")
        for note in common_notes:
            html.append(
                f"<div style='font-size:10px;margin-bottom:3px;'>・{note['注意事項文章']}</div>"
            )

    html.append("</div>")
    return "".join(html)


# ===== Streamlit UI =====
st.title("📄 説明書生成")
st.caption("患者さん説明用のスケジュール表を画面に表示します。コピーしてWordに貼り付けてご利用ください。")
st.divider()

if st.button("🔄 データを最新化する", key="btn_refresh_6"):
    st.cache_data.clear()
    st.cache_resource.clear()
    st.rerun()

with st.spinner("データを読み込み中..."):
    basic_data, drug_data, master_data, notes_data = load_all_data()

if not basic_data:
    st.error("データの読み込みに失敗しました。")
    st.stop()

regimen_list = [
    f"{b['プロトコールNo']}　{b['レジメン名']}"
    for b in reversed(basic_data)
    if b.get('プロトコールNo', '').strip()
]

if not regimen_list:
    st.error("レジメン一覧が取得できませんでした。")
    st.stop()

selected = st.selectbox(
    "レジメンを選択してください",
    options=regimen_list,
    index=0,
    key="selectbox_schedule",
)

protocol_no = selected.split('　')[0].strip()

st.divider()
st.subheader("📋 プレビュー（下の枠内を選択してコピーし、Wordに貼り付けてください）")
st.caption(
    "💡 操作方法：枠内でクリック→Ctrl+A（全選択）→Ctrl+C（コピー）→"
    "Wordを開いてCtrl+V（貼り付け）。表の書式を保ったまま貼り付けられます。"
)

html_content = build_schedule_html(protocol_no, basic_data, drug_data, master_data, notes_data)

if html_content:
    st.markdown(
        f"<div style='border:2px dashed #4472C4; padding:10px; background:#fff;'>"
        f"{html_content}</div>",
        unsafe_allow_html=True,
    )
else:
    st.error("このレジメンのデータが取得できませんでした。")

st.divider()
st.caption("⚠️ 生成された内容は必ず確認してから使用してください")
