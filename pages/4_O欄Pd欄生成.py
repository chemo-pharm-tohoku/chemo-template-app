import streamlit as st
import re
from datetime import date

st.set_page_config(
    page_title="O欄・Pd欄 テキスト生成",
    page_icon="📋",
    layout="centered",
    initial_sidebar_state="expanded",
    menu_items={}
)

st.sidebar.title("メニュー")

import gspread
from google.oauth2 import service_account

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
def fetch_sheet_realtime(sheet_name):
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
    pd_data     = fetch_sheet_realtime("Pd")
    ae_data     = fetch_sheet_realtime("抗がん剤副作用マスタ")
    return basic_data, drug_data, master_data, pd_data, ae_data


# ===== 投与量計算ヘルパー =====

def format_dose_text(drug):
    try:
        _rv = str(drug.get('投与量数値', '') or '').strip()
        _rv = ''.join(c for c in _rv if c.isdigit() or c == '.')
        dose = float(_rv or 0)
    except:
        dose = 0
    unit_input = str(drug.get('投与単位', '')).strip()
    dose_base  = str(drug.get('用量根拠', ''))
    v_to_mg    = drug.get('1V当たりmg', '')
    if unit_input.upper() == 'V':
        if v_to_mg != '' and str(v_to_mg).strip() != '':
            try:
                dose = dose * float(v_to_mg)
                unit_input = 'mg'
            except:
                pass
    if unit_input in ('mg', 'mg/body', 'mg/ body'):
        display_unit = 'mg'
    elif unit_input == '':
        display_unit = '' if dose_base == 'AUC依存' else 'mg'
    else:
        display_unit = unit_input
    dose_str = str(int(dose)) if dose == int(dose) else str(dose)
    return dose_str, display_unit


def calc_dose(drug, bsa, bw, ccr):
    """投与量を計算して文字列で返す。未入力・計算不能の場合は '未入力' を返す。"""
    dose_base = str(drug.get('用量根拠', ''))
    try:
        _rv = str(drug.get('投与量数値', '') or '').strip()
        _rv = ''.join(c for c in _rv if c.isdigit() or c == '.')
        dose_num = float(_rv or 0)
    except:
        dose_num = 0
    dose_str, unit_str = format_dose_text(drug)
    if dose_base == '固定用量':
        return f"{dose_str}{unit_str}"
    elif dose_base == 'BSA依存':
        if bsa is None:
            return '未入力'
        val = round(bsa * dose_num, 1)
        val_str = str(int(val)) if val == int(val) else str(val)
        return f"{val_str}mg"
    elif dose_base == 'AUC依存':
        if ccr is None:
            return '未入力'
        val = round((ccr + 25) * dose_num, 0)
        return f"{int(val)}mg"
    elif dose_base == 'BW依存':
        if bw is None:
            return '未入力'
        val = round(bw * dose_num, 1)
        val_str = str(int(val)) if val == int(val) else str(val)
        return f"{val_str}mg"
    else:
        return f"{dose_str}{unit_str}"


def calc_dose_num(drug, bsa, bw, ccr):
    """計算値を数値で返す（達成率計算用）。計算不能な場合はNoneを返す。"""
    dose_base = str(drug.get('用量根拠', ''))
    try:
        _rv = str(drug.get('投与量数値', '') or '').strip()
        _rv = ''.join(c for c in _rv if c.isdigit() or c == '.')
        dose_num = float(_rv or 0)
    except:
        dose_num = 0
    if dose_base == '固定用量':
        return dose_num
    elif dose_base == 'BSA依存':
        if bsa is None:
            return None
        return round(bsa * dose_num, 1)
    elif dose_base == 'AUC依存':
        if ccr is None:
            return None
        return round((ccr + 25) * dose_num, 0)
    elif dose_base == 'BW依存':
        if bw is None:
            return None
        return round(bw * dose_num, 1)
    return None


def calc_ccr(bw, age, scr, sex):
    """Cockcroft-Gault 式で Ccr を計算"""
    try:
        ccr = (140 - age) * bw / (72 * scr)
        if sex == '女':
            ccr *= 0.85
        return round(ccr, 1)
    except:
        return None


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
            '一般名（全角）'     : master.get('一般名（全角）', ''),
            '採用商品名（全角）' : master.get('採用商品名（全角）', ''),
            '薬剤区分'          : master.get('薬剤区分', ''),
            '支持療法分類'      : master.get('支持療法分類', ''),
            '1V当たりmg'        : master.get('1V当たりmg', ''),
            '患者向け説明'      : master.get('患者向け説明', ''),
        })
        drugs.append(merged)
    return {'basic': basic, 'drugs': drugs, 'master_dict': master_dict}


# ===== 副作用フラグ取得 =====

def get_ae_flags(protocol_no, drug_data, ae_data):
    cancer_codes = list(dict.fromkeys([
        str(d.get('管理コード', '')).strip()
        for d in drug_data
        if str(d.get('プロトコールNo', '')).strip() == protocol_no
        and str(d.get('管理コード', '')).strip().startswith('AC')
    ]))
    ae_dict  = {str(r.get('管理コード', '')).strip(): r for r in ae_data}
    ae_flags = {}
    for code in cancer_codes:
        ae_row = ae_dict.get(code, {})
        for col_name, val in ae_row.items():
            if str(val).strip() == '○':
                ae_flags[col_name] = True
    return ae_flags


# ===== O欄テキスト生成 =====

def to_half_kana(text):
    table = {
        'ア':'ｱ','イ':'ｲ','ウ':'ｳ','エ':'ｴ','オ':'ｵ',
        'カ':'ｶ','キ':'ｷ','ク':'ｸ','ケ':'ｹ','コ':'ｺ',
        'サ':'ｻ','シ':'ｼ','ス':'ｽ','セ':'ｾ','ソ':'ｿ',
        'タ':'ﾀ','チ':'ﾁ','ツ':'ﾂ','テ':'ﾃ','ト':'ﾄ',
        'ナ':'ﾅ','ニ':'ﾆ','ヌ':'ﾇ','ネ':'ﾈ','ノ':'ﾉ',
        'ハ':'ﾊ','ヒ':'ﾋ','フ':'ﾌ','ヘ':'ﾍ','ホ':'ﾎ',
        'マ':'ﾏ','ミ':'ﾐ','ム':'ﾑ','メ':'ﾒ','モ':'ﾓ',
        'ヤ':'ﾔ','ユ':'ﾕ','ヨ':'ﾖ',
        'ラ':'ﾗ','リ':'ﾘ','ル':'ﾙ','レ':'ﾚ','ロ':'ﾛ',
        'ワ':'ﾜ','ヲ':'ｦ','ン':'ﾝ',
        'ァ':'ｧ','ィ':'ｨ','ゥ':'ｩ','ェ':'ｪ','ォ':'ｫ',
        'ッ':'ｯ','ャ':'ｬ','ュ':'ｭ','ョ':'ｮ',
        'ガ':'ｶﾞ','ギ':'ｷﾞ','グ':'ｸﾞ','ゲ':'ｹﾞ','ゴ':'ｺﾞ',
        'ザ':'ｻﾞ','ジ':'ｼﾞ','ズ':'ｽﾞ','ゼ':'ｾﾞ','ゾ':'ｿﾞ',
        'ダ':'ﾀﾞ','ヂ':'ﾁﾞ','ヅ':'ﾂﾞ','デ':'ﾃﾞ','ド':'ﾄﾞ',
        'バ':'ﾊﾞ','ビ':'ﾋﾞ','ブ':'ﾌﾞ','ベ':'ﾍﾞ','ボ':'ﾎﾞ',
        'パ':'ﾊﾟ','ピ':'ﾋﾟ','プ':'ﾌﾟ','ペ':'ﾍﾟ','ポ':'ﾎﾟ',
        'ー':'ｰ','ヴ':'ｳﾞ','・':'･',
    }
    result = ''
    for char in str(text):
        result += table.get(char, char)
    return result

INJECTION_ORDER = {
    'NK1':1,'5HT3':2,'ステロイド':3,'G-CSF':4,
    '利尿薬':5,'解毒薬':6,'抗アレルギー':7,
    'H2ブロッカー':8,'電解質補正':9,'その他注射':10
}

def build_o_text(protocol_no, basic_data, drug_data, master_data,
                 bsa, bw, ccr, start_date, course_num):
    """O欄テキストを文字列で生成して返す。"""
    result = get_regimen(protocol_no, basic_data, drug_data, master_data)
    if result is None:
        return ''
    basic       = result['basic']
    drugs       = result['drugs']
    master_dict = result['master_dict']

    cancer_drugs     = [d for d in drugs if str(d.get('①O欄_抗がん剤', '')) == '○']
    support_inj_all  = sorted(
        [d for d in drugs
         if str(d.get('①O欄_支持療法', '')) == '○'
         and str(d.get('投与順序', '')) != '内服'],
        key=lambda d: INJECTION_ORDER.get(str(d.get('支持療法分類', '')), 99)
    )
    support_oral_all = [d for d in drugs
                        if str(d.get('①O欄_支持療法', '')) == '○'
                        and str(d.get('投与順序', '')) == '内服']

    lines = []

    # 1行目：レジメン名
    lines.append(
        f"O；【{protocol_no}】{basic['レジメン名']}"
        f"(1ｸｰﾙ{basic['1コース日数']}日)"
    )

    # 2行目：患者情報
    patient_parts = []
    if bsa is not None:
        patient_parts.append(f"BSA：{bsa:.3f} m2")
    if bw is not None:
        patient_parts.append(f"BW：{bw:.1f} kg")
    if ccr is not None:
        patient_parts.append(f"Ccr：{ccr:.1f} mL/min")
    if patient_parts:
        lines.append("  " + "  ".join(patient_parts))

    # 開始日・コース目
    if start_date:
        course_marks = ["①","②","③","④","⑤","⑥","⑦","⑧","⑨","⑩"]
        mark = course_marks[course_num - 1] if 1 <= course_num <= 10 else f"{course_num}"
        lines.append(f"  {mark}コース目　{start_date.strftime('%Y/%m/%d')}～")

    lines.append("")

    # 抗がん剤行
    for drug in cancer_drugs:
        code      = str(drug['管理コード'])
        master    = master_dict.get(code, {})
        name_half = to_half_kana(str(master.get('一般名（全角）', '')))
        dose_base = str(drug.get('用量根拠', ''))
        dose_str, unit_str = format_dose_text(drug)
        try:
            _rv = str(drug.get('投与量数値', '') or '').strip()
            _rv = ''.join(c for c in _rv if c.isdigit() or c == '.')
            dose_num = float(_rv or 0)
        except:
            dose_num = 0
        day_str = str(drug.get('投与Day文字', ''))

        # 計算用量
        calc     = calc_dose(drug, bsa, bw, ccr)
        calc_num = calc_dose_num(drug, bsa, bw, ccr)

        # 実投与量（セッションステートから取得）
        _actual_key = f"p6_actual_{protocol_no}_{code}"
        _actual_val = st.session_state.get(_actual_key, 0.0)
        _actual_mg  = float(_actual_val) if _actual_val else 0.0

        # 投与量表示文字列
        if dose_base == '固定用量':
            _dose_display = f"投与量：{dose_str}{unit_str}"
        elif _actual_mg > 0:
            _actual_str = str(int(_actual_mg)) if _actual_mg == int(_actual_mg) else str(_actual_mg)
            if calc_num and calc_num > 0:
                _rate = _actual_mg / calc_num * 100
                _dose_display = f"計算値：{calc}　処方量：{_actual_str}mg（{_rate:.1f}%）"
            else:
                _dose_display = f"処方量：{_actual_str}mg"
        else:
            _dose_display = f"投与量：{calc}"

        if dose_base == '固定用量':
            lines.append(
                f"{name_half} ({dose_str}{unit_str})"
                f"  {_dose_display}  {day_str}"
            )
        elif dose_base == 'AUC依存':
            lines.append(
                f"{name_half} (AUC{int(dose_num)})"
                f"  {_dose_display}  {day_str}"
            )
        else:
            _dn = int(dose_num) if dose_num == int(dose_num) else dose_num
            lines.append(
                f"{name_half} ({_dn}{unit_str})"
                f"  {_dose_display}  {day_str}"
            )

    # 支持療法行
    if support_inj_all:
        inj_parts = []
        for d in support_inj_all:
            name = to_half_kana(
                str(d.get('商品名', '') or d.get('採用商品名（全角）', ''))
            )
            ds, us = format_dose_text(d)
            day = str(d.get('投与Day文字', ''))
            inj_parts.append(f"{name} {ds}{us}({day})")
        lines.append("支持療法：" + "､".join(inj_parts))

    if support_oral_all:
        oral_parts = []
        for d in support_oral_all:
            name = to_half_kana(
                str(d.get('商品名', '') or d.get('採用商品名（全角）', ''))
            )
            ds, us = format_dose_text(d)
            day = str(d.get('投与Day文字', ''))
            oral_parts.append(f"{name} {ds}{us}({day})")
        lines.append("　　　　　" + "､".join(oral_parts))

    return "\n".join(lines)


# ===== O欄＋Pd欄テキスト生成 =====

AE_ITEMS = [
    ("●嘔吐　　　　　　□なし　□あり　(嘔吐回数;　　　)",                       True,  None),
    ("●悪心　　　　　　□なし　□G1　□G2　□G3",                                True,  None),
    ("●食欲不振　　　　□なし　□G1　□G2　□G3　□G4",                          True,  None),
    ("●便秘　　　　　　□なし　□あり　（ベースライン 回数：　　　、BS：　　　）", True,  None),
    ("●倦怠感　　　　　□なし　□G1　□G2　□G3",                                True,  None),
    (
        "●骨髄抑制\n"
        "　　WBC　　□なし　□G1　□G2　□G3　□G4\n"
        "　　Neut　　□なし　□G1　□G2　□G3　□G4\n"
        "　　Hb　　　□なし　□G1　□G2　□G3　□G4\n"
        "　　PLT　　　□なし　□G1　□G2　□G3　□G4",
        True, None
    ),
    ("●肝機能障害　　　□なし　□あり",                                          True,  None),
    ("●腎機能障害　　　□なし　□あり",                                          True,  None),
    ("●電解質異常　　　□なし　□あり",                                          True,  None),
    ("●下痢　　　　　　□なし　□あり　（ベースライン 回数：　　　、BS：　　　）", False, "下痢"),
    ("●口腔粘膜炎　　　□なし　□G1　□G2　□G3　□G4",                         False, "口腔粘膜炎"),
    ("●脱毛　　　　　　□なし　□G1　□G2",                                     False, "脱毛"),
    ("●末梢神経障害　　□なし　□G1　□G2　□G3　□G4",                         False, "末梢神経障害"),
    ("●味覚異常　　　　□なし　□G1　□G2",                                     False, "味覚異常"),
    ("●IRR　　　　　　 □なし　□あり",                                         False, "IRR"),
    ("●手足症候群　　　□なし　□G1　□G2　□G3",                               False, "手足症候群"),
    ("●皮膚障害　　　　□なし　□あり",                                         False, "皮膚障害"),
    ("●間質性肺炎　　　□なし　□あり",                                         False, "間質性肺炎"),
    ("●心障害　　　　　□なし　□あり",                                         False, "心障害"),
    ("●その他（　　　　　　　　　　　）",                                       True,  None),
]

IRAE_ITEMS = [
    "　●IRR　　　　　　　　　　　　□なし　□あり",
    "　●IP様症状(空咳･呼吸困難感)　□なし　□あり；/ KL-6：",
    "　●肝機能異常　　　　　　　　　□なし　□あり",
    "　●腎機能異常　　　　　　　　　□なし　□あり",
    "　●甲状腺機能異常　　　　　　　□なし　□あり；/ FT3：　FT4：　TSH：",
    "　●下垂体機能異常　　　　　　　□なし　□あり；/ ACTH：",
    "　●副腎機能障害　　　　　　　　□なし　□あり；/ コルチゾール：",
    "　●高血糖　　　　　　　　　　　□なし　□あり；/ HbA1c：",
    "　●静脈血栓症　　　　　　　　　□なし　□あり；/ D-dimer：",
    "　●皮膚障害　　　　　　　　　　□なし　□あり",
    "　●下痢・大腸炎　　　　　　　　□なし　□あり",
    "　●眼症状(ぶどう膜炎)　　　　　□なし　□あり",
    "　●筋炎･横紋筋融解症/重症筋無力症/心筋炎　□なし　□あり；/ CK：",
    "　●神経障害(ギラン・バレー症候群）　　　　 □なし　□あり",
    "　●脳炎/髄膜炎　　　　　　　　□なし　□あり",
    "　●膵炎　　　　　　　　　　　　□なし　□あり；/ リパーゼ：　アミラーゼ：",
    "　●その他（　　　　　　　　　　　）",
]

PD_TEXTS = {
    "骨髄抑制": (
        "【骨髄抑制】化学療法開始から1～2週間後に白血球/好中球が減少します。"
        "感染症にかかりやすい状態になりますので、手洗い・うがい・マスク着用などの"
        "感染予防を徹底しましょう。発熱（38℃以上）があればすぐにご連絡ください。\n"
        "　血小板が減ると出血しやすくなります。鼻血・歯ぐきからの出血・皮下出血に注意してください。"
    ),
    "悪心嘔吐": (
        "【悪心嘔吐】症状がある場合は、脂っぽい食事は避けるとよいでしょう。\n"
        "　脱水症状にならないように、スポーツドリンクなどで水分をとりましょう。\n"
        "　吐き気を抑える薬が処方されている場合は、我慢せずに飲みましょう。"
    ),
    "末梢神経障害": (
        "【末梢神経障害】抗がん剤により手足のしびれや感覚異常が出ることがあり、"
        "1ヶ月〜数ヶ月経ってから現れることもあります。"
        "症状の出方には個人差がありますが、日常生活に支障が出る前に"
        "状態を医療スタッフへお伝えください。"
    ),
    "口腔粘膜炎": (
        "【口腔粘膜炎】炎症が起こりやすくなります。"
        "毎朝お口の中（歯ぐき・頬・舌など）をチェックし、"
        "食前のうがいと食後の優しい歯磨きで口腔内を清潔に保ちましょう。"
    ),
    "脱毛": (
        "【脱毛】治療開始から2～3週間後に脱毛が始まることがあります。"
        "治療終了後に毛髪は回復することがほとんどです。"
        "ウィッグや帽子などを活用しましょう。"
    ),
    "下痢": (
        "【下痢】水分補給を十分に行いましょう。"
        "脂っぽいもの・刺激物・乳製品は控えめにしてください。"
        "下痢が続く場合や血便がある場合はすぐにご連絡ください。"
    ),
    "手足症候群": (
        "【手足症候群】手のひらや足の裏が赤くなったり、ひび割れ・水疱・痛みが"
        "出ることがあります。保湿クリームでのケアが重要です。"
        "靴擦れや締め付けを避け、皮膚への刺激を減らしましょう。"
    ),
    "irAE": (
        "【irAE（免疫関連有害事象）】免疫チェックポイント阻害薬では、"
        "免疫が過剰に働くことで様々な臓器に炎症が起きることがあります。"
        "息苦しさ・発熱・皮疹・下痢・倦怠感など普段と異なる症状が出た場合は"
        "すぐにご連絡ください。"
    ),
    "IRR": (
        "【IRR（注入関連反応）】点滴中や点滴後に発熱・寒気・蕁麻疹・"
        "呼吸困難感などの症状が出ることがあります。"
        "症状を感じたらすぐに看護師にお知らせください。"
    ),
}


def build_opd_text(protocol_no, basic_data, drug_data,
                   master_data, ae_data, pd_data,
                   bsa, bw, ccr, start_date, course_num):
    """O欄＋Pd欄 全テキストを生成して返す。"""
    ae_flags = get_ae_flags(protocol_no, drug_data, ae_data)
    has_irae = ae_flags.get("irAE", False)

    basic_row   = next(
        (b for b in basic_data if b['プロトコールNo'] == protocol_no), {}
    )
    pd_cat_raw  = str(basic_row.get('Pdカテゴリ', '')).strip()
    pd_cat_list = [x.strip() for x in pd_cat_raw.split('|') if x.strip()]

    pda_list = sorted(
        [p for p in (pd_data or [])
         if str(p.get('種別', '')).strip() == 'A'],
        key=lambda x: int(x['優先順位'])
        if str(x.get('優先順位', '')).isdigit() else 99
    )
    matched_pda = [p for p in pda_list if p['カテゴリID'] in pd_cat_list]

    pda007 = next((p for p in (pd_data or []) if p['カテゴリID'] == 'PDA007'), None)
    if pda007 and ae_flags.get('脱毛', False):
        if not any(p['カテゴリID'] == 'PDA007' for p in matched_pda):
            matched_pda.append(pda007)
            matched_pda = sorted(
                matched_pda,
                key=lambda x: int(x['優先順位'])
                if str(x.get('優先順位', '')).isdigit() else 99
            )

    lines = []

    # ブロック1：O欄本文
    o_text = build_o_text(
        protocol_no, basic_data, drug_data, master_data,
        bsa, bw, ccr, start_date, course_num
    )
    lines.append(o_text)
    lines.append("")

    # ブロック2：HBV
    lines.append("●B型肝炎スクリーニング")
    lines.append("　HBs-AG：□－　□＋　、HBs-AB：□－　□＋　、HBc-AB：□－　□＋")
    lines.append("　※＋がある場合→HBV-DNA　□検出限界以下（　　　　）")
    lines.append("")

    # ブロック3：副作用評価
    lines.append("◇副作用")
    lines.append(
        "*評価方法　評価日がd1：前クール全体で最も悪い検査値や症状\n"
        "　　　　　 上記以外　：d1～評価日の最も悪い検査値や症状"
    )
    for text, always, flag in AE_ITEMS:
        if not always and not ae_flags.get(flag, False):
            continue
        lines.append(text)
    lines.append("")

    # ブロック4：irAE
    if has_irae:
        lines.append("◇irAE　※検査値はベースラインを記載、最新値は【薬学的管理】参照")
        for item in IRAE_ITEMS:
            lines.append(item)
        lines.append("")

    # ブロック5：Pd記載
    lines.append(
        "Pd；ご本人に対して初回面談実施。服薬状況、服薬理解度および"
        "有害事象の発現状況の確認を行った。"
    )
    lines.append(
        "化学療法のしおり、メーカー作成パンフレット（パンフレット名記載）、"
        "添付する説明書を用いて、化学療法について説明"
        "（治療スケジュール、支持療法、副作用/対策）を行った。"
    )

    ae_label_parts = ["骨髄抑制", "口内炎", "脱毛", "悪心嘔吐等"]
    if has_irae:
        ae_label_parts.append("irAE")
    lines.append(
        "・代表的な有害事象（" + "・".join(ae_label_parts) + "）の対処法は以下の通り指導した。"
    )

    lines.append(PD_TEXTS["骨髄抑制"])
    lines.append(PD_TEXTS["悪心嘔吐"])

    CONDITIONAL_PD = [
        ("末梢神経障害", "末梢神経障害"),
        ("口腔粘膜炎",  "口腔粘膜炎"),
        ("脱毛",        "脱毛"),
        ("下痢",        "下痢"),
        ("手足症候群",  "手足症候群"),
        ("irAE",        "irAE"),
        ("IRR",         "IRR"),
    ]
    for flag_key, pd_key in CONDITIONAL_PD:
        if ae_flags.get(flag_key, False) and pd_key in PD_TEXTS:
            lines.append(PD_TEXTS[pd_key])

    for pda in matched_pda:
        cat_name = str(pda.get('カテゴリ名', '')).strip()
        text     = str(pda.get('説明文', '')).strip()
        if text:
            text_clean = text.replace("\r\n", "\n").replace("\r", "\n")
            lines.append(f"【{cat_name}】\n{text_clean}")
        else:
            lines.append(f"【{cat_name}】")

    lines.append("※個別の薬剤に関する説明事項等も追記（必須）")

    return "\n".join(lines)


# ===== Streamlit UI =====

st.title("📋 O欄・Pd欄 テキスト生成")
st.caption("レジメンを選択してパラメーターを入力するとカルテ貼り付け用テキストを生成します")
st.divider()

if st.button("🔄 データを最新化する", key="btn_refresh_6"):
    st.cache_data.clear()
    st.cache_resource.clear()
    st.rerun()

with st.spinner("データを読み込み中..."):
    basic_data, drug_data, master_data, pd_data, ae_data = load_all_data()

if not basic_data:
    st.error("データの読み込みに失敗しました。")
    st.stop()

# ===== レジメン選択 =====
st.subheader("① レジメン選択")

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
    key="selectbox_6"
)

if selected is None:
    st.stop()

protocol_no = selected.split('　')[0].strip()

# レジメン切り替え時のセッションリセット
if st.session_state.get("last_protocol_no_6") != protocol_no:
    for key in list(st.session_state.keys()):
        if key.startswith("p6_"):
            del st.session_state[key]
    st.session_state["last_protocol_no_6"] = protocol_no

selected_basic = next(
    (b for b in basic_data if b['プロトコールNo'] == protocol_no), None
)

if selected_basic:
    col1, col2, col3 = st.columns(3)
    col1.metric("プロトコールNo", protocol_no)
    col2.metric("1コース日数",    f"{selected_basic.get('1コース日数', '?')}日")
    col3.metric("対象疾患",       selected_basic.get('対象疾患', '?'))

    pd_cats = str(selected_basic.get('Pdカテゴリ', '')).strip()
    if pd_cats:
        st.success(f"✅ Pdカテゴリ：{pd_cats}")
    else:
        st.warning("⚠️ Pdカテゴリ未設定です。3_テンプレート生成ページから設定してください。")

# 必要なパラメーターを判定
result = get_regimen(protocol_no, basic_data, drug_data, master_data)
cancer_drugs = []
if result:
    cancer_drugs = [d for d in result['drugs']
                    if str(d.get('①O欄_抗がん剤', '')) == '○']

need_bsa = any(str(d.get('用量根拠', '')) == 'BSA依存'             for d in cancer_drugs)
need_ccr = any(str(d.get('用量根拠', '')) == 'AUC依存'             for d in cancer_drugs)
need_bw  = any(str(d.get('用量根拠', '')) in ('BW依存', 'AUC依存') for d in cancer_drugs)
need_age = need_ccr
need_sex = need_ccr

st.divider()

# ===== パラメーター入力 =====
st.subheader("② パラメーター入力")

bsa = bw = scr = age = ccr = None
sex = None

if need_bsa or need_bw or need_ccr:
    cols = st.columns(3)
    col_idx = 0

    if need_bw:
        with cols[col_idx % 3]:
            bw_input = st.number_input(
                "体重 (kg)", min_value=0.0, max_value=200.0,
                value=0.0, step=0.1, format="%.1f", key="p6_bw"
            )
            bw = bw_input if bw_input > 0 else None
        col_idx += 1

    if need_bsa:
        with cols[col_idx % 3]:
            bsa_input = st.number_input(
                "BSA (m²)", min_value=0.0, max_value=3.0,
                value=0.0, step=0.001, format="%.3f", key="p6_bsa"
            )
            bsa = bsa_input if bsa_input > 0 else None
        col_idx += 1

    if need_ccr:
        with cols[col_idx % 3]:
            scr_input = st.number_input(
                "SCr", min_value=0.0, max_value=20.0,
                value=0.0, step=0.01, format="%.2f", key="p6_scr"
            )
            scr = scr_input if scr_input > 0 else None
        col_idx += 1

    if need_age:
        with cols[col_idx % 3]:
            age_input = st.number_input(
                "年齢", min_value=0, max_value=120,
                value=0, step=1, key="p6_age"
            )
            age = age_input if age_input > 0 else None
        col_idx += 1

    if need_sex:
        with cols[col_idx % 3]:
            sex = st.selectbox(
                "性別", options=["", "男", "女"], key="p6_sex"
            )
        col_idx += 1

    if need_ccr:
        if bw and age and scr and sex:
            ccr = calc_ccr(bw, age, scr, sex)
            st.info(f"📊 Ccr（Cockcroft-Gault）：**{ccr} mL/min**")
        else:
            st.caption("💡 体重・年齢・SCr・性別を入力すると Ccr を自動計算します")

# 開始日・コース目
cols2 = st.columns(2)
with cols2[0]:
    start_date = st.date_input(
        "開始日", value=date.today(), key="p6_start_date"
    )
with cols2[1]:
    course_num = st.number_input(
        "コース目", min_value=1, max_value=10,
        value=1, step=1, key="p6_course"
    )

# ===== 実投与量入力（BSA/AUC/BW依存のみ） =====
_need_actual = any(
    str(d.get('用量根拠', '')) in ('BSA依存', 'AUC依存', 'BW依存')
    for d in cancer_drugs
)

if _need_actual and result:
    st.divider()
    st.subheader("② - 2　実際の投与量を入力（任意）")
    st.caption("計算値と異なる用量を投与する場合に入力してください。達成率を自動計算します。")

    for _d in cancer_drugs:
        _dose_base = str(_d.get('用量根拠', ''))
        if _dose_base not in ('BSA依存', 'AUC依存', 'BW依存'):
            continue
        _code     = str(_d.get('管理コード', ''))
        _mast     = result['master_dict'].get(_code, {})
        _name     = str(_mast.get('一般名（全角）', '') or _d.get('商品名', '') or _code)
        _calc     = calc_dose(_d, bsa, bw, ccr)
        _calc_num = calc_dose_num(_d, bsa, bw, ccr)
        _key      = f"p6_actual_{protocol_no}_{_code}"

        st.markdown(f"**{_name}**　計算値：{_calc}")
        _actual_input = st.number_input(
            f"{_name} 実際の投与量 (mg)",
            min_value=0.0,
            value=0.0,
            step=1.0,
            format="%.1f",
            key=_key,
        )
        if _actual_input > 0 and _calc_num and _calc_num > 0:
            _rate = _actual_input / _calc_num * 100
            st.caption(f"　達成率：{_rate:.1f}%")

st.divider()

# ===== テキスト生成 =====
st.subheader("③ テキスト生成")

if st.button(
    "📋 O欄・Pd欄テキストを生成する",
    type="primary",
    use_container_width=True,
    key="btn_generate_6"
):
    with st.spinner("テキスト生成中..."):
        full_text = build_opd_text(
            protocol_no, basic_data, drug_data,
            master_data, ae_data, pd_data,
            bsa, bw, ccr, start_date, course_num
        )
    st.session_state["p6_generated_text"] = full_text

# 生成済みテキストの表示
if "p6_generated_text" in st.session_state:
    text = st.session_state["p6_generated_text"]

    st.success("✅ 生成完了！下のテキストをコピーしてカルテに貼り付けてください。")

    st.text_area(
        label="📄 O欄・Pd欄テキスト（全選択してコピー）",
        value=text,
        height=600,
        key="p6_textarea"
    )

    import streamlit.components.v1 as components
    copy_js = f"""
    <button
        onclick="
            navigator.clipboard.writeText(
                document.getElementById('copy_target').innerText
            ).then(() => {{
                this.innerText = '✅ コピーしました！';
                setTimeout(() => {{ this.innerText = '📋 全文をコピーする'; }}, 2000);
            }});
        "
        style="
            background-color: #4CAF50;
            color: white;
            border: none;
            padding: 12px 24px;
            font-size: 16px;
            border-radius: 6px;
            cursor: pointer;
            width: 100%;
            margin-bottom: 8px;
        "
    >📋 全文をコピーする</button>
    <pre id="copy_target" style="display:none;">{text}</pre>
    """
    components.html(copy_js, height=60)

    st.caption("⚠️ 生成されたテキストは必ず内容を確認してから使用してください")

st.divider()
st.caption("💡 Pdカテゴリの設定・変更は **3_テンプレート生成** ページから行ってください")
