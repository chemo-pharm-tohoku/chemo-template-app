import streamlit as st
from google import genai
from google.genai import types
import gspread
from google.oauth2.service_account import Credentials
from collections import defaultdict
from pptx import Presentation
from pptx.util import Pt, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree
import io, json, re
from datetime import date, datetime
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from chemo_utils import (
    to_half_kana, shorten_regimen_name, parse_days_num,
    format_dose_text, make_support_line, get_regimen,
    create_pptx, INJECTION_ORDER
)

today_str = date.today().strftime("%Y%m%d")

st.title("📋 レジメン情報抽出・登録")
st.caption("PDFをアップロードしてAIが自動解析→スプレッドシートに登録→パワポ生成まで一気通貫")
st.divider()

# ===== 認証・初期化 =====
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

@st.cache_resource
def get_spreadsheet():
    gc = get_gspread_client()
    return gc.open_by_url(st.secrets["spreadsheet"]["url"])

@st.cache_data(ttl=600)
def load_definition():
    gc = get_gspread_client()
    SPREADSHEET_ID = "1dLEUYSZlrIK1uHqEtEAfS1jSAPpXCIiAiAk_iaRuY-8"
    sh = gc.open_by_key(SPREADSHEET_ID)
    ws = sh.worksheet("抽出定義書")
    return ws.acell("A1").value

@st.cache_data(ttl=300)
def load_master_data():
    sh = get_spreadsheet()
    master_data = sh.worksheet("薬品マスタ").get_all_records()
    notes_data  = sh.worksheet("注意事項").get_all_records()
    return master_data, notes_data

# ===== ヘルパー関数 =====
def get_val(d, *keys, default=""):
    for key in keys:
        if key in d and d[key] is not None:
            return d[key]
    return default

def get_drugs(data):
    return data.get("drug_info") or data.get("drugs") or []

def get_basic(data):
    return data.get("basic_info") or {}
