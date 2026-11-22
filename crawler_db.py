#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
[대산포크 B2B 인텔리전스] 4대 축산 플랫폼(미트박스, 금천미트, 미트프렌즈, 탑미트) 한돈 전수 크롤링,
시계열 누적 관리(SQLite `handon_master.db` Append Only & `한돈_누적관리대장.xlsx` / 구글시트 연동),
생산일자 대비 최저·최고가 분석, 중점관리 체크박스 누적 트래킹 엔진 (슬랙 발송 100% 제외)
"""

import json
import logging
import os
import random
import re
import sqlite3
import sys
import time
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict, dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple
from urllib.parse import parse_qs, quote, urlparse

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

import openpyxl
import pandas as pd
import requests
import urllib3
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger("HandonMasterDB")

WORK_DIR = Path(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = WORK_DIR / "handon_master.db"
EXCEL_MASTER_PATH = WORK_DIR / "한돈_누적관리대장.xlsx"

# ---------------------------------------------------------------------------
# 1. 수집 타겟 URL (4대 플랫폼 한돈)
# ---------------------------------------------------------------------------
MEATBOX_TARGET_URL = (
    "https://www.meatbox.co.kr/fo/product/productListPage.do?displayCategorySeq=10010"
)
MEATBOX_API_URL = "https://mb-api.meatbox.co.kr/product/api/v1/search/filter"

EKCM_TARGET_URL = (
    "https://www.ekcm.co.kr/dp/subMain?dispCtgNo=31&uprTopDispCtgNo=3"
    "&dispCtgNm=%EA%B5%AD%EB%82%B4%EC%82%B0+%EB%8F%88%EC%9C%A1&leafCtgNo&dispCtgNoList"
)
EKCM_CATEGORY_API_URL = (
    "https://gw.ekcm.co.kr/api/display/v1/displayCategory/getCornContListInfo"
    "?vueCmptId=Disp_leaf_category&dispCtgNo=31&saleStat=10&sortCtgYn=N"
)
EKCM_GOODS_API_URL = "https://gw.ekcm.co.kr/api/goods/v1/goods/dispGoodsList"
EKCM_DETAIL_PAGE_BASE = "https://www.ekcm.co.kr/pd/productDetail"
EKCM_DETAIL_API_BASE = "https://gw.ekcm.co.kr/api/goods/v1/goods/goodsDetail/01"

MEATFRIENDS_TARGET_URL = "https://www.meatfriends.co.kr/display/displayIndex.do?catNo=7"
MEATFRIENDS_LIST_AJAX_URL = "https://www.meatfriends.co.kr/display/selectDisplayListPage.do"
MEATFRIENDS_DETAIL_URL = "https://www.meatfriends.co.kr/display/selectDisplayDetail.do"

TOPMEAT_TARGET_URL = "https://www.topmeat.co.kr/shop/list.php?ca_id=40"

# 흑돼지 및 수입육 제외 키워드
BLACK_PORK_KEYWORDS = [
    "흑돼지",
    "흑돈",
    "버크셔",
    "우리흑돈",
    "지리산흑돈",
    "제주흑돈",
    "(흑)",
    "[흑]",
    "흑통시",
    "흑미박",
    "거멍",
    "탐라애돈",
]

IMPORTED_MEAT_KEYWORDS = [
    "수입",
    "미국산",
    "스페인산",
    "칠레산",
    "캐나다산",
    "독일산",
    "네덜란드산",
    "오스트리아산",
    "멕시코산",
    "브라질산",
    "덴마크산",
    "프랑스산",
    "벨기에산",
    "아일랜드산",
    "이베리코",
    "듀록(수입)",
]

STANDARD_PART_ORDER = [
    "삼겹살",
    "미박삼겹살",
    "목살",
    "앞다리살(전지)",
    "미박앞다리살(미박전지)",
    "뒷다리살(후지)",
    "미박뒷다리살(미박후지)",
    "갈비",
    "등갈비",
    "안심",
    "등심",
    "항정살",
    "가브리살",
    "갈매기살",
    "사태",
    "미박사태",
    "사골",
    "등뼈",
    "장족",
    "장족(뒤)",
    "단족/미니족",
    "오돌뼈",
    "지방돈피/돼지껍데기",
    "돈지방(A지방)",
    "잡육",
    "꼬리",
    "머리고기/내장",
]
STANDARD_PARTS_ORDER = STANDARD_PART_ORDER

# 4대 플랫폼별 고유 카테고리/품목코드 1차 매핑 사전 (전수 코드 매핑)
TOPMEAT_CA_ID_MAP: Dict[str, Tuple[str, str]] = {
    "4010": ("삼겹살", "박피"),
    "4011": ("삼겹살", "박피"),
    "4020": ("삼겹살", "박피"),
    "40q0": ("삼겹살", "박피"),
    "40q1": ("삼겹살", "박피"),
    "40q2": ("삼겹살", "박피"),
    "40r0": ("삼겹살", "박피"),
    "40r2": ("삼겹살", "박피"),
    "4030": ("미박삼겹살", "미박"),
    "4031": ("미박삼겹살", "미박"),
    "4040": ("미박삼겹살", "미박"),
    "40s0": ("미박삼겹살", "미박"),
    "40t0": ("미박삼겹살", "미박"),
    "4000": ("목살", "일반"),
    "4050": ("목살", "일반"),
    "4053": ("목살", "일반"),
    "4059": ("목살", "일반"),
    "40q3": ("목살", "일반"),
    "40u0": ("목살", "일반"),
    "40u1": ("목살", "일반"),
    "4070": ("앞다리살(전지)", "박피"),
    "4080": ("앞다리살(전지)", "박피"),
    "4001": ("앞다리살(전지)", "박피"),
    "4002": ("앞다리살(전지)", "박피"),
    "4004": ("앞다리살(전지)", "박피"),
    "40r1": ("앞다리살(전지)", "박피"),
    "40a0": ("미박앞다리살(미박전지)", "미박"),
    "40b0": ("미박앞다리살(미박전지)", "미박"),
    "4003": ("미박앞다리살(미박전지)", "미박"),
    "40b1": ("미박앞다리살(미박전지)", "미박"),
    "40c0": ("뒷다리살(후지)", "박피"),
    "40o0": ("뒷다리살(후지)", "박피"),
    "40p0": ("뒷다리살(후지)", "박피"),
    "40d0": ("미박뒷다리살(미박후지)", "미박"),
    "40o1": ("미박뒷다리살(미박후지)", "미박"),
    "40l0": ("갈비", "일반"),
    "40y0": ("갈비", "일반"),
    "4061": ("등갈비", "일반"),
    "40z0": ("등갈비", "일반"),
    "40g0": ("안심", "일반"),
    "40g1": ("안심", "일반"),
    "40w1": ("안심", "일반"),
    "40g2": ("안심", "일반"),
    "40f0": ("등심", "일반"),
    "40w0": ("등심", "일반"),
    "40i0": ("항정살", "일반"),
    "40v2": ("항정살", "일반"),
    "40h0": ("가브리살", "일반"),
    "40v1": ("가브리살", "일반"),
    "40j0": ("갈매기살", "일반"),
    "40v3": ("갈매기살", "일반"),
    "40l1": ("미박사태", "미박"),
    "40l2": ("미박사태", "미박"),
    "40l4": ("미박사태", "미박"),
    "40i6": ("미박사태", "미박"),
    "40l5": ("미박사태", "미박"),
    "40l6": ("미박사태", "미박"),
    "40z4": ("사골", "일반"),
    "40v0": ("등뼈", "일반"),
    "40x0": ("등뼈", "일반"),
    "40z8": ("장족", "미박"),
    "40x1": ("장족", "미박"),
    "40z5": ("장족(뒤)", "미박"),
    "40x2": ("장족(뒤)", "미박"),
    "40z7": ("단족/미니족", "미박"),
    "40y7": ("오돌뼈", "일반"),
    "40z3": ("지방돈피/돼지껍데기", "미박"),
    "40z2": ("돈지방(A지방)", "일반"),
    "40z1": ("잡육", "일반"),
    "40z9": ("꼬리", "일반"),
}

TOPMEAT_FEMALE_CA_IDS: Set[str] = {
    "4020", "4040", "4050", "4080", "40b0",
    "40r0", "40t0", "40u1", "4002", "40b1",
}

# 금천미트 리프 카테고리 번호(dispCtgNo) 전수 매핑 사전
EKCM_LEAF_CTG_MAP: Dict[str, Tuple[str, str]] = {
    "310101": ("삼겹살", "박피"),
    "310102": ("삼겹살", "박피"),
    "310103": ("삼겹살", "박피"),
    "310104": ("삼겹살", "박피"),
    "310105": ("삼겹살", "박피"),
    "310123": ("삼겹살", "박피"),
    "310125": ("삼겹살", "박피"),
    "310126": ("삼겹살", "박피"),
    "813754": ("삼겹살", "박피"),
    "816494": ("삼겹살", "박피"),
    "310106": ("미박삼겹살", "미박"),
    "310107": ("미박삼겹살", "미박"),
    "810974": ("미박삼겹살", "미박"),
    "813744": ("미박삼겹살", "미박"),
    "813764": ("미박삼겹살", "미박"),
    "815474": ("미박삼겹살", "미박"),
    "310108": ("목살", "일반"),
    "310109": ("목살", "일반"),
    "310124": ("목살", "일반"),
    "310127": ("목살", "일반"),
    "818324": ("목살", "미박"),
    "818424": ("목살", "일반"),
    "310110": ("앞다리살(전지)", "박피"),
    "310112": ("앞다리살(전지)", "박피"),
    "310128": ("앞다리살(전지)", "박피"),
    "310111": ("미박앞다리살(미박전지)", "미박"),
    "310113": ("미박앞다리살(미박전지)", "미박"),
    "815484": ("미박앞다리살(미박전지)", "미박"),
    "310206": ("뒷다리살(후지)", "박피"),
    "310129": ("뒷다리살(후지)", "박피"),
    "310114": ("미박뒷다리살(미박후지)", "미박"),
    "815471": ("미박뒷다리살(미박후지)", "미박"),
    "310115": ("등심", "일반"),
    "310131": ("등심", "일반"),
    "811205": ("등심", "일반"),
    "813734": ("등심", "일반"),
    "310116": ("안심", "일반"),
    "310709": ("안심", "일반"),
    "310132": ("안심", "일반"),
    "310117": ("갈비", "일반"),
    "310133": ("갈비", "일반"),
    "310118": ("등갈비", "일반"),
    "310134": ("등갈비", "일반"),
    "310119": ("항정살", "일반"),
    "310135": ("항정살", "일반"),
    "310120": ("가브리살", "일반"),
    "310137": ("가브리살", "일반"),
    "310121": ("갈매기살", "일반"),
    "310136": ("갈매기살", "일반"),
    "818165": ("사태", "박피"),
    "819204": ("사태", "박피"),
    "818155": ("사태", "박피"),
    "310408": ("미박사태", "미박"),
    "816394": ("미박사태", "미박"),
    "816395": ("미박사태", "미박"),
    "818175": ("미박사태", "미박"),
    "819234": ("미박사태", "미박"),
    "818704": ("미박사태", "미박"),
    "310147": ("미박사태", "미박"),
    "310148": ("미박사태", "미박"),
    "310138": ("잡육", "일반"),
    "310140": ("등뼈", "일반"),
    "310139": ("지방돈피/돼지껍데기", "미박"),
    "310145": ("장족", "미박"),
    "310146": ("장족(뒤)", "미박"),
    "816544": ("단족/미니족", "미박"),
    "816554": ("단족/미니족", "미박"),
    "815504": ("꼬리", "일반"),
    "818174": ("오돌뼈", "일반"),
    "311201": ("머리고기/내장", "일반"),
    "311202": ("머리고기/내장", "일반"),
    "811884": ("머리고기/내장", "일반"),
    "311204": ("머리고기/내장", "일반"),
    "311206": ("머리고기/내장", "일반"),
    "311207": ("머리고기/내장", "일반"),
}

EKCM_FEMALE_CTG_IDS: Set[str] = {
    "310102", "813764", "310105", "310107", "310109", "310112", "310113",
}

# 미트박스 표준 카테고리(itemCatName) 매핑 사전
MEATBOX_ITEM_CAT_MAP: Dict[str, Tuple[str, str]] = {
    "삼겹살": ("삼겹살", "박피"),
    "삼겹살(암퇘지)": ("삼겹살", "박피"),
    "삼겹살(미추리)": ("삼겹살", "박피"),
    "삼겹살(모돈)": ("삼겹살", "박피"),
    "두판 삼겹살": ("삼겹살", "박피"),
    "뼈삼겹살": ("삼겹살", "박피"),
    "미박삼겹살/오겹살": ("미박삼겹살", "미박"),
    "미박삼겹살(암퇘지)": ("미박삼겹살", "미박"),
    "미박삼겹살(수육용)": ("미박삼겹살", "미박"),
    "미박삼겹살(듀록)": ("미박삼겹살", "미박"),
    "미박 삼겹살(미추리)": ("미박삼겹살", "미박"),
    "목심": ("목살", "일반"),
    "목심(암퇘지)": ("목살", "일반"),
    "목심(정제목심)": ("목살", "일반"),
    "목심(듀록)": ("목살", "일반"),
    "전지/앞다리": ("앞다리살(전지)", "박피"),
    "전지/앞다리(암퇘지)": ("앞다리살(전지)", "박피"),
    "미박전지": ("미박앞다리살(미박전지)", "미박"),
    "미박전지(암퇘지)": ("미박앞다리살(미박전지)", "미박"),
    "후지": ("뒷다리살(후지)", "박피"),
    "후지(암퇘지)": ("뒷다리살(후지)", "박피"),
    "미박후지": ("미박뒷다리살(미박후지)", "미박"),
    "미박후지(암퇘지)": ("미박뒷다리살(미박후지)", "미박"),
    "갈비": ("갈비", "일반"),
    "LA갈비": ("갈비", "일반"),
    "등갈비": ("등갈비", "일반"),
    "안심": ("안심", "일반"),
    "등심": ("등심", "일반"),
    "뼈등심": ("등심", "일반"),
    "스킨등심": ("등심", "일반"),
    "항정살": ("항정살", "일반"),
    "항정살(암퇘지)": ("항정살", "일반"),
    "등심덧살(가브리살)": ("가브리살", "일반"),
    "갈매기살": ("갈매기살", "일반"),
    "갈매기살(막)": ("갈매기살", "일반"),
    "사태": ("사태", "박피"),
    "미박사태": ("미박사태", "미박"),
    "사골": ("사골", "일반"),
    "등뼈": ("등뼈", "일반"),
    "장족": ("장족", "미박"),
    "장족(앞)": ("장족", "미박"),
    "장족(뒤)": ("장족(뒤)", "미박"),
    "단족/미니족": ("단족/미니족", "미박"),
    "오돌뼈": ("오돌뼈", "일반"),
    "돈피/돼지껍데기": ("지방돈피/돼지껍데기", "미박"),
    "지방돈피/돼지껍데기": ("지방돈피/돼지껍데기", "미박"),
    "지방(A)": ("돈지방(A지방)", "일반"),
    "잡육": ("잡육", "일반"),
    "꼬리": ("꼬리", "일반"),
}

# 미트프렌즈 축종/품목 카테고리 매핑 사전
MEATFRIENDS_CAT_MAP: Dict[str, Tuple[str, str]] = {
    "삼겹살": ("삼겹살", "박피"),
    "미박삼겹살": ("미박삼겹살", "미박"),
    "목심": ("목살", "일반"),
    "앞다리": ("앞다리살(전지)", "박피"),
    "미박앞다리": ("미박앞다리살(미박전지)", "미박"),
    "뒷다리": ("뒷다리살(후지)", "박피"),
    "미박뒷다리": ("미박뒷다리살(미박후지)", "미박"),
    "갈비": ("갈비", "일반"),
    "등갈비": ("등갈비", "일반"),
    "안심": ("안심", "일반"),
    "등심": ("등심", "일반"),
    "항정살": ("항정살", "일반"),
    "등심덧살": ("가브리살", "일반"),
    "갈매기살": ("갈매기살", "일반"),
    "사태": ("사태", "박피"),
    "미사태": ("미박사태", "미박"),
    "미박사태": ("미박사태", "미박"),
    "사골": ("사골", "일반"),
    "등뼈": ("등뼈", "일반"),
    "장족": ("장족", "미박"),
    "단족": ("단족/미니족", "미박"),
    "미니족": ("단족/미니족", "미박"),
    "돈피": ("지방돈피/돼지껍데기", "미박"),
    "돈지방": ("돈지방(A지방)", "일반"),
    "잡육": ("잡육", "일반"),
    "돈꼬리": ("꼬리", "일반"),
    "꼬리": ("꼬리", "일반"),
    "돈두롤(롤삼겹)": ("머리고기/내장", "일반"),
    "돈토시살": ("머리고기/내장", "일반"),
}

EKCM_ARTC_EXACT_MAP: Dict[str, Tuple[str, str]] = {
    "1005030005": ("미박삼겹살", "미박"),
    "1005030187": ("미박삼겹살", "미박"),
    "1005030437": ("미박삼겹살", "미박"),
    "1005030453": ("미박삼겹살", "미박"),
    "1005030478": ("미박삼겹살", "미박"),
    "1005030197": ("등갈비", "일반"),
    "1005030454": ("등갈비", "일반"),
    "1005030462": ("등갈비", "일반"),
    "1005030475": ("등갈비", "일반"),
    "1005030479": ("등갈비", "일반"),
    "1005030032": ("등갈비", "일반"),
    "1005030458": ("등갈비", "일반"),
    "1005030018": ("갈매기살", "일반"),
    "1005030168": ("갈매기살", "일반"),
    "1005030198": ("갈매기살", "일반"),
    "1005030260": ("갈매기살", "일반"),
    "1005030425": ("갈매기살", "일반"),
    "1005050003": ("미박앞다리살(미박전지)", "미박"),
    "1005050123": ("미박앞다리살(미박전지)", "미박"),
    "1005050304": ("미박앞다리살(미박전지)", "미박"),
    "1005050310": ("미박앞다리살(미박전지)", "미박"),
    "1005050004": ("미박앞다리살(미박전지)", "미박"),
    "1005050319": ("미박앞다리살(미박전지)", "미박"),
    "1005050029": ("항정살", "일반"),
    "1005050129": ("항정살", "일반"),
    "1005050166": ("항정살", "일반"),
    "1005050293": ("항정살", "일반"),
    "1005050305": ("항정살", "일반"),
    "1005060002": ("미박뒷다리살(미박후지)", "미박"),
    "1005060038": ("미박뒷다리살(미박후지)", "미박"),
    "1005060135": ("미박뒷다리살(미박후지)", "미박"),
    "1005060146": ("미박뒷다리살(미박후지)", "미박"),
    "1005060137": ("미박뒷다리살(미박후지)", "미박"),
    "1005020004": ("가브리살", "일반"),
    "1005020140": ("가브리살", "일반"),
    "1005020151": ("가브리살", "일반"),
    "1005020157": ("가브리살", "일반"),
    "1005020161": ("가브리살", "일반"),
    "1005100563": ("사태", "박피"),
    "1005100600": ("미박사태", "미박"),
    "1005100026": ("미박사태", "미박"),
    "1005100062": ("미박사태", "미박"),
    "1005100559": ("미박사태", "미박"),
    "1005100029": ("미박사태", "미박"),
    "1005100065": ("미박사태", "미박"),
    "1005100546": ("미박사태", "미박"),
    "1005100547": ("미박사태", "미박"),
    "1005100021": ("등뼈", "일반"),
    "1005100577": ("등뼈", "일반"),
    "1005100024": ("장족", "미박"),
    "1005100060": ("장족", "미박"),
    "1005100027": ("장족(뒤)", "미박"),
    "1005100533": ("장족(뒤)", "미박"),
    "1005100604": ("장족(뒤)", "미박"),
    "1005100385": ("단족/미니족", "미박"),
    "1005100557": ("단족/미니족", "미박"),
    "1005100564": ("오돌뼈", "일반"),
    "1005100020": ("잡육", "일반"),
    "1005100538": ("꼬리", "일반"),
}

KEEPINGS_ORDER = ["냉장", "냉동"]
SEX_TYPES_ORDER = ["일반", "암퇘지"]

PLATFORM_5_GROUPS = [
    "미트박스 - 근본한돈",
    "미트박스 - 일반 한돈",
    "금천미트",
    "미트프렌즈",
    "탑미트",
]
PLATFORM_GROUPS_ORDER = PLATFORM_5_GROUPS

PROCESSED_CUT_PATTERN = re.compile(
    r"세절|슬라이스|대패|다짐|분쇄|깍두기|채썰기|탕수육채|탕수육용채|큐브|돌돌이|롤삼겹|꽃삼겹|돈두롤|"
    r"키로육|구이용|구이/찜용|국/탕/찜용|장조림용|찌개용|불고기용|돈까스용|돈가스용|카레|짜장|잡채용|제육용|"
    r"절단|컷팅|포기김치|양념|시즈닝|마리네이드|\d+(?:\.\d+)?\s*mm|"
    r"\(낱개\)|\[낱개|낱개상품|한팩|두팩|소분|세절&분쇄|\d+g\s*\*\s*\d+|\b\d+g\b"
)

# 원물 박스육의 지방 정선 두께 표기(예: '지방두께 3mm', '지방 10mm 작업', '뒷다리살 3mm', 'mm등심')는 세절육 오탐에서 제외
FAT_TRIM_SPEC_REGEX = re.compile(
    r"지방\s*두께\s*\d+(?:\.\d+)?\s*mm|지방\s*\d+(?:\.\d+)?\s*mm(?:\s*작업)?|"
    r"뒷다리살\s*[03]\s*mm|\bmm등심\b",
    re.I,
)

FEMALE_PORK_REGEX = re.compile(
    r"암퇘지|암돼지|\(암\)|\[암\]|암돈|미경산|(?:^|[\s/\-_])암\)"
)


@dataclass
class PorkSnapshotRecord:
    """SQLite `product_snapshots` 테이블 및 누적 엑셀 적재용 표준 데이터 레코드"""

    run_id: int
    collected_at: str         # YYYY-MM-DD HH:MM:SS
    collected_date: str       # YYYY-MM-DD
    item_key: str             # '{source_site}:{product_id}' (중점관리 체크박스 연결 고유키)
    platform_group: str       # 5개 비교군
    source_site: str          # 4대 플랫폼
    is_geunbon: int           # 1 or 0
    product_id: str
    raw_name: str
    raw_category: str
    brand: str                # 브랜드명 (근본한돈, 도드람, 선진포크, 하이포크, 목우촌 등)
    keeping: str              # '냉장' or '냉동'
    standard_part: str        # 13개 표준부위
    skin_type: str            # '박피', '미박', '일반'
    sub_spec: str             # '일반', '암퇘지', '수육/특수정선', '모돈'
    processing_type: str      # '원물/박스육' vs '세절/가공육'
    mapping_status: str       # 이중 검증 상태
    unit_price_kg: int        # 1kg당 최저 판매단가 (원/kg)
    max_price_kg: int         # 1kg당 최고 판매단가 (원/kg)
    avg_price_kg: int         # 1kg당 평균 판매단가 (원/kg)
    price_norm_method: str    # 단가 환산 방식
    price_display: str        # 원본 가격 표기
    min_proc_date: str        # 최저가 박스 생산(가공)일자 (YYYY-MM-DD)
    max_proc_date: str        # 최고가 박스 생산(가공)일자 (YYYY-MM-DD)
    proc_date_range: str      # 전체 생산일자 범위
    exp_date_range: str       # 유통/소비기한 범위
    prod_date_bucket: str     # 생산일자 구간 ('당월 신상품(최근 14일)', '전월/일반 생산(15~45일)', '장기보관/기한임박(45일 경과)', '주문당일생산/상시')
    is_sold_out: int          # 1(품절) or 0(판매중)
    status_text: str
    stock_qty: int            # 당일 잔여 재고 수량 (박스/팩)
    option_cnt: int           # 유효 옵션 등록 수
    product_url: str          # 최저가/대표 상품상세 하이퍼링크 URL
    max_product_url: str = "" # 최고가 박스/옵션 상품상세 하이퍼링크 URL (미지정 시 product_url 사용)


def create_http_session() -> requests.Session:
    session = requests.Session()
    retry = Retry(
        total=3,
        read=3,
        connect=3,
        backoff_factor=0.6,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset(["GET", "POST"]),
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=35, pool_maxsize=35)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    session.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
            ),
            "Accept": "application/json, text/html, application/xhtml+xml, */*",
            "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
        }
    )
    return session


build_retry_session = create_http_session


# ---------------------------------------------------------------------------
# 표준화 매핑 & 브랜드 정제 & 생산일자 구간 분류기
# ---------------------------------------------------------------------------
class PorkStandardizer:
    @staticmethod
    def is_excluded_black_or_imported(text: str) -> Tuple[bool, str]:
        norm = str(text or "").strip()
        for kw in BLACK_PORK_KEYWORDS:
            if kw in norm:
                return True, f"흑돼지제외({kw})"
        if re.search(r"흑돼지|흑돈|버크셔|우리흑돈|지리산흑돈|제주흑돈|흑미박|\(흑\)|\[흑\]", norm):
            return True, "흑돼지제외(정규식)"
        for ikw in IMPORTED_MEAT_KEYWORDS:
            if ikw in norm:
                return True, f"수입육제외({ikw})"
        return False, ""

    @staticmethod
    def clean_brand_name(raw_brand: str, raw_name: str, source_site: str) -> str:
        """상품명 및 브랜드 필드에서 핵심 브랜드명(근본한돈, 도드람, 선진포크, 하이포크, 목우촌 등)을 깔끔하게 정규화"""
        b = str(raw_brand or "").strip()
        n = str(raw_name or "").strip()
        combined = f"{b} {n}"

        if "근본한돈" in combined:
            return "근본한돈"

        # 대표 한돈 브랜드 사전 매칭
        major_brands = [
            "도드람",
            "선진포크",
            "하이포크",
            "목우촌",
            "농협안심한돈",
            "포크밸리",
            "보성녹돈",
            "당당한돈",
            "빽포크",
            "한돈명가",
            "미트체인지",
            "천금한돈",
            "안동참마돼지",
            "골든포크",
            "몬트락제주",
            "태흥한돈",
            "금천한돈",
            "산들내음",
            "우경축산",
            "박달재한돈",
            "제주도니",
            "탐라포크",
            "대산포크",
            "엘리트포크",
            "배산장군포크",
            "문경약돌돼지",
            "좋은생각",
            "고기랑",
            "돈마루",
            "미소찬한돈",
            "청정원",
            "바른한돈",
            "더조은한돈",
            "다드림한돈",
            "초이스한돈",
            "내포원",
        ]
        for mb in major_brands:
            if mb in combined:
                return mb

        if b and b not in ("브랜드없음", "기타", "한돈", "국내산", "국산"):
            # '피그마스터 - 1등급이상' -> '피그마스터'
            b_clean = b.split(" - ")[0].strip()
            return b_clean

        m_bracket = re.match(r"^\[([^\]]+)\]", n)
        if m_bracket:
            cand = m_bracket.group(1).strip()
            if cand not in ("국산", "국내산", "한돈", "냉장", "냉동"):
                return cand

        if "|" in n:
            tail = n.split("|")[-1].strip()
            tail_clean = tail.split(" - ")[0].strip()
            if tail_clean and len(tail_clean) <= 15:
                return tail_clean

        return f"{source_site}직영/한돈"

    @staticmethod
    def normalize_keeping(raw_keeping: str, raw_name: str, raw_cat: str) -> str:
        combined = f"{raw_keeping} {raw_cat} {raw_name}"
        if "냉동" in raw_keeping or "동결" in raw_keeping:
            return "냉동"
        if "냉장" in raw_keeping:
            if "냉동" in raw_cat or "[냉동]" in raw_name or "냉동" in raw_name:
                return "냉동"
            return "냉장"
        if "냉동" in combined or "동결" in combined:
            return "냉동"
        return "냉장"

    @staticmethod
    def detect_sex_type(
        raw_name: str,
        raw_cat: str,
        sub_spec: str = "",
        source_site: str = "",
        product_id: str = "",
        sex_hint: str = "",
    ) -> str:
        """
        암퇘지 상품 전용 판별기 ('암퇘지' vs '일반'):
        - 모돈/후보돈은 번식돈 원료육이므로 '암퇘지(미경산 암퇘지)' 비교군에서 제외('일반' 내 모돈으로 분리)
        - 플랫폼별 암퇘지 전용 카테고리 코드(탑미트 ca_id, 금천미트 dispCtgNo/lventSexGbNm) 및 상품명·등급 텍스트 교차 검증
        """
        combined = f"{raw_name} {raw_cat} {sex_hint}".strip()
        if re.search(r"모돈|후보돈", combined) or sub_spec == "모돈":
            return "일반"
        if sub_spec == "암퇘지" or sex_hint in ("암", "암퇘지"):
            return "암퇘지"
        pid_parts = str(product_id or "").split("_")
        pid_prefix = pid_parts[0].strip() if pid_parts else ""
        pid_suffix = pid_parts[1].strip() if len(pid_parts) > 1 else ""
        if source_site == "탑미트" and pid_prefix in TOPMEAT_FEMALE_CA_IDS:
            return "암퇘지"
        if source_site == "금천미트" and pid_suffix in EKCM_FEMALE_CTG_IDS:
            return "암퇘지"
        if FEMALE_PORK_REGEX.search(combined):
            return "암퇘지"
        return "일반"

    @staticmethod
    def get_row_sex_type(row: Dict[str, Any]) -> str:
        if row.get("sex_type") in ("암퇘지", "일반"):
            return str(row["sex_type"])
        return PorkStandardizer.detect_sex_type(
            raw_name=str(row.get("raw_name") or ""),
            raw_cat=str(row.get("raw_category") or ""),
            sub_spec=str(row.get("sub_spec") or ""),
            source_site=str(row.get("source_site") or ""),
            product_id=str(row.get("product_id") or ""),
        )

    @staticmethod
    def map_standard_part_step1(
        raw_name: str,
        raw_cat: str,
        skin_hint: str = "",
        source_site: str = "",
        product_id: str = "",
    ) -> Tuple[str, str]:
        """
        [1차 매핑 + 2차 부위/미박·박피 교차검증]
        - 1단계: 4대 플랫폼별 고유 카테고리/품목코드(탑미트 ca_id, 금천미트 dispCtgNo·artcCd, 미트박스 itemCatName, 미트프렌즈 축종/품목) 우선 매핑
        - 2단계: 상품명 및 스펙 텍스트를 교차 검증하여
          * 삼겹살 vs 미박삼겹살 (오겹살/미박삼겹) 완벽 분리
          * 앞다리살(전지) = 앞다리 = 전지 vs 미박앞다리살(미박전지) = 미박앞다리 = 미박전지 = 피전지 완벽 분리
          * 뒷다리살(후지) = 뒷다리 = 후지 vs 미박뒷다리살(미박후지) = 미박뒷다리 = 미박후지 = 피후지 완벽 분리
          * 사태 vs 미박사태(미사태/앞미박사태/뒷미박사태/미박알사태) 완벽 분리
          * 특수부위/부산물 통합 명칭 폐기 -> 사골, 등뼈, 장족, 장족(뒤), 단족/미니족, 오돌뼈, 지방돈피/돼지껍데기, 돈지방(A지방), 잡육, 꼬리, 머리고기/내장으로 정밀 분류
        """
        clean_cat = re.sub(r"^(한돈/|금천한돈/|국산돼지/|냉장/|냉동/|세절/)+", "", str(raw_cat or "")).strip()
        clean_cat_no_grade = re.sub(r"\[암퇘지\]$", "", clean_cat).strip()
        clean_name = re.sub(r"\([^)]*(?:제거|미포함|제외)[^)]*\)", " ", str(raw_name or ""))
        clean_name = re.sub(
            r"(?:머리|날개|지방|근막|꽃살|항정|양막|돈피|껍데기|오돌뼈|연골)[,\s_/=]*(?:완벽)?(?:제거|미포함|제외)",
            " ",
            clean_name,
        )
        target = f"{clean_cat_no_grade} {clean_name}".strip()
        pid_parts = str(product_id or "").split("_")
        pid_prefix = pid_parts[0].strip() if pid_parts else ""
        pid_suffix = pid_parts[1].strip() if len(pid_parts) > 1 else ""

        is_mibak_text = bool(
            re.search(
                r"미박|오겹|피삼겹|피전지|스킨전지|탕박전지|피후지|스킨후지|탕박후지|미사태|앞미사태|뒷미사태",
                target,
            )
            or "미박" in str(skin_hint)
        )
        skin_type = "미박" if is_mibak_text else ("박피" if "박피" in f"{target} {skin_hint}" else "일반")

        mapped_part: Optional[str] = None
        mapped_skin: str = skin_type

        # -------------------------------------------------------------------
        # [Step 1-A] 4대 플랫폼별 고유 카테고리/품목코드 기반 1차 직접 매핑
        # -------------------------------------------------------------------
        if source_site == "탑미트" and pid_prefix in TOPMEAT_CA_ID_MAP:
            mapped_part, mapped_skin = TOPMEAT_CA_ID_MAP[pid_prefix]
        elif source_site == "금천미트" and pid_suffix in EKCM_LEAF_CTG_MAP:
            mapped_part, mapped_skin = EKCM_LEAF_CTG_MAP[pid_suffix]
        elif source_site == "금천미트" and pid_prefix in EKCM_ARTC_EXACT_MAP:
            mapped_part, mapped_skin = EKCM_ARTC_EXACT_MAP[pid_prefix]
        elif source_site == "금천미트" and len(pid_prefix) >= 6:
            prefix6 = pid_prefix[:6]
            if prefix6 == "100501":
                mapped_part, mapped_skin = "안심", "일반"
            elif prefix6 == "100502":
                mapped_part, mapped_skin = ("가브리살", "일반") if "덧살" in target else ("등심", "일반")
            elif prefix6 == "100503":
                if "등갈비" in target:
                    mapped_part, mapped_skin = "등갈비", "일반"
                elif "갈매기" in target:
                    mapped_part, mapped_skin = "갈매기살", "일반"
                elif is_mibak_text:
                    mapped_part, mapped_skin = "미박삼겹살", "미박"
                else:
                    mapped_part, mapped_skin = "삼겹살", "박피"
            elif prefix6 == "100504":
                mapped_part, mapped_skin = "목살", "일반"
            elif prefix6 == "100505":
                if "항정" in target:
                    mapped_part, mapped_skin = "항정살", "일반"
                elif is_mibak_text:
                    mapped_part, mapped_skin = "미박앞다리살(미박전지)", "미박"
                else:
                    mapped_part, mapped_skin = "앞다리살(전지)", "박피"
            elif prefix6 == "100506":
                if is_mibak_text:
                    mapped_part, mapped_skin = "미박뒷다리살(미박후지)", "미박"
                else:
                    mapped_part, mapped_skin = "뒷다리살(후지)", "박피"
            elif prefix6 == "100507":
                mapped_part, mapped_skin = "갈비", "일반"
            elif prefix6 == "100509":
                mapped_part, mapped_skin = "머리고기/내장", "일반"
        elif source_site == "미트박스" and clean_cat_no_grade in MEATBOX_ITEM_CAT_MAP:
            mapped_part, mapped_skin = MEATBOX_ITEM_CAT_MAP[clean_cat_no_grade]
        elif source_site == "미트프렌즈" and clean_cat_no_grade in MEATFRIENDS_CAT_MAP:
            if clean_cat_no_grade == "장족" and re.search(r"장족\(뒤\)|뒷장족|손질\s*뒷장족", target):
                mapped_part, mapped_skin = "장족(뒤)", "미박"
            else:
                mapped_part, mapped_skin = MEATFRIENDS_CAT_MAP[clean_cat_no_grade]

        # -------------------------------------------------------------------
        # [Step 1-B] 카테고리명 & 상품명 정규식 기반 정밀 분류 (복합/부산물/세절 카테고리 대응)
        # -------------------------------------------------------------------
        if mapped_part is None:
            if re.search(r"사골|돈사골", target):
                mapped_part, mapped_skin = "사골", "일반"
            elif re.search(r"단족|미니족|뒷미니족|앞미니족|아기족|벌족", target):
                mapped_part, mapped_skin = "단족/미니족", "미박"
            elif re.search(r"장족\(뒤\)|뒷장족|손질\s*뒷장족", target):
                mapped_part, mapped_skin = "장족(뒤)", "미박"
            elif re.search(r"장족\(앞\)|앞장족|장족", target):
                if re.search(r"장족\(뒤\)|뒷장족", target):
                    mapped_part, mapped_skin = "장족(뒤)", "미박"
                else:
                    mapped_part, mapped_skin = "장족", "미박"
            elif re.search(r"오돌뼈|전지연골|돼지연골|\b연골\b", target) and not re.search(r"삼겹", clean_cat_no_grade):
                mapped_part, mapped_skin = "오돌뼈", "일반"
            elif re.search(r"지방돈피|돈피|돼지껍데기|껍데기", target):
                mapped_part, mapped_skin = "지방돈피/돼지껍데기", "미박"
            elif re.search(r"지방\(A\)|A지방|등지방|판지방|돈우지|돼지지방", target):
                mapped_part, mapped_skin = "돈지방(A지방)", "일반"
            elif re.search(r"등뼈|목뼈|척뼈|모듬뼈|통뼈", target) and not re.search(r"등갈비|뼈삼겹|뼈\s*삼겹|뼈등심", target):
                mapped_part, mapped_skin = "등뼈", "일반"
            elif re.search(r"잡육", target):
                mapped_part, mapped_skin = "잡육", "일반"
            elif re.search(r"꼬리|돼지꼬리|꼬리살", target):
                mapped_part, mapped_skin = "꼬리", "일반"
            elif re.search(
                r"돈두롤|돈두|통머리|발골머리|머리차돌|돼지머리|뽈항정|두항정|볼살|뽈살|두볼살|덜미살|관자살|뒷고기|혀밑살|"
                r"소창|염통|막창|대창|오소리|애기보|돼지간|돼지허파|돼지귀|돼지코|돼지혀|돼지유통|돼지콩팥|돼지지라|"
                r"돈갈비살|돈토시살|토시살|지육\(암퇘지\)|지육\(거세\)",
                target,
            ):
                mapped_part, mapped_skin = "머리고기/내장", "일반"
            elif re.search(r"등갈비|쪽갈비|스페어립|폭립", target):
                mapped_part, mapped_skin = "등갈비", "일반"
            elif re.search(r"등심덧살|가브리살|가브리", target) and not re.search(r"지등심", target):
                mapped_part, mapped_skin = "가브리살", "일반"
            elif re.search(r"갈매기살|갈매기", target):
                mapped_part, mapped_skin = "갈매기살", "일반"
            elif re.search(r"항정살|항정", target):
                mapped_part, mapped_skin = "항정살", "일반"
            elif re.search(r"미박\s*삼겹|미박삼겹|오겹살|오겹|피삼겹", target):
                mapped_part, mapped_skin = "미박삼겹살", "미박"
            elif re.search(r"삼겹살|삼겹|박피삼겹|롤삼겹|꽃삼겹", target):
                mapped_part, mapped_skin = ("미박삼겹살", "미박") if is_mibak_text else ("삼겹살", "박피")
            elif re.search(r"목살|목심|알목심|정제목심|숄더랙", target):
                mapped_part, mapped_skin = "목살", "일반"
            elif re.search(r"미박\s*앞다리|미박앞다리|미박\s*전지|미박전지|피전지|스킨전지|탕박전지", target):
                mapped_part, mapped_skin = "미박앞다리살(미박전지)", "미박"
            elif re.search(r"앞다리살|앞다리|전지|박피전지|스키너\s*앞다리", target):
                mapped_part, mapped_skin = ("미박앞다리살(미박전지)", "미박") if is_mibak_text else ("앞다리살(전지)", "박피")
            elif re.search(r"미박\s*뒷다리|미박뒷다리|미박\s*후지|미박후지|피후지|스킨후지|탕박후지", target):
                mapped_part, mapped_skin = "미박뒷다리살(미박후지)", "미박"
            elif re.search(r"뒷다리살|뒷다리|후지|박피후지", target):
                mapped_part, mapped_skin = ("미박뒷다리살(미박후지)", "미박") if is_mibak_text else ("뒷다리살(후지)", "박피")
            elif re.search(r"갈비|돼지갈비|찜갈비|LA갈비", target):
                mapped_part, mapped_skin = "갈비", "일반"
            elif re.search(r"안심", target):
                mapped_part, mapped_skin = "안심", "일반"
            elif re.search(r"등심|지등심|알등심|돈등심|뼈등심|스킨등심", target):
                mapped_part, mapped_skin = "등심", "일반"
            elif re.search(r"미박사태|미사태|앞미박사태|뒷미박사태|앞미사태|뒷미사태|미박알사태|미박앞사태|미박뒷사태", target):
                mapped_part, mapped_skin = "미박사태", "미박"
            elif re.search(r"사태|아롱사태|전사태|후사태|알사태|앞사태|뒷사태", target):
                mapped_part, mapped_skin = ("미박사태", "미박") if is_mibak_text else ("사태", "박피")
            else:
                mapped_part, mapped_skin = "머리고기/내장", "일반"

        # -------------------------------------------------------------------
        # [Step 2-A] 1차 코드 매핑 결과에 대한 상품명·스펙 2차 교차 검증 및 자동 보정
        #   - 미박앞다리살 = 미박전지 / 앞다리살 = 전지 교차 검증
        #   - 미박뒷다리살 = 미박후지 / 뒷다리살 = 후지 교차 검증
        #   - 미박삼겹살 = 오겹살 / 삼겹살 교차 검증
        # -------------------------------------------------------------------
        if re.search(r"돈두롤|머리차돌", target):
            mapped_part, mapped_skin = "머리고기/내장", "일반"
        elif mapped_part == "삼겹살" and is_mibak_text:
            mapped_part, mapped_skin = "미박삼겹살", "미박"
        elif mapped_part == "미박삼겹살" and "박피" in target and not is_mibak_text:
            mapped_part, mapped_skin = "삼겹살", "박피"
        elif mapped_part == "앞다리살(전지)" and is_mibak_text:
            mapped_part, mapped_skin = "미박앞다리살(미박전지)", "미박"
        elif mapped_part == "미박앞다리살(미박전지)" and "박피" in target and not is_mibak_text:
            mapped_part, mapped_skin = "앞다리살(전지)", "박피"
        elif mapped_part == "뒷다리살(후지)" and is_mibak_text:
            mapped_part, mapped_skin = "미박뒷다리살(미박후지)", "미박"
        elif mapped_part == "미박뒷다리살(미박후지)" and "박피" in target and not is_mibak_text:
            mapped_part, mapped_skin = "뒷다리살(후지)", "박피"
        elif mapped_part == "사태" and is_mibak_text:
            mapped_part, mapped_skin = "미박사태", "미박"
        elif mapped_part == "장족" and re.search(r"장족\(뒤\)|뒷장족|손질\s*뒷장족", target):
            mapped_part, mapped_skin = "장족(뒤)", "미박"

        return mapped_part, mapped_skin

    @staticmethod
    def validate_step2(
        source_site: str,
        raw_name: str,
        raw_cat: str,
        raw_proc: str,
        brand: str,
        standard_part: str,
        skin_type: str,
        unit_price_kg: int,
        product_id: str = "",
        sex_hint: str = "",
    ) -> Tuple[str, str, str]:
        combined = f"{raw_name} {raw_cat} {raw_proc} {brand}"
        pid_prefix = str(product_id or "").split("_")[0].strip()

        is_processed = False
        if source_site == "미트박스" and raw_proc and raw_proc not in ("원물", "지육"):
            is_processed = True
        elif source_site == "미트프렌즈" and re.search(r"세절|분쇄|슬라이스|다짐|절단|소분|양념|육가공", str(raw_proc or "")):
            is_processed = True
        if "육마트(키로육)" in brand or "키로육" in raw_name:
            is_processed = True

        # 지방 정선 두께 표기('지방두께 3mm', '지방 10mm 작업', '뒷다리살 3mm', 'mm등심')를 제거한 뒤 세절/소분 정규식 검사
        scrubbed_for_cut = FAT_TRIM_SPEC_REGEX.sub(" ", combined)
        if PROCESSED_CUT_PATTERN.search(scrubbed_for_cut):
            is_processed = True

        processing_type = "세절/가공육" if is_processed else "원물/박스육"

        sex_type = PorkStandardizer.detect_sex_type(
            raw_name=raw_name,
            raw_cat=raw_cat,
            source_site=source_site,
            product_id=product_id,
            sex_hint=sex_hint,
        )

        cleaned_special = re.sub(r"미추리\s*짧음", "", combined)
        if re.search(r"모돈|후보돈", combined):
            sub_spec = "모돈"
        elif sex_type == "암퇘지":
            sub_spec = "암퇘지"
        elif re.search(
            r"수육용|보쌈용|삼겹\s*미추리|미추리\s*삼겹|\(미추리\)|정제목심|두판|더좋은|풍미삼겹|삼겹양지|"
            r"돈차돌|립삼겹|돈마호크|설하목심|갈비목심|갈비삼겹|뼈\s*삼겹|통삼겹|복박이|오돌삼겹|지등심|"
            r"듀록|YBD|난축맛돈|얼룩도야지",
            cleaned_special,
        ):
            sub_spec = "수육/특수정선"
        else:
            sub_spec = "일반"

        code_tag = f"코드:{pid_prefix}·" if pid_prefix else ""
        sex_tag = "암퇘지전용분리" if sex_type == "암퇘지" else sub_spec
        if processing_type == "세절/가공육":
            mapping_status = f"2차검증완료({code_tag}세절/소분가공/{sex_tag})"
        elif sub_spec == "모돈":
            mapping_status = f"2차검증완료({code_tag}모돈원물 분리)"
        elif sex_type == "암퇘지":
            mapping_status = f"2차검증완료({code_tag}암퇘지 별도분리/{standard_part})"
        elif sub_spec == "수육/특수정선":
            mapping_status = f"2차검증완료({code_tag}특수정선/수육 분리)"
        elif standard_part in ("미박삼겹살", "미박앞다리살(미박전지)", "미박뒷다리살(미박후지)", "미박사태"):
            mapping_status = f"2차검증완료({code_tag}미박품목분리/{sub_spec})"
        elif standard_part in ("삼겹살", "앞다리살(전지)", "뒷다리살(후지)", "사태"):
            mapping_status = f"2차검증완료({code_tag}박피·일반분리/{sub_spec})"
        else:
            mapping_status = f"2차검증완료({code_tag}{standard_part}/{sub_spec})"

        return processing_type, sub_spec, mapping_status

    @staticmethod
    def classify_prod_date_bucket(min_proc_date: str, keeping: str, ref_date_str: str = "") -> str:
        """
        생산일자(제조일/가공일)를 기준으로 신선도 구간 분류:
        - '당월 신상품(최근 14일 이내)': 최근 14일 이내 생산된 신선 물량
        - '전월/일반 생산(15~45일)': 생산 후 15~45일 경과 물량
        - '장기보관/기한임박(45일 경과)': 생산 후 45일 초과 경과 물량 (노후 재고 덤핑 착시 주의 구간)
        - '주문당일생산/상시': 당일작업 또는 상세일정 상시 출고
        """
        if not min_proc_date or min_proc_date in ("-", "당일작업", "상시생산"):
            return "주문당일생산/상시"
        m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", str(min_proc_date).strip())
        if not m:
            return "주문당일생산/상시"
        try:
            p_dt = datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            ref_dt = datetime.strptime(ref_date_str[:10], "%Y-%m-%d") if ref_date_str else datetime.now()
            age_days = (ref_dt - p_dt).days
            if age_days <= 14:
                return "당월 신상품(최근 14일)"
            elif age_days <= 45:
                return "전월/일반 생산(15~45일)"
            else:
                return "장기보관/기한임박(45일 경과)"
        except Exception:
            return "주문당일생산/상시"


# ---------------------------------------------------------------------------
# 4대 축산 플랫폼 전수 크롤러 (브랜드 + 생산일자 최저/최고가 + 하이퍼링크 완벽 파싱)
# ---------------------------------------------------------------------------
def crawl_meatbox(
    session: requests.Session, run_id: int, collected_at: str, collected_date: str
) -> Tuple[List[PorkSnapshotRecord], Dict[str, int]]:
    logger.info("=== [1/4] 미트박스(Meatbox) 한돈 전수 수집 및 옵션 생산일자 파싱 시작 ===")
    stats = {"raw": 0, "dropped": 0, "valid": 0, "wonmul": 0, "processed": 0}
    deliveries = [
        {"productSeq": 184784, "type": "fresh", "distTypeCd": "AD03", "isJejuAddr": False},
        {"productSeq": 409260, "type": "meatmart", "distTypeCd": "AD03", "isJejuAddr": False},
    ]
    headers = {
        "Content-Type": "application/json",
        "Origin": "https://www.meatbox.co.kr",
        "Referer": MEATBOX_TARGET_URL,
    }

    raw_items: List[Dict[str, Any]] = []
    seen_seqs: Set[str] = set()
    page = 1
    page_size = 40

    while page <= 40:
        try:
            resp = session.post(
                MEATBOX_API_URL,
                params={"page": page, "size": page_size},
                json={"displayCategorySeq": [10010], "deliveries": deliveries, "operator": "or"},
                headers=headers,
                timeout=15,
            )
            if resp.status_code != 200:
                break
            data = resp.json()
            search_node = data.get("data", {}).get("search", {})
            items = search_node.get("items", []) or []
            if not items:
                break
            added = 0
            for it in items:
                seq = str(it.get("productSeq") or "")
                if seq and seq not in seen_seqs:
                    seen_seqs.add(seq)
                    raw_items.append(it)
                    added += 1
            total_cnt = search_node.get("total")
            if len(items) < page_size or (isinstance(total_cnt, int) and len(raw_items) >= total_cnt) or added == 0:
                break
            page += 1
            time.sleep(random.uniform(0.10, 0.22))
        except Exception as exc:
            logger.warning("미트박스 목록 페이지(%d) 조회 경고: %s", page, exc)
            break

    stats["raw"] = len(raw_items)
    records: List[PorkSnapshotRecord] = []

    for it in raw_items:
        seq = str(it.get("productSeq") or "")
        name = str(it.get("productName") or "").strip()
        raw_brand = str(it.get("brandName") or "").strip()
        item_kind = str(it.get("itemKindName") or "").strip()
        item_cat = str(it.get("itemCatName") or "").strip()
        raw_cat = f"{item_kind}/{item_cat}".strip("/")
        keeping_raw = str(it.get("keepingName") or "").strip()
        proc_raw = str(it.get("procName") or "").strip()
        event_msg = str(it.get("eventMsg") or "").strip()

        is_excl, _ = PorkStandardizer.is_excluded_black_or_imported(f"{name} {raw_brand} {raw_cat} {event_msg}")
        if is_excl:
            stats["dropped"] += 1
            continue

        brand = PorkStandardizer.clean_brand_name(raw_brand, name, "미트박스")

        price_text = str(it.get("priceText") or "").strip()
        unit_type_text = str(it.get("priceUnitTypeText") or "").strip()
        disp_kg_text = str(it.get("displayPriceForKg") or "").strip()

        unit_price_kg = 0
        norm_method = "1kg단가 직수집"
        kg_match = re.search(r"kg당\s*([\d,]+)\s*원", disp_kg_text)
        if kg_match:
            unit_price_kg = int(kg_match.group(1).replace(",", ""))
            norm_method = "displayPriceForKg 환산" if unit_type_text != "(kg당)" else "1kg단가 직수집"
        else:
            m_100g = re.search(r"100g당\s*([\d,]+)\s*원", f"{disp_kg_text} {price_text} {unit_type_text}")
            if m_100g:
                unit_price_kg = int(m_100g.group(1).replace(",", "")) * 10
                norm_method = "100g단가×10 환산"
            else:
                p_match = re.search(r"([\d,]+)\s*원", price_text)
                if p_match:
                    raw_p = int(p_match.group(1).replace(",", ""))
                    if unit_type_text == "(kg당)" or raw_p <= 60000:
                        unit_price_kg = raw_p
                    else:
                        m_w = re.search(r"(\d+(?:\.\d+)?)\s*kg", f"{disp_kg_text} {name}", re.I)
                        if m_w and float(m_w.group(1)) > 0:
                            unit_price_kg = int(round(raw_p / float(m_w.group(1))))
                            norm_method = "박스총액÷중량 환산"
                        else:
                            unit_price_kg = raw_p

        if unit_price_kg <= 0:
            continue

        price_display = f"{price_text} {unit_type_text}".strip()
        if disp_kg_text:
            price_display += f" [{disp_kg_text}]"

        is_sold_out_flag = bool(it.get("isSoldOut")) or str(it.get("soldOutYn") or "").upper() == "Y"
        status_name = str(it.get("statusName") or "판매중").strip()
        stock_val = it.get("stockAmount")
        try:
            stock_qty = int(str(stock_val).replace(",", "").strip()) if stock_val is not None else 0
        except ValueError:
            stock_qty = 0

        is_sold_out = is_sold_out_flag or ("품절" in status_name) or (stock_qty <= 0)
        is_virtual_stock = stock_qty >= 9999
        real_stock_qty = 0 if (is_sold_out or is_virtual_stock) else stock_qty
        status_text = (
            "일시품절"
            if is_sold_out
            else (f"판매중(상시:{stock_qty:,})" if is_virtual_stock else f"판매중(재고:{stock_qty:,}박스)")
        )

        is_geunbon = 1 if (("근본한돈" in name) or ("근본한돈" in brand)) else 0
        platform_group = "미트박스 - 근본한돈" if is_geunbon else "미트박스 - 일반 한돈"

        keeping = PorkStandardizer.normalize_keeping(keeping_raw, name, raw_cat)
        std_part, skin_type = PorkStandardizer.map_standard_part_step1(
            name, raw_cat, source_site="미트박스", product_id=seq
        )
        proc_type, sub_spec, map_status = PorkStandardizer.validate_step2(
            "미트박스", name, raw_cat, proc_raw, brand, std_part, skin_type, unit_price_kg, product_id=seq
        )

        if proc_type == "원물/박스육":
            stats["wonmul"] += 1
        else:
            stats["processed"] += 1

        records.append(
            PorkSnapshotRecord(
                run_id=run_id,
                collected_at=collected_at,
                collected_date=collected_date,
                item_key=f"미트박스:{seq}",
                platform_group=platform_group,
                source_site="미트박스",
                is_geunbon=is_geunbon,
                product_id=seq,
                raw_name=name,
                raw_category=raw_cat,
                brand=brand,
                keeping=keeping,
                standard_part=std_part,
                skin_type=skin_type,
                sub_spec=sub_spec,
                processing_type=proc_type,
                mapping_status=map_status,
                unit_price_kg=unit_price_kg,
                max_price_kg=unit_price_kg,
                avg_price_kg=unit_price_kg,
                price_norm_method=norm_method,
                price_display=price_display,
                min_proc_date="-",
                max_proc_date="-",
                proc_date_range="상시생산",
                exp_date_range="-",
                prod_date_bucket="주문당일생산/상시",
                is_sold_out=1 if is_sold_out else 0,
                status_text=status_text,
                stock_qty=real_stock_qty,
                option_cnt=0 if is_sold_out else 1,
                product_url=f"https://www.meatbox.co.kr/fo/product/productViewPage.do?productSeq={seq}",
            )
        )

    # 원물/박스육 판매중 상품(근본한돈 전체 + 일반 원물 주요 상품)에 대해 상세 옵션 API를 조회해 생산일자와 최저/최고가 추출
    wonmul_active = [r for r in records if not r.is_sold_out and r.processing_type == "원물/박스육"]
    logger.info("미트박스 판매중 원물 %d건에 대해 상세 옵션(생산일자·유통기한·옵션별 최저/최고가) 병렬 조회 중...", len(wonmul_active))

    def _enrich_mb_opt(rec: PorkSnapshotRecord) -> None:
        try:
            r_opt = session.get(
                f"https://mb-api.meatbox.co.kr/product/api/v1/items/{rec.product_id}/retrieves",
                headers=headers,
                timeout=10,
            )
            if r_opt.status_code != 200:
                return
            d_opt = r_opt.json().get("data", {}) or {}
            opts = d_opt.get("sortedPricePerKgItems") or d_opt.get("productOptions") or []
            if not opts:
                return
            in_stk = [o for o in opts if int(o.get("restAmount") or o.get("amount") or 0) > 0]
            target_opts = in_stk if in_stk else opts

            pairs: List[Tuple[int, str, str]] = []
            for o in target_opts:
                p_kg = int(o.get("finalPricePerKg") or rec.unit_price_kg)
                if p_kg <= 0:
                    p_kg = rec.unit_price_kg
                p_dt = str(o.get("procDate") or o.get("bcrDate") or "-").strip()
                e_dt = str(o.get("expDate") or "-").strip()
                pairs.append((p_kg, p_dt, e_dt))

            if not pairs:
                return
            pairs.sort(key=lambda x: (x[0], x[1]))
            rec.unit_price_kg = pairs[0][0]
            rec.min_proc_date = pairs[0][1] if pairs[0][1] else "-"
            rec.max_price_kg = pairs[-1][0]
            rec.max_proc_date = pairs[-1][1] if pairs[-1][1] else "-"
            rec.avg_price_kg = int(round(sum(x[0] for x in pairs) / len(pairs)))
            rec.option_cnt = len(target_opts)

            p_dates = sorted({x[1] for x in pairs if x[1] and x[1] != "-"})
            e_dates = sorted({x[2] for x in pairs if x[2] and x[2] != "-"})
            if p_dates:
                rec.proc_date_range = p_dates[0] if len(p_dates) == 1 else f"{p_dates[0]}~{p_dates[-1]}"
                if rec.min_proc_date == "-":
                    rec.min_proc_date = p_dates[0]
                if rec.max_proc_date == "-":
                    rec.max_proc_date = p_dates[-1]
            if e_dates:
                rec.exp_date_range = e_dates[0] if len(e_dates) == 1 else f"{e_dates[0]}~{e_dates[-1]}"

            rec.prod_date_bucket = PorkStandardizer.classify_prod_date_bucket(
                rec.min_proc_date, rec.keeping, collected_date
            )
        except Exception:
            pass

    with ThreadPoolExecutor(max_workers=14) as ex:
        list(ex.map(_enrich_mb_opt, wonmul_active))

    stats["valid"] = len(records)
    logger.info("미트박스 수집 완료: 유효 %d건 (근본한돈 %d건, 일반 %d건)", len(records), sum(r.is_geunbon for r in records), sum(1 for r in records if not r.is_geunbon))
    return records, stats


def build_ekcm_product_detail_url(
    goods_no: str, artc_cd: str, disp_ctg_no: str, disp_ctg_nm: str = ""
) -> str:
    """금천미트 개별 상품상세페이지(/pd/productDetail) 정밀 URL 생성"""
    g_no = str(goods_no or "").strip()
    a_cd = str(artc_cd or "").strip()
    c_no = str(disp_ctg_no or "").strip()
    if g_no and a_cd and c_no:
        q_nm = f"&dispCtgNm={quote(disp_ctg_nm)}" if disp_ctg_nm else ""
        return f"{EKCM_DETAIL_PAGE_BASE}?goodsNo={g_no}&artcCd={a_cd}&dispCtgNo={c_no}&uprTopDispCtgNo=3{q_nm}"
    if c_no:
        return f"https://www.ekcm.co.kr/pd/product?dispCtgNo=31&uprTopDispCtgNo=3&leafCtgNo={c_no}&dispCtgNoList={c_no}"
    return EKCM_TARGET_URL


def verify_ekcm_detail_api(
    session: requests.Session, goods_no: str, disp_ctg_no: str, headers: Optional[Dict[str, str]] = None
) -> bool:
    """금천미트 상품상세 백엔드 API(goodsDetail)를 호출하여 하이퍼링크 클릭 시 실제 상품이 표시되는지 검증"""
    g_no = str(goods_no or "").strip()
    c_no = str(disp_ctg_no or "").strip()
    if not g_no or not c_no:
        return False
    req_headers = headers or {
        "Content-Type": "application/json",
        "Origin": "https://www.ekcm.co.kr",
        "Referer": "https://www.ekcm.co.kr/",
    }
    try:
        resp = session.post(f"{EKCM_DETAIL_API_BASE}/{g_no}/{c_no}/N", headers=req_headers, json={}, timeout=12)
        if resp.status_code != 200:
            return False
        body = resp.json()
        if str(body.get("code") or "") != "0000":
            return False
        g_info = (body.get("payload") or {}).get("goodsInfo") or {}
        return str(g_info.get("goodsNo") or "") == g_no
    except Exception:
        return False


def crawl_ekcm(
    session: requests.Session, run_id: int, collected_at: str, collected_date: str
) -> Tuple[List[PorkSnapshotRecord], Dict[str, int]]:
    logger.info("=== [2/4] 금천미트(ekcm) 국내산 돈육 전수 수집·생산일자 파싱 및 상품상세 링크 검증 시작 ===")
    stats = {"raw": 0, "dropped": 0, "valid": 0, "wonmul": 0, "processed": 0, "boxes": 0, "url_verified": 0}
    headers = {
        "Content-Type": "application/json",
        "Origin": "https://www.ekcm.co.kr",
        "Referer": "https://www.ekcm.co.kr/",
    }

    try:
        r_cat = session.get(EKCM_CATEGORY_API_URL, headers=headers, timeout=15)
        groups = r_cat.json().get("leftCateCornCont", []) or []
    except Exception as exc:
        logger.error("금천미트 카테고리 조회 오류: %s", exc)
        return [], stats

    leaf_tasks: List[Tuple[str, Dict[str, Any]]] = []
    for g in groups:
        keeping_grp = str(g.get("ctgTypCdNm") or "").strip()
        for c in g.get("displayCategoryList", []) or []:
            disp_ctg_nm = str(c.get("dispCtgNm") or "").strip()
            cnt = int(c.get("goodsCnt") or 0)
            if cnt <= 0:
                continue
            is_excl, _ = PorkStandardizer.is_excluded_black_or_imported(disp_ctg_nm)
            if is_excl:
                stats["dropped"] += cnt
                continue
            leaf_tasks.append((keeping_grp, c))

    def _fetch_leaf(task: Tuple[str, Dict[str, Any]]) -> List[Dict[str, Any]]:
        keeping_grp, c = task
        disp_ctg_nm = str(c.get("dispCtgNm") or "").strip()
        disp_ctg_no = str(c.get("dispCtgNo") or "").strip()
        ctg_list = [x.strip() for x in str(c.get("dispCtgNoList") or disp_ctg_no).split(",") if x.strip()]
        payload = {
            "dispCtgNoList": ctg_list,
            "brandNoList": [], "lsprdGrdCdList": [], "homeCdList": [], "ppYmdList": [],
            "strgMthdGbCdList": [], "workMethTypCdList": [], "deliProcTypCdList": [],
            "recomBkindList": [], "artcGbCdList": [], "qualityList": [], "insfatGrdList": [],
            "mffldList": [], "estNoList": [], "sortTpCd": "", "pageNo": 1, "pageSize": 2000,
            "aplyPsbMediaCd": "02", "curCtgNo": disp_ctg_no, "noDispCtgRegYn": "N", "mbrNo": "",
        }
        try:
            resp = session.post(EKCM_GOODS_API_URL, headers=headers, json=payload, timeout=20)
            if resp.status_code != 200:
                return []
            boxes = resp.json().get("payload", []) or []
            for b in boxes:
                b["_leafCtgNm"] = disp_ctg_nm
                b["_leafCtgNo"] = disp_ctg_no
                b["_keepingGrp"] = keeping_grp
            return boxes
        except Exception:
            return []

    all_boxes: List[Dict[str, Any]] = []
    seen_gno: Set[str] = set()
    with ThreadPoolExecutor(max_workers=10) as ex:
        for boxes in ex.map(_fetch_leaf, leaf_tasks):
            for b in boxes:
                gno = str(b.get("goodsNo") or "")
                if not gno or gno in seen_gno:
                    continue
                seen_gno.add(gno)
                stats["raw"] += 1
                disp_ctg_nm = str(b.get("_leafCtgNm") or "")
                artc_nm = str(b.get("artcNm") or "")
                goods_nm = str(b.get("goodsNm") or "")
                brand_nm = str(b.get("brandNm") or "")
                is_excl, _ = PorkStandardizer.is_excluded_black_or_imported(f"{disp_ctg_nm} {artc_nm} {goods_nm} {brand_nm}")
                if is_excl:
                    stats["dropped"] += 1
                    continue
                all_boxes.append(b)

    stats["boxes"] = len(all_boxes)

    def _fmt_date(raw_dt: Any) -> str:
        s = re.sub(r"[^\d]", "", str(raw_dt or ""))
        return f"{s[:4]}-{s[4:6]}-{s[6:8]}" if len(s) == 8 else "-"

    sku_groups: Dict[Tuple[str, ...], Dict[str, Any]] = {}
    for box in all_boxes:
        artc_cd = str(box.get("artcCd") or "").strip()
        artc_nm = str(box.get("artcNm") or "").strip()
        goods_nm = str(box.get("goodsNm") or "").strip()
        raw_brand = str(box.get("brandNm") or "금천한돈").strip()
        keeping_nm = str(box.get("strgMthdGbNm") or box.get("_keepingGrp") or "냉장").strip()
        leaf_nm = str(box.get("_leafCtgNm") or "").strip()
        leaf_no = str(box.get("_leafCtgNo") or "").strip()
        box_leaf_no = str(box.get("leafCtgNo") or box.get("dispCtgNo") or leaf_no).strip()
        skng_nm = str(box.get("skngStatNm") or "").strip()
        sex_nm = str(box.get("lventSexGbNm") or "").strip()
        goods_no_str = str(box.get("goodsNo") or "").strip()

        ucost = int(box.get("realUcost") or box.get("supPcost") or box.get("ucost") or 0)
        if ucost <= 0:
            continue
        stk_qty = int(box.get("stkQty") or 0)
        sale_fnsh = str(box.get("saleFnshYn") or "N").upper() == "Y"
        is_avail = (not sale_fnsh) and (stk_qty > 0)
        pp_dt = _fmt_date(box.get("ppYmd") or box.get("ppDt"))
        exp_dt = _fmt_date(box.get("expYmd") or box.get("expDt"))

        key = (artc_cd, artc_nm, goods_nm, raw_brand, keeping_nm, leaf_nm, skng_nm, sex_nm)
        if key not in sku_groups:
            sku_groups[key] = {
                "artc_cd": artc_cd,
                "goods_nm": goods_nm,
                "artc_nm": artc_nm,
                "raw_brand": raw_brand,
                "keeping_nm": keeping_nm,
                "leaf_ctg_nm": leaf_nm,
                "leaf_ctg_no": leaf_no,
                "skng_nm": skng_nm,
                "sex_nm": sex_nm,
                "boxes": [],
                "sample_goods_no": goods_no_str,
                "sample_box_leaf_no": box_leaf_no,
            }
        sku_groups[key]["boxes"].append(
            {
                "ucost": ucost,
                "stk_qty": stk_qty if is_avail else 0,
                "is_avail": is_avail,
                "pp_dt": pp_dt,
                "exp_dt": exp_dt,
                "goods_no": goods_no_str,
                "box_leaf_no": box_leaf_no,
            }
        )

    records: List[PorkSnapshotRecord] = []
    record_candidates: List[Tuple[PorkSnapshotRecord, str, str, List[Tuple[str, str]], List[Tuple[str, str]]]] = []

    for _, grp in sku_groups.items():
        boxes_list = grp["boxes"]
        avail_boxes = [b for b in boxes_list if b["is_avail"]]
        target_boxes = avail_boxes if avail_boxes else boxes_list
        target_boxes.sort(key=lambda x: (x["ucost"], x["pp_dt"]))

        min_box = target_boxes[0]
        max_box = target_boxes[-1]
        min_ucost = min_box["ucost"]
        min_pp_dt = min_box["pp_dt"]
        min_goods_no = min_box.get("goods_no") or grp.get("sample_goods_no") or ""
        min_leaf_no = min_box.get("box_leaf_no") or grp["leaf_ctg_no"]

        max_ucost = max_box["ucost"]
        max_pp_dt = max_box["pp_dt"]
        max_goods_no = max_box.get("goods_no") or min_goods_no
        max_leaf_no = max_box.get("box_leaf_no") or min_leaf_no

        avg_ucost = int(round(sum(b["ucost"] for b in target_boxes) / len(target_boxes)))

        p_dates = sorted({b["pp_dt"] for b in target_boxes if b["pp_dt"] != "-"})
        e_dates = sorted({b["exp_dt"] for b in target_boxes if b["exp_dt"] != "-"})
        proc_range = "상시생산" if not p_dates else (p_dates[0] if len(p_dates) == 1 else f"{p_dates[0]}~{p_dates[-1]}")
        exp_range = "-" if not e_dates else (e_dates[0] if len(e_dates) == 1 else f"{e_dates[0]}~{e_dates[-1]}")

        stk_qty = sum(b["stk_qty"] for b in avail_boxes)
        opt_cnt = len(avail_boxes)
        is_sold_out = stk_qty <= 0

        goods_nm = grp["goods_nm"]
        artc_nm = grp["artc_nm"]
        leaf_nm = grp["leaf_ctg_nm"]
        keeping_nm = grp["keeping_nm"]
        raw_brand = grp["raw_brand"]
        skng_nm = grp["skng_nm"]
        sex_nm = grp.get("sex_nm") or ""

        full_name = f"[{raw_brand}] {goods_nm} ({artc_nm})"
        brand = PorkStandardizer.clean_brand_name(raw_brand, full_name, "금천미트")
        keeping = PorkStandardizer.normalize_keeping(keeping_nm, full_name, leaf_nm)
        pid = f"{grp['artc_cd']}_{grp['leaf_ctg_no']}"
        std_part, skin_type = PorkStandardizer.map_standard_part_step1(
            full_name, leaf_nm, skin_hint=skng_nm, source_site="금천미트", product_id=pid
        )
        proc_type, sub_spec, map_status = PorkStandardizer.validate_step2(
            "금천미트",
            full_name,
            leaf_nm,
            "원물",
            brand,
            std_part,
            skin_type,
            min_ucost,
            product_id=pid,
            sex_hint=sex_nm,
        )

        if proc_type == "원물/박스육":
            stats["wonmul"] += 1
        else:
            stats["processed"] += 1

        prod_bucket = PorkStandardizer.classify_prod_date_bucket(min_pp_dt, keeping, collected_date)
        min_url = build_ekcm_product_detail_url(min_goods_no, grp["artc_cd"], min_leaf_no, leaf_nm)
        max_url = build_ekcm_product_detail_url(max_goods_no, grp["artc_cd"], max_leaf_no, leaf_nm)

        rec = PorkSnapshotRecord(
            run_id=run_id,
            collected_at=collected_at,
            collected_date=collected_date,
            item_key=f"금천미트:{pid}",
            platform_group="금천미트",
            source_site="금천미트",
            is_geunbon=0,
            product_id=pid,
            raw_name=full_name,
            raw_category=f"{keeping_nm}/{leaf_nm}",
            brand=brand,
            keeping=keeping,
            standard_part=std_part,
            skin_type=skin_type,
            sub_spec=sub_spec,
            processing_type=proc_type,
            mapping_status=map_status,
            unit_price_kg=min_ucost,
            max_price_kg=max_ucost,
            avg_price_kg=avg_ucost,
            price_norm_method="1kg단가 직수집(realUcost)",
            price_display=f"{min_ucost:,}~{max_ucost:,}원/kg" if min_ucost != max_ucost else f"{min_ucost:,}원/kg",
            min_proc_date=min_pp_dt,
            max_proc_date=max_pp_dt,
            proc_date_range=proc_range,
            exp_date_range=exp_range,
            prod_date_bucket=prod_bucket,
            is_sold_out=1 if is_sold_out else 0,
            status_text="일시품절" if is_sold_out else f"판매중(재고:{stk_qty:,}박스)",
            stock_qty=stk_qty,
            option_cnt=opt_cnt,
            product_url=min_url,
            max_product_url=max_url,
        )
        records.append(rec)

        min_cands = [(b["goods_no"], b["box_leaf_no"]) for b in target_boxes if b.get("goods_no")]
        if min_goods_no and grp["leaf_ctg_no"] != min_leaf_no:
            min_cands.append((min_goods_no, grp["leaf_ctg_no"]))
        max_cands = [(b["goods_no"], b["box_leaf_no"]) for b in reversed(target_boxes) if b.get("goods_no")]
        if max_goods_no and grp["leaf_ctg_no"] != max_leaf_no:
            max_cands.append((max_goods_no, grp["leaf_ctg_no"]))
        record_candidates.append((rec, grp["artc_cd"], leaf_nm, min_cands, max_cands))

    # [크롤링 스케줄 내장 검증] 금천미트 전 SKU 상품상세 하이퍼링크(/pd/productDetail) 실시간 응답 검증 및 자동 보정
    def _verify_and_repair_ekcm_rec(
        item_tuple: Tuple[PorkSnapshotRecord, str, str, List[Tuple[str, str]], List[Tuple[str, str]]]
    ) -> bool:
        rec_obj, artc_cd_val, leaf_nm_val, min_cand_list, max_cand_list = item_tuple
        min_ok = False
        chosen_min_pair: Optional[Tuple[str, str]] = None
        for g_no, c_no in min_cand_list[:4]:
            if verify_ekcm_detail_api(session, g_no, c_no, headers):
                rec_obj.product_url = build_ekcm_product_detail_url(g_no, artc_cd_val, c_no, leaf_nm_val)
                min_ok = True
                chosen_min_pair = (g_no, c_no)
                break

        max_ok = False
        if max_cand_list and chosen_min_pair and max_cand_list[0] == chosen_min_pair:
            rec_obj.max_product_url = rec_obj.product_url
            max_ok = True
        else:
            for g_no, c_no in max_cand_list[:4]:
                if verify_ekcm_detail_api(session, g_no, c_no, headers):
                    rec_obj.max_product_url = build_ekcm_product_detail_url(g_no, artc_cd_val, c_no, leaf_nm_val)
                    max_ok = True
                    break
            if not max_ok and min_ok:
                rec_obj.max_product_url = rec_obj.product_url
                max_ok = True
        return min_ok and max_ok

    verified_ok_cnt = 0
    with ThreadPoolExecutor(max_workers=12) as ex:
        for is_ok in ex.map(_verify_and_repair_ekcm_rec, record_candidates):
            if is_ok:
                verified_ok_cnt += 1

    stats["valid"] = len(records)
    stats["url_verified"] = verified_ok_cnt
    logger.info(
        "금천미트 수집 및 하이퍼링크 검증 완료: 유효 SKU %d건 (개별 박스 %d박스 / 상품상세페이지 링크 검증 %d/%d건 정상)",
        len(records),
        stats["boxes"],
        verified_ok_cnt,
        len(records),
    )
    return records, stats


def crawl_meatfriends(
    session: requests.Session, run_id: int, collected_at: str, collected_date: str
) -> Tuple[List[PorkSnapshotRecord], Dict[str, int]]:
    logger.info("=== [3/4] 미트프렌즈(MeatFriends) 한돈 전수 수집 및 상세 생산일자 파싱 시작 ===")
    stats = {"raw": 0, "dropped": 0, "valid": 0, "wonmul": 0, "processed": 0, "options": 0}
    ajax_headers = {"Referer": MEATFRIENDS_TARGET_URL, "X-Requested-With": "XMLHttpRequest"}

    list_items: List[Dict[str, Any]] = []
    seen_prd: Set[str] = set()

    for cat_no, keep_hint_default in [(7, ""), (8, "냉장"), (9, "냉동"), (26, "세절")]:
        page = 1
        while page <= 30:
            try:
                resp = session.post(
                    MEATFRIENDS_LIST_AJAX_URL,
                    data={"catNo": cat_no, "pIdx": page, "mbrTypeGb": "99", "lnb": "Y"},
                    headers=ajax_headers,
                    timeout=15,
                )
                if resp.status_code != 200:
                    break
                soup = BeautifulSoup(resp.text, "html.parser")
                lis = soup.select("ul.prdList > li")
                if not lis:
                    break
                added = 0
                for li in lis:
                    thumb = li.select_one("div.thumb")
                    if not thumb:
                        continue
                    m = re.search(
                        r"fnProductDetail\('(\d+)'\s*,\s*'([^']*)'\s*,\s*'([^']*)'\s*,\s*'([^']*)'\s*,\s*'([^']*)'\)",
                        thumb.get("onclick", ""),
                    )
                    if not m:
                        continue
                    prd_no = m.group(1).strip()
                    if not prd_no or prd_no in seen_prd:
                        continue
                    seen_prd.add(prd_no)
                    unit_span = li.select_one("dd.prdPrice span.unit")
                    unit_txt = unit_span.get_text(" ", strip=True) if unit_span else ""
                    list_items.append(
                        {
                            "prd_no": prd_no,
                            "title": m.group(2).strip(),
                            "list_price": int(m.group(3) or 0),
                            "keep_hint": m.group(5).strip() or keep_hint_default,
                            "slb_stk": int(thumb.get("slbstkcnt") or 0),
                            "unit_txt": unit_txt,
                        }
                    )
                    added += 1
                if len(lis) < 12:
                    break
                page += 1
                time.sleep(random.uniform(0.04, 0.10))
            except Exception:
                break

    stats["raw"] = len(list_items)

    def _fetch_detail(item: Dict[str, Any]) -> Dict[str, Any]:
        prd_no = item["prd_no"]
        detail_url = f"{MEATFRIENDS_DETAIL_URL}?prdNo={prd_no}"
        try:
            r = session.get(detail_url, headers={"Referer": MEATFRIENDS_TARGET_URL}, timeout=12)
            html = r.text if r.status_code == 200 else ""
        except Exception:
            html = ""

        opts: List[Dict[str, Any]] = []
        m_json = re.search(r"var\s+jsonData\s*=\s*'(\[.*?\])';", html, re.DOTALL)
        if m_json:
            try:
                opts = json.loads(m_json.group(1).replace("\r", "").replace("\n", " "))
            except Exception:
                opts = []

        info_map: Dict[str, str] = {}
        if html:
            soup_d = BeautifulSoup(html, "html.parser")
            for tbl in soup_d.select("div.details-1 table"):
                cap = tbl.find("caption")
                if cap and "상품필수정보" in cap.get_text():
                    for tr in tbl.find_all("tr"):
                        for th, td in zip(tr.find_all("th"), tr.find_all("td")):
                            k = th.get_text(" ", strip=True)
                            v = td.get_text(" ", strip=True)
                            if k:
                                info_map[k] = v
                    break
        item["opts"] = opts
        item["info"] = info_map
        return item

    with ThreadPoolExecutor(max_workers=14) as ex:
        futures = [ex.submit(_fetch_detail, it) for it in list_items]
        for fut in as_completed(futures):
            fut.result()

    def _parse_title_weight_kg(title: str) -> float:
        m_k_mult = re.search(r"(\d+(?:\.\d+)?)\s*k(?:g)?\s*\*\s*(\d+)", title, re.I)
        if m_k_mult:
            return float(m_k_mult.group(1)) * float(m_k_mult.group(2))
        m_g_mult = re.search(r"(\d+(?:\.\d+)?)\s*g\s*\*\s*(\d+)", title, re.I)
        if m_g_mult:
            return (float(m_g_mult.group(1)) * float(m_g_mult.group(2))) / 1000.0
        m_k_single = re.search(r"(\d+(?:\.\d+)?)\s*k(?:g)?(?:/|\)|b|\s)", title, re.I)
        if m_k_single:
            return float(m_k_single.group(1))
        return 0.0

    records: List[PorkSnapshotRecord] = []
    for it in list_items:
        prd_no = it["prd_no"]
        title = it["title"]
        info = it.get("info", {})
        opts = it.get("opts", [])

        raw_cat = info.get("축종/품목", "").strip()
        grade_raw = info.get("등급", "").strip()
        proc_raw = info.get("가공방법", "").strip()
        keep_raw = info.get("보관방법", "").strip() or it["keep_hint"]
        brand_raw = info.get("원산지/브랜드", "").strip()
        raw_brand = brand_raw.split("/", 1)[1].strip() if "/" in brand_raw else brand_raw

        is_excl, _ = PorkStandardizer.is_excluded_black_or_imported(f"{title} {raw_brand} {raw_cat} {grade_raw}")
        if is_excl:
            stats["dropped"] += 1
            continue

        brand = PorkStandardizer.clean_brand_name(raw_brand, title, "미트프렌즈")

        m_title_kg = re.search(r"kg[,\s당]*([\d,]+)\s*원", title, re.I)
        title_kg_price = int(m_title_kg.group(1).replace(",", "")) if m_title_kg else 0
        title_wgt_kg = _parse_title_weight_kg(title)

        parsed_opts: List[Dict[str, Any]] = []
        norm_method = "1kg단가 직수집"

        for op in opts:
            stk = int(op.get("slbStkCnt") or 0)
            prc_gb = str(op.get("prcGb") or "01").strip()
            sel_prc = int(op.get("selPrc") or 0)
            avg_wgt = float(op.get("avgWgt") or op.get("maxWgt") or op.get("minWgt") or 0.0)

            if prc_gb == "01":
                u_price = sel_prc
            else:
                if title_kg_price > 0:
                    u_price = title_kg_price
                    norm_method = "상품명 1kg단가 추출"
                elif avg_wgt > 0 and sel_prc > 0:
                    u_price = int(round(sel_prc / avg_wgt))
                    norm_method = f"박스총액({sel_prc:,}원)÷평균중량({avg_wgt:g}kg) 환산"
                elif title_wgt_kg > 0 and sel_prc > 0:
                    u_price = int(round(sel_prc / title_wgt_kg))
                    norm_method = f"박스총액({sel_prc:,}원)÷표기중량({title_wgt_kg:g}kg) 환산"
                else:
                    u_price = int(round(sel_prc / 10.0)) if sel_prc >= 15000 else sel_prc
                    norm_method = "박스총액÷표준박스중량(10kg) 환산"

            prs_day = str(op.get("prsDay") or "-").strip()
            exp_term = str(op.get("expTerm") or "-").strip()
            if prs_day.startswith("203"):
                prs_day = "당일작업"
            if exp_term.startswith("203"):
                exp_term = "상시"

            if u_price > 0:
                parsed_opts.append({"price": u_price, "proc_date": prs_day, "exp_date": exp_term, "stock": max(stk, 0)})

        active_opts = [o for o in parsed_opts if o["stock"] > 0]
        target_opts = active_opts if active_opts else parsed_opts

        if target_opts:
            stats["options"] += len(active_opts)
            target_opts.sort(key=lambda x: (x["price"], x["proc_date"]))
            unit_price_kg = target_opts[0]["price"]
            min_proc_date = target_opts[0]["proc_date"]
            max_price_kg = target_opts[-1]["price"]
            max_proc_date = target_opts[-1]["proc_date"]
            tot_w = sum(o["stock"] for o in target_opts)
            if tot_w > 0:
                avg_price_kg = int(round(sum(o["price"] * o["stock"] for o in target_opts) / tot_w))
            else:
                avg_price_kg = int(round(sum(o["price"] for o in target_opts) / len(target_opts)))

            p_dates = sorted({o["proc_date"] for o in target_opts if o["proc_date"] not in ("-", "")})
            e_dates = sorted({o["exp_date"] for o in target_opts if o["exp_date"] not in ("-", "")})
            proc_range = "상시생산" if not p_dates else (p_dates[0] if len(p_dates) == 1 else f"{p_dates[0]}~{p_dates[-1]}")
            exp_range = "-" if not e_dates else (e_dates[0] if len(e_dates) == 1 else f"{e_dates[0]}~{e_dates[-1]}")
            stock_qty = sum(o["stock"] for o in active_opts)
            opt_cnt = len(active_opts)
            is_sold_out = len(active_opts) == 0
        else:
            raw_lp = int(it["list_price"] or 0)
            if raw_lp <= 0:
                continue
            if "Box" in it["unit_txt"] or raw_lp >= 40000:
                w_kg = title_wgt_kg if title_wgt_kg > 0 else 10.0
                unit_price_kg = int(round(raw_lp / w_kg))
                norm_method = f"목록박스가({raw_lp:,}원)÷중량({w_kg:g}kg) 환산"
            else:
                unit_price_kg = raw_lp
            max_price_kg = unit_price_kg
            avg_price_kg = unit_price_kg
            min_proc_date = "-"
            max_proc_date = "-"
            proc_range = "상시생산"
            exp_range = "-"
            stock_qty = max(int(it["slb_stk"] or 0), 0)
            opt_cnt = 1 if stock_qty > 0 else 0
            is_sold_out = stock_qty <= 0

        keeping = PorkStandardizer.normalize_keeping(keep_raw, title, raw_cat)
        is_mf_female = bool(FEMALE_PORK_REGEX.search(f"{title} {grade_raw}")) or grade_raw in ("암", "암퇘지")
        stored_cat = f"{keep_raw}/{raw_cat}".strip("/")
        if is_mf_female and "[암퇘지]" not in stored_cat and "암퇘지" not in stored_cat:
            stored_cat = f"{stored_cat}[암퇘지]"
        std_part, skin_type = PorkStandardizer.map_standard_part_step1(
            title, stored_cat, source_site="미트프렌즈", product_id=prd_no
        )
        proc_type, sub_spec, map_status = PorkStandardizer.validate_step2(
            "미트프렌즈",
            title,
            stored_cat,
            proc_raw,
            brand,
            std_part,
            skin_type,
            unit_price_kg,
            product_id=prd_no,
            sex_hint="암퇘지" if is_mf_female else "",
        )

        if proc_type == "원물/박스육":
            stats["wonmul"] += 1
        else:
            stats["processed"] += 1

        prod_bucket = PorkStandardizer.classify_prod_date_bucket(min_proc_date, keeping, collected_date)

        records.append(
            PorkSnapshotRecord(
                run_id=run_id,
                collected_at=collected_at,
                collected_date=collected_date,
                item_key=f"미트프렌즈:{prd_no}",
                platform_group="미트프렌즈",
                source_site="미트프렌즈",
                is_geunbon=0,
                product_id=prd_no,
                raw_name=title,
                raw_category=stored_cat,
                brand=brand,
                keeping=keeping,
                standard_part=std_part,
                skin_type=skin_type,
                sub_spec=sub_spec,
                processing_type=proc_type,
                mapping_status=map_status,
                unit_price_kg=unit_price_kg,
                max_price_kg=max_price_kg,
                avg_price_kg=avg_price_kg,
                price_norm_method=norm_method,
                price_display=f"{unit_price_kg:,}~{max_price_kg:,}원/kg" if unit_price_kg != max_price_kg else f"{unit_price_kg:,}원/kg",
                min_proc_date=min_proc_date,
                max_proc_date=max_proc_date,
                proc_date_range=proc_range,
                exp_date_range=exp_range,
                prod_date_bucket=prod_bucket,
                is_sold_out=1 if is_sold_out else 0,
                status_text="일시품절" if is_sold_out else f"판매중(재고:{stock_qty:,}박스)",
                stock_qty=stock_qty,
                option_cnt=opt_cnt,
                product_url=f"{MEATFRIENDS_DETAIL_URL}?prdNo={prd_no}",
            )
        )

    stats["valid"] = len(records)
    logger.info("미트프렌즈 수집 완료: 유효 SKU %d건 (유효 옵션 %d건)", len(records), stats["options"])
    return records, stats


def crawl_topmeat(
    session: requests.Session, run_id: int, collected_at: str, collected_date: str
) -> Tuple[List[PorkSnapshotRecord], Dict[str, int]]:
    logger.info("=== [4/4] 탑미트(Topmeat) 한돈 전수 수집 및 생산일자 파싱 시작 ===")
    stats = {"raw": 0, "dropped": 0, "valid": 0, "wonmul": 0, "processed": 0, "boxes": 0}
    brand_code_map = {
        "topm": "탑미트",
        "bae": "배산장군포크",
        "gogi": "고기랑",
        "good": "좋은생각",
        "jeju": "대국축산(제주)",
        "mko": "문경약돌돼지",
        "new": "뉴웅비푸드",
        "elite": "엘리트포크",
        "deasan": "대산포크",
        "deasan2": "대산포크",
        "dodram": "도드람",
        "gjs": "피밀리",
    }

    sub_tasks: List[Tuple[str, str, str, int]] = []
    try:
        r = session.get(TOPMEAT_TARGET_URL, timeout=15)
        html = r.content.decode("cp949", errors="replace")
        pos_froz = html.rfind("냉동상품")
        seen_cids: Set[str] = set()
        for m in re.finditer(r"goods_kind_form_submit\('([^']+)'\)[^>]*>(.*?)</td>", html, re.S):
            cid = m.group(1).strip()
            if cid in seen_cids:
                continue
            seen_cids.add(cid)
            raw_lbl = re.sub(r"<[^>]+>", "", m.group(2)).replace("&nbsp;", " ").strip()
            m_cnt = re.search(r"\((\d+)\)", raw_lbl)
            cnt = int(m_cnt.group(1)) if m_cnt else 0
            part_lbl = re.sub(r"\(\d+\)", "", raw_lbl).strip()
            keep = "냉동" if (pos_froz != -1 and m.start() > pos_froz) else "냉장"
            is_excl, _ = PorkStandardizer.is_excluded_black_or_imported(part_lbl)
            if is_excl:
                stats["dropped"] += cnt
                continue
            if cnt > 0:
                sub_tasks.append((keep, cid, part_lbl, cnt))
    except Exception as exc:
        logger.error("탑미트 카테고리 조회 오류: %s", exc)
        return [], stats

    def _fetch_sub(task: Tuple[str, str, str, int]) -> List[Dict[str, Any]]:
        keep_lbl, cid, part_lbl, expected_cnt = task
        collected_boxes: List[Dict[str, Any]] = []
        seen_box_ids: Set[str] = set()
        page = 1
        while page <= 15:
            url = f"https://www.topmeat.co.kr/shop/list.php?ca_id={cid}&page={page}"
            try:
                r_sub = session.get(url, timeout=20)
                if r_sub.status_code != 200:
                    break
                h_sub = r_sub.content.decode("cp949", errors="replace")
                soup = BeautifulSoup(h_sub, "html.parser")
                lis = soup.select("li.lists__item")
                if not lis:
                    break
                new_in_page = 0
                for li in lis:
                    inps: Dict[str, str] = {}
                    for inp in li.find_all("input"):
                        nm = str(inp.get("name") or "")
                        base_nm = re.sub(r"\[\d*\]$", "", nm)
                        if base_nm:
                            inps[base_nm] = str(inp.get("value") or "").strip()
                    it_id = inps.get("it_id") or inps.get("it_name_2") or ""
                    lot_no = inps.get("it_lot_no") or ""
                    uniq_key = f"{it_id}_{lot_no}_{inps.get('it_origin1', '')}"
                    if uniq_key in seen_box_ids:
                        continue
                    seen_box_ids.add(uniq_key)
                    new_in_page += 1
                    inps["_keep"] = keep_lbl
                    inps["_cid"] = cid
                    inps["_part_lbl"] = part_lbl
                    collected_boxes.append(inps)
                if new_in_page == 0 or len(collected_boxes) >= expected_cnt:
                    break
                page += 1
            except Exception:
                break
        return collected_boxes

    all_boxes: List[Dict[str, Any]] = []
    with ThreadPoolExecutor(max_workers=8) as ex:
        for b_list in ex.map(_fetch_sub, sub_tasks):
            for b in b_list:
                stats["raw"] += 1
                it_name = b.get("it_name") or ""
                part_lbl = b.get("_part_lbl") or ""
                cmpny = b.get("it_cmpny") or ""
                is_excl, _ = PorkStandardizer.is_excluded_black_or_imported(f"{it_name} {part_lbl} {cmpny}")
                if is_excl:
                    stats["dropped"] += 1
                    continue
                all_boxes.append(b)

    stats["boxes"] = len(all_boxes)

    sku_groups: Dict[Tuple[str, str, str, str, str], Dict[str, Any]] = {}
    for b in all_boxes:
        cid = b["_cid"]
        part_lbl = b["_part_lbl"]
        keep_lbl = b["_keep"]
        it_name = b.get("it_name") or ""
        it_spec = b.get("it_name_1") or b.get("it_basic") or ""
        full_name = f"{it_name} {it_spec}".strip()

        b_code = (b.get("it_name_cod") or "").lower()
        b_label = brand_code_map.get(b_code, b_code.upper() if b_code else "탑미트")
        cmpny = b.get("it_cmpny") or ""
        raw_brand = f"{b_label}({cmpny})" if (cmpny and cmpny not in b_label) else (b_label or cmpny or "탑미트")

        norm_method = "1kg단가 직수집(it_maker2)"
        try:
            kg_price = int(re.sub(r"[^\d]", "", b.get("it_maker2") or "0") or 0)
        except ValueError:
            kg_price = 0
        if kg_price <= 0:
            try:
                tot_amt = int(re.sub(r"[^\d]", "", b.get("it_amount") or "0") or 0)
                wgt = float(b.get("it_origin1") or 0.0)
                if tot_amt > 0 and wgt > 0:
                    kg_price = int(round(tot_amt / wgt))
                    norm_method = f"박스총액({tot_amt:,}원)÷중량({wgt:g}kg) 환산"
            except Exception:
                kg_price = 0
        if kg_price <= 0:
            continue

        proc_dt = str(b.get("it_maker1") or "-").strip()
        if not re.match(r"^\d{4}-\d{2}-\d{2}$", proc_dt):
            proc_dt = "-"

        key = (cid, part_lbl, full_name, raw_brand, keep_lbl)
        if key not in sku_groups:
            sku_groups[key] = {
                "cid": cid,
                "part_lbl": part_lbl,
                "full_name": full_name,
                "raw_brand": raw_brand,
                "keep_lbl": keep_lbl,
                "norm_method": norm_method,
                "boxes": [],
            }
        sku_groups[key]["boxes"].append({"kg_price": kg_price, "proc_dt": proc_dt})

    records: List[PorkSnapshotRecord] = []
    for _, grp in sku_groups.items():
        boxes_list = grp["boxes"]
        boxes_list.sort(key=lambda x: (x["kg_price"], x["proc_dt"]))
        min_p = boxes_list[0]["kg_price"]
        min_dt = boxes_list[0]["proc_dt"]
        max_p = boxes_list[-1]["kg_price"]
        max_dt = boxes_list[-1]["proc_dt"]
        avg_p = int(round(sum(x["kg_price"] for x in boxes_list) / len(boxes_list)))

        dates = sorted({x["proc_dt"] for x in boxes_list if x["proc_dt"] != "-"})
        proc_range = "상시생산" if not dates else (dates[0] if len(dates) == 1 else f"{dates[0]}~{dates[-1]}")

        box_cnt = len(boxes_list)
        full_name = grp["full_name"]
        part_lbl = grp["part_lbl"]
        raw_brand = grp["raw_brand"]
        keep_lbl = grp["keep_lbl"]

        disp_name = f"[{raw_brand}] {full_name} ({part_lbl})"
        brand = PorkStandardizer.clean_brand_name(raw_brand, disp_name, "탑미트")
        keeping = PorkStandardizer.normalize_keeping(keep_lbl, disp_name, part_lbl)
        pid = f"{grp['cid']}_{re.sub(r'[^A-Za-z0-9가-힣]', '', raw_brand)[:10]}"
        std_part, skin_type = PorkStandardizer.map_standard_part_step1(
            disp_name, part_lbl, source_site="탑미트", product_id=pid
        )
        proc_type, sub_spec, map_status = PorkStandardizer.validate_step2(
            "탑미트", disp_name, part_lbl, "원물", brand, std_part, skin_type, min_p, product_id=pid
        )

        if proc_type == "원물/박스육":
            stats["wonmul"] += 1
        else:
            stats["processed"] += 1

        prod_bucket = PorkStandardizer.classify_prod_date_bucket(min_dt, keeping, collected_date)

        records.append(
            PorkSnapshotRecord(
                run_id=run_id,
                collected_at=collected_at,
                collected_date=collected_date,
                item_key=f"탑미트:{pid}",
                platform_group="탑미트",
                source_site="탑미트",
                is_geunbon=0,
                product_id=pid,
                raw_name=disp_name,
                raw_category=f"{keep_lbl}/{part_lbl}",
                brand=brand,
                keeping=keeping,
                standard_part=std_part,
                skin_type=skin_type,
                sub_spec=sub_spec,
                processing_type=proc_type,
                mapping_status=map_status,
                unit_price_kg=min_p,
                max_price_kg=max_p,
                avg_price_kg=avg_p,
                price_norm_method=grp["norm_method"],
                price_display=f"{min_p:,}~{max_p:,}원/kg" if min_p != max_p else f"{min_p:,}원/kg",
                min_proc_date=min_dt,
                max_proc_date=max_dt,
                proc_date_range=proc_range,
                exp_date_range="-",
                prod_date_bucket=prod_bucket,
                is_sold_out=0,
                status_text=f"판매중(재고:{box_cnt:,}박스)",
                stock_qty=box_cnt,
                option_cnt=box_cnt,
                product_url=f"https://www.topmeat.co.kr/shop/list.php?ca_id={grp['cid']}",
            )
        )

    stats["valid"] = len(records)
    logger.info("탑미트 수집 완료: 유효 SKU %d건 (개별 박스 %d박스)", len(records), stats["boxes"])
    return records, stats


# ---------------------------------------------------------------------------
# SQLite 시계열 누적 DB (handon_master.db) 초기화 및 관리
# ---------------------------------------------------------------------------
def get_db_connection(db_path: Path = DB_PATH) -> sqlite3.Connection:
    conn = sqlite3.connect(str(db_path), timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA synchronous=NORMAL;")
    return conn


def init_db(db_path: Path = DB_PATH) -> None:
    conn = get_db_connection(db_path)
    with conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS crawl_runs (
                run_id INTEGER PRIMARY KEY AUTOINCREMENT,
                collected_at TEXT NOT NULL,
                collected_date TEXT NOT NULL,
                trigger_source TEXT NOT NULL DEFAULT 'manual',
                total_raw_cnt INTEGER NOT NULL DEFAULT 0,
                black_excluded_cnt INTEGER NOT NULL DEFAULT 0,
                valid_sku_cnt INTEGER NOT NULL DEFAULT 0,
                wonmul_sku_cnt INTEGER NOT NULL DEFAULT 0,
                processed_sku_cnt INTEGER NOT NULL DEFAULT 0,
                geunbon_sku_cnt INTEGER NOT NULL DEFAULT 0,
                meatbox_cnt INTEGER NOT NULL DEFAULT 0,
                ekcm_cnt INTEGER NOT NULL DEFAULT 0,
                meatfriends_cnt INTEGER NOT NULL DEFAULT 0,
                topmeat_cnt INTEGER NOT NULL DEFAULT 0,
                slack_sent_cnt INTEGER NOT NULL DEFAULT 0,
                note TEXT DEFAULT ''
            );
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS product_snapshots (
                snapshot_id INTEGER PRIMARY KEY AUTOINCREMENT,
                run_id INTEGER NOT NULL,
                collected_at TEXT NOT NULL,
                collected_date TEXT NOT NULL,
                item_key TEXT NOT NULL,
                platform_group TEXT NOT NULL,
                source_site TEXT NOT NULL,
                is_geunbon INTEGER NOT NULL DEFAULT 0,
                product_id TEXT NOT NULL,
                raw_name TEXT NOT NULL,
                raw_category TEXT NOT NULL,
                brand TEXT NOT NULL,
                keeping TEXT NOT NULL,
                standard_part TEXT NOT NULL,
                skin_type TEXT NOT NULL,
                sub_spec TEXT NOT NULL,
                processing_type TEXT NOT NULL,
                mapping_status TEXT NOT NULL,
                unit_price_kg INTEGER NOT NULL,
                max_price_kg INTEGER NOT NULL,
                avg_price_kg INTEGER NOT NULL,
                price_norm_method TEXT NOT NULL,
                price_display TEXT NOT NULL,
                min_proc_date TEXT NOT NULL,
                max_proc_date TEXT NOT NULL,
                proc_date_range TEXT NOT NULL,
                exp_date_range TEXT NOT NULL,
                prod_date_bucket TEXT NOT NULL,
                is_sold_out INTEGER NOT NULL DEFAULT 0,
                status_text TEXT NOT NULL,
                stock_qty INTEGER NOT NULL DEFAULT 0,
                option_cnt INTEGER NOT NULL DEFAULT 0,
                product_url TEXT NOT NULL,
                max_product_url TEXT NOT NULL DEFAULT '',
                FOREIGN KEY (run_id) REFERENCES crawl_runs(run_id)
            );
            """
        )
        existing_cols = {r[1] for r in conn.execute("PRAGMA table_info(product_snapshots);").fetchall()}
        if "max_product_url" not in existing_cols:
            conn.execute("ALTER TABLE product_snapshots ADD COLUMN max_product_url TEXT NOT NULL DEFAULT '';")
            conn.execute("UPDATE product_snapshots SET max_product_url = product_url WHERE max_product_url = '';")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_snapshots_run_id ON product_snapshots(run_id);")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_snapshots_item_key ON product_snapshots(item_key);")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_snapshots_part_keep ON product_snapshots(standard_part, keeping);")
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS tracked_items (
                item_key TEXT PRIMARY KEY,
                source_site TEXT NOT NULL,
                platform_group TEXT NOT NULL,
                product_id TEXT NOT NULL,
                raw_name TEXT NOT NULL,
                brand TEXT NOT NULL,
                standard_part TEXT NOT NULL,
                keeping TEXT NOT NULL,
                product_url TEXT NOT NULL,
                memo TEXT DEFAULT '',
                is_active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS carcass_market_prices (
                price_date TEXT PRIMARY KEY,
                year INTEGER NOT NULL,
                month TEXT NOT NULL,
                iso_week TEXT NOT NULL,
                price_ex_jeju INTEGER NOT NULL DEFAULT 0,
                price_national INTEGER NOT NULL DEFAULT 0,
                grade_1plus INTEGER NOT NULL DEFAULT 0,
                grade_1 INTEGER NOT NULL DEFAULT 0,
                grade_2 INTEGER NOT NULL DEFAULT 0,
                farm_price_110kg INTEGER NOT NULL DEFAULT 0,
                retail_samgyeop INTEGER NOT NULL DEFAULT 0,
                auction_head_cnt INTEGER NOT NULL DEFAULT 0,
                is_preliminary INTEGER NOT NULL DEFAULT 0,
                updated_at TEXT NOT NULL
            );
            """
        )
        conn.execute("CREATE INDEX IF NOT EXISTS idx_carcass_year_month ON carcass_market_prices(year, month);")

        # 스키마 매핑 버전(v3):
        # 1) 4대 플랫폼 전수 카테고리/품목코드 사전 매핑 (미박앞다리살=미박전지, 미박뒷다리살=미박후지 완벽 통합)
        # 2) 암퇘지 상품 별도 구분(sex_type='암퇘지' vs '일반') 및 지방두께(3mm/10mm)·박스내 팩수 원물 오탐 보정
        db_ver = conn.execute("PRAGMA user_version;").fetchone()[0]
        legacy_cnt = conn.execute(
            """
            SELECT COUNT(*) FROM product_snapshots
            WHERE standard_part IN ('특수부위/부산물', '앞다리살', '뒷다리살')
               OR (standard_part = '삼겹살' AND skin_type = '미박')
            """
        ).fetchone()[0]
        if db_ver < 3 or legacy_cnt > 0:
            rows_to_fix = conn.execute(
                """
                SELECT snapshot_id, source_site, product_id, raw_name, raw_category,
                       brand, skin_type, sub_spec, processing_type, unit_price_kg
                FROM product_snapshots
                """
            ).fetchall()
            updates: List[Tuple[str, str, str, str, str, str, int]] = []
            for r in rows_to_fix:
                site = r["source_site"]
                proc_hint = "세절" if (site == "미트박스" and r["processing_type"] == "세절/가공육") else ""
                was_female = r["sub_spec"] == "암퇘지"
                raw_cat_val = r["raw_category"]
                if was_female and site == "미트프렌즈" and "암" not in f"{r['raw_name']} {raw_cat_val}":
                    raw_cat_val = f"{raw_cat_val}[암퇘지]"
                new_part, new_skin = PorkStandardizer.map_standard_part_step1(
                    r["raw_name"],
                    raw_cat_val,
                    skin_hint=r["skin_type"],
                    source_site=site,
                    product_id=r["product_id"],
                )
                new_proc, new_sub, new_status = PorkStandardizer.validate_step2(
                    site,
                    r["raw_name"],
                    raw_cat_val,
                    proc_hint,
                    r["brand"],
                    new_part,
                    new_skin,
                    int(r["unit_price_kg"] or 0),
                    product_id=r["product_id"],
                    sex_hint="암퇘지" if was_female else "",
                )
                updates.append(
                    (raw_cat_val, new_part, new_skin, new_proc, new_sub, new_status, int(r["snapshot_id"]))
                )
            if updates:
                conn.executemany(
                    """
                    UPDATE product_snapshots
                    SET raw_category = ?, standard_part = ?, skin_type = ?, processing_type = ?, sub_spec = ?, mapping_status = ?
                    WHERE snapshot_id = ?
                    """,
                    updates,
                )
                # 각 회차(crawl_runs)의 원물/세절 집계도 최신 보정 결과와 동기화
                for run_r in conn.execute("SELECT run_id FROM crawl_runs").fetchall():
                    rid_val = int(run_r["run_id"])
                    w_cnt = conn.execute(
                        "SELECT COUNT(*) FROM product_snapshots WHERE run_id = ? AND processing_type = '원물/박스육'",
                        (rid_val,),
                    ).fetchone()[0]
                    p_cnt = conn.execute(
                        "SELECT COUNT(*) FROM product_snapshots WHERE run_id = ? AND processing_type != '원물/박스육'",
                        (rid_val,),
                    ).fetchone()[0]
                    conn.execute(
                        "UPDATE crawl_runs SET wonmul_sku_cnt = ?, processed_sku_cnt = ? WHERE run_id = ?",
                        (w_cnt, p_cnt, rid_val),
                    )
            # tracked_items 테이블의 standard_part도 최신 매핑으로 동기화
            t_rows = conn.execute("SELECT item_key, source_site, product_id, raw_name FROM tracked_items").fetchall()
            for tr in t_rows:
                snap_m = conn.execute(
                    "SELECT standard_part FROM product_snapshots WHERE item_key = ? ORDER BY run_id DESC LIMIT 1",
                    (tr["item_key"],),
                ).fetchone()
                if snap_m:
                    conn.execute(
                        "UPDATE tracked_items SET standard_part = ? WHERE item_key = ?",
                        (snap_m["standard_part"], tr["item_key"]),
                    )
            conn.execute("PRAGMA user_version = 3;")
            logger.info("기존 DB 스냅샷 %d건에 대해 v3(암퇘지 별도분리 & 미박앞다리살=미박전지/미박뒷다리살=미박후지 전수매핑) 마이그레이션 완료", len(updates))
    conn.close()


def toggle_tracked_item(
    item_key: str,
    is_active: Optional[bool] = None,
    memo: Optional[str] = None,
    db_path: Path = DB_PATH,
) -> Dict[str, Any]:
    """
    상품 체크박스(중점관리 체크) 상태를 DB(tracked_items)에 즉시 반영.
    새로고침이나 다음 크롤링 이후에도 영구 유지됨.
    """
    init_db(db_path)
    conn = get_db_connection(db_path)
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with conn:
        existing = conn.execute("SELECT * FROM tracked_items WHERE item_key = ?", (item_key,)).fetchone()
        if existing:
            new_active = (0 if existing["is_active"] else 1) if is_active is None else (1 if is_active else 0)
            new_memo = existing["memo"] if memo is None else memo
            conn.execute(
                "UPDATE tracked_items SET is_active = ?, memo = ?, updated_at = ? WHERE item_key = ?",
                (new_active, new_memo, now_str, item_key),
            )
            snap = dict(existing)
            snap["is_active"] = new_active
            snap["memo"] = new_memo
            snap["updated_at"] = now_str
        else:
            snap_row = conn.execute(
                "SELECT * FROM product_snapshots WHERE item_key = ? ORDER BY snapshot_id DESC LIMIT 1",
                (item_key,),
            ).fetchone()
            if not snap_row:
                conn.close()
                raise ValueError(f"상품 키({item_key})를 스냅샷 DB에서 찾을 수 없습니다.")
            new_active = 1 if is_active is None else (1 if is_active else 0)
            new_memo = memo or "중점관리 모니터링 대상"
            conn.execute(
                """
                INSERT INTO tracked_items (
                    item_key, source_site, platform_group, product_id, raw_name,
                    brand, standard_part, keeping, product_url, memo, is_active, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    item_key,
                    snap_row["source_site"],
                    snap_row["platform_group"],
                    snap_row["product_id"],
                    snap_row["raw_name"],
                    snap_row["brand"],
                    snap_row["standard_part"],
                    snap_row["keeping"],
                    snap_row["product_url"],
                    new_memo,
                    new_active,
                    now_str,
                    now_str,
                ),
            )
            snap = {
                "item_key": item_key,
                "source_site": snap_row["source_site"],
                "platform_group": snap_row["platform_group"],
                "product_id": snap_row["product_id"],
                "raw_name": snap_row["raw_name"],
                "brand": snap_row["brand"],
                "standard_part": snap_row["standard_part"],
                "keeping": snap_row["keeping"],
                "product_url": snap_row["product_url"],
                "memo": new_memo,
                "is_active": new_active,
                "created_at": now_str,
                "updated_at": now_str,
            }
    conn.close()
    return snap


# ---------------------------------------------------------------------------
# 데이터 집계 및 문제해결(신호등 알림, 5자 비교, 생산일자 분석) 엔진
# ---------------------------------------------------------------------------
def get_low_stock_threshold(keeping: str) -> int:
    """품절/주의 재고 기준: 냉장 5개 미만, 냉동 10개 미만"""
    return 5 if keeping == "냉장" else 10


def is_item_low_stock(is_sold_out: int, stock_qty: int, keeping: str) -> bool:
    """품절이 아닌 판매중 상품 중 잔여 재고가 기준(냉장<5, 냉동<10) 미만인 주의 품목 여부 (상시재고 0 표기는 제외)"""
    if is_sold_out:
        return False
    qty = int(stock_qty or 0)
    return 0 < qty < get_low_stock_threshold(keeping)


def build_comparison_matrix_from_rows(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    27개 표준부위 × 2개 보관방법(냉장/냉동) × 2개 돈종구분(일반한돈 / 암퇘지 전용)에 대해
    5개 비교군(미트박스-근본한돈, 미트박스-일반한돈, 금천미트, 미트프렌즈, 탑미트)의
    최저가, 최고가, 평균가, 재고수량, 최저·최고가 브랜드/생산일자/하이퍼링크 URL 등을 완벽 분리 산출.
    """
    matrix: List[Dict[str, Any]] = []

    for part in STANDARD_PARTS_ORDER:
        for keep in KEEPINGS_ORDER:
            low_limit = get_low_stock_threshold(keep)
            for sex_type in SEX_TYPES_ORDER:
                sub = [
                    r
                    for r in rows
                    if r["standard_part"] == part
                    and r["keeping"] == keep
                    and PorkStandardizer.get_row_sex_type(r) == sex_type
                ]
                if not sub:
                    continue

                display_part = part if sex_type == "일반" else f"{part} [암퇘지]"

                row_data: Dict[str, Any] = {
                    "standard_part": part,
                    "keeping": keep,
                    "sex_type": sex_type,
                    "display_part": display_part,
                    "low_stock_limit": low_limit,
                    "groups": {},
                    "total_sku_cnt": len(sub),
                    "total_stock_qty": sum(int(r["stock_qty"] or 0) for r in sub if not r["is_sold_out"]),
                }

                group_min_candidates: List[Tuple[int, str, Dict[str, Any]]] = []
                group_max_candidates: List[Tuple[int, str, Dict[str, Any]]] = []

                for grp in PLATFORM_GROUPS_ORDER:
                    g_all = [r for r in sub if r["platform_group"] == grp]
                    g_avail = [r for r in g_all if not r["is_sold_out"] and r["unit_price_kg"] > 0]

                    if sex_type == "암퇘지":
                        g_std_wonmul = [r for r in g_avail if r["processing_type"] == "원물/박스육"]
                        g_non_modon_wonmul = g_std_wonmul
                    else:
                        g_std_wonmul = [
                            r for r in g_avail if r["processing_type"] == "원물/박스육" and r["sub_spec"] == "일반"
                        ]
                        g_non_modon_wonmul = [
                            r for r in g_avail if r["processing_type"] == "원물/박스육" and r["sub_spec"] != "모돈"
                        ]
                    g_any_wonmul = [r for r in g_avail if r["processing_type"] == "원물/박스육"]

                    if g_std_wonmul:
                        comp_pool = g_std_wonmul
                        basis = "원물(암퇘지)" if sex_type == "암퇘지" else "원물(표준)"
                    elif g_non_modon_wonmul:
                        comp_pool = g_non_modon_wonmul
                        basis = "원물(특수정선)"
                    elif g_any_wonmul:
                        comp_pool = g_any_wonmul
                        basis = "원물(모돈포함)"
                    else:
                        comp_pool = g_avail
                        basis = "세절/가공포함" if g_avail else "-"

                    sold_out_cnt = sum(1 for r in g_all if r["is_sold_out"])
                    avail_cnt = len(g_avail)
                    stock_boxes = sum(int(r["stock_qty"] or 0) for r in g_avail)
                    is_grp_low_stock = 1 if (avail_cnt > 0 and 0 < stock_boxes < low_limit) else 0

                    if comp_pool:
                        comp_pool_sorted = sorted(comp_pool, key=lambda x: (x["unit_price_kg"], x["min_proc_date"]))
                        best_item = comp_pool_sorted[0]
                        worst_item = sorted(
                            comp_pool, key=lambda x: (x["max_price_kg"], x["max_proc_date"]), reverse=True
                        )[0]
                        min_p = int(best_item["unit_price_kg"])
                        max_p = int(worst_item["max_price_kg"])
                        avg_p = int(round(sum(x["unit_price_kg"] for x in comp_pool) / len(comp_pool)))

                        g_info = {
                            "min_price": min_p,
                            "max_price": max_p,
                            "avg_price": avg_p,
                            "sku_cnt": len(g_all),
                            "avail_cnt": avail_cnt,
                            "sold_out_cnt": sold_out_cnt,
                            "stock_qty": stock_boxes,
                            "is_low_stock": is_grp_low_stock,
                            "basis": basis,
                            "best_brand": best_item["brand"],
                            "best_name": best_item["raw_name"],
                            "best_proc_date": best_item["min_proc_date"],
                            "best_prod_bucket": best_item["prod_date_bucket"],
                            "best_url": best_item["product_url"],
                            "best_item_key": best_item["item_key"],
                            "worst_brand": worst_item["brand"],
                            "worst_name": worst_item["raw_name"],
                            "worst_proc_date": worst_item["max_proc_date"],
                            "worst_prod_bucket": worst_item["prod_date_bucket"],
                            "worst_url": worst_item.get("max_product_url") or worst_item["product_url"],
                            "worst_item_key": worst_item["item_key"],
                        }
                        group_min_candidates.append((min_p, grp, g_info))
                        group_max_candidates.append((max_p, grp, g_info))
                    else:
                        g_info = {
                            "min_price": None,
                            "max_price": None,
                            "avg_price": None,
                            "sku_cnt": len(g_all),
                            "avail_cnt": 0,
                            "sold_out_cnt": sold_out_cnt,
                            "stock_qty": 0,
                            "is_low_stock": 0,
                            "basis": "품절/미취급" if g_all else "미취급",
                            "best_brand": "-",
                            "best_name": "-",
                            "best_proc_date": "-",
                            "best_prod_bucket": "-",
                            "best_url": "",
                            "best_item_key": "",
                            "worst_brand": "-",
                            "worst_name": "-",
                            "worst_proc_date": "-",
                            "worst_prod_bucket": "-",
                            "worst_url": "",
                            "worst_item_key": "",
                        }

                    row_data["groups"][grp] = g_info

                if group_min_candidates:
                    group_min_candidates.sort(key=lambda x: x[0])
                    overall_min_p, overall_winner_grp, overall_winner_info = group_min_candidates[0]
                    second_min_p = group_min_candidates[1][0] if len(group_min_candidates) > 1 else overall_min_p
                    price_gap_1_2 = second_min_p - overall_min_p
                else:
                    overall_min_p = None
                    overall_winner_grp = "-"
                    overall_winner_info = {}
                    price_gap_1_2 = 0

                if group_max_candidates:
                    group_max_candidates.sort(key=lambda x: x[0], reverse=True)
                    overall_max_p, overall_highest_grp, overall_highest_info = group_max_candidates[0]
                else:
                    overall_max_p = None
                    overall_highest_grp = "-"
                    overall_highest_info = {}

                gb_info = row_data["groups"]["미트박스 - 근본한돈"]
                gb_min_p = gb_info["min_price"]
                gb_is_low_stock = bool(gb_info.get("is_low_stock"))
                gb_is_sold_out = gb_info["sku_cnt"] > 0 and gb_info["avail_cnt"] == 0

                # 타사/일반 최저가 (근본한돈 제외한 나머지 4개 그룹 중 최저가)
                other_candidates = [c for c in group_min_candidates if c[1] != "미트박스 - 근본한돈"]
                competitor_min_p = other_candidates[0][0] if other_candidates else None
                competitor_winner_grp = other_candidates[0][1] if other_candidates else "-"
                competitor_winner_info = other_candidates[0][2] if other_candidates else {}

                if gb_min_p is not None and overall_min_p is not None:
                    gb_vs_lowest_diff = gb_min_p - overall_min_p
                    gb_vs_comp_diff = (gb_min_p - competitor_min_p) if competitor_min_p is not None else 0
                    if gb_min_p == overall_min_p:
                        if competitor_min_p is not None and competitor_min_p > gb_min_p:
                            gb_status = f"최저가 단독우위 (-{competitor_min_p - gb_min_p:,}원)"
                        else:
                            gb_status = "최저가 공동1위"
                        gb_signal = "safe"
                        if gb_is_low_stock:
                            gb_status += f" · 재고주의({gb_info['stock_qty']}박스)"
                    else:
                        gb_status = f"타사대비 +{gb_vs_lowest_diff:,}원 높음 ({overall_winner_grp} 우위)"
                        gb_signal = "danger" if gb_vs_lowest_diff >= 300 else "warning"
                        if gb_is_low_stock:
                            gb_status += f" · 재고주의({gb_info['stock_qty']}박스)"
                elif gb_is_sold_out:
                    gb_vs_lowest_diff = None
                    gb_vs_comp_diff = None
                    gb_status = "근본한돈 일시품절 (재고보충 시급)"
                    gb_signal = "warning"
                else:
                    gb_vs_lowest_diff = None
                    gb_vs_comp_diff = None
                    gb_status = "근본한돈 미취급 품목"
                    gb_signal = "neutral"

                has_any_caution = int(
                    gb_is_sold_out
                    or gb_is_low_stock
                    or any(
                        (g["sku_cnt"] > 0 and g["avail_cnt"] == 0) or g.get("is_low_stock")
                        for g in row_data["groups"].values()
                    )
                )

                row_data.update(
                    {
                        "overall_min_price": overall_min_p,
                        "overall_winner_group": overall_winner_grp,
                        "overall_winner_brand": overall_winner_info.get("best_brand", "-"),
                        "overall_winner_proc_date": overall_winner_info.get("best_proc_date", "-"),
                        "overall_winner_url": overall_winner_info.get("best_url", ""),
                        "overall_max_price": overall_max_p,
                        "overall_highest_group": overall_highest_grp,
                        "overall_highest_brand": overall_highest_info.get("worst_brand", "-"),
                        "overall_highest_proc_date": overall_highest_info.get("worst_proc_date", "-"),
                        "overall_highest_url": overall_highest_info.get("worst_url", ""),
                        "price_gap_1_2": price_gap_1_2,
                        "competitor_min_price": competitor_min_p,
                        "competitor_winner_group": competitor_winner_grp,
                        "competitor_winner_brand": competitor_winner_info.get("best_brand", "-"),
                        "competitor_winner_proc_date": competitor_winner_info.get("best_proc_date", "-"),
                        "competitor_winner_url": competitor_winner_info.get("best_url", ""),
                        "gb_min_price": gb_min_p,
                        "gb_max_price": gb_info["max_price"],
                        "gb_is_low_stock": 1 if gb_is_low_stock else 0,
                        "gb_is_caution": 1 if (gb_is_sold_out or gb_is_low_stock) else 0,
                        "has_any_caution": has_any_caution,
                        "gb_vs_lowest_diff": gb_vs_lowest_diff,
                        "gb_vs_comp_diff": gb_vs_comp_diff,
                        "gb_status": gb_status,
                        "gb_signal": gb_signal,
                    }
                )
                matrix.append(row_data)

    return matrix


def build_production_date_analysis(rows: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    생산일자(제조일/가공일) 구간별 최저가·최고가 분석 (근본한돈 메인 기준 대비 타플랫폼 데이터 구성):
    - 근본한돈(메인)의 최저가·최고가·생산일자·재고·하이퍼링크를 중심축으로 배치하고,
    - 타플랫폼(미트박스일반·금천미트·미트프렌즈·탑미트)의 생산일자 구간별 최저·최고가 및 하이퍼링크를 1:1 대조.
    """
    buckets_order = [
        "당월 신상품(최근 14일)",
        "전월/일반 생산(15~45일)",
        "장기보관/기한임박(45일 경과)",
        "주문당일생산/상시",
    ]
    active_wonmul = [
        r for r in rows if not r["is_sold_out"] and r["processing_type"] == "원물/박스육" and r["unit_price_kg"] > 0
    ]

    def _summarize_bucket_items(b_items: List[Dict[str, Any]], b_name: str) -> Dict[str, Any]:
        if not b_items:
            return {
                "bucket": b_name,
                "sku_cnt": 0,
                "stock_qty": 0,
                "min_price": 0,
                "min_url": "",
                "min_platform": "-",
                "min_brand": "-",
                "min_proc_date": "-",
                "max_price": 0,
                "max_url": "",
                "max_platform": "-",
                "max_brand": "-",
                "max_proc_date": "-",
                "avg_price": 0,
                "gb_min_price": None,
                "gb_min_url": "",
            }
        best_it = sorted(b_items, key=lambda x: (x["unit_price_kg"], x["min_proc_date"]))[0]
        worst_it = sorted(b_items, key=lambda x: (x["max_price_kg"], x["max_proc_date"]), reverse=True)[0]
        gb_b = [x for x in b_items if x["platform_group"] == "미트박스 - 근본한돈"]
        gb_best = sorted(gb_b, key=lambda x: x["unit_price_kg"])[0] if gb_b else None
        return {
            "bucket": b_name,
            "sku_cnt": len(b_items),
            "stock_qty": sum(int(r["stock_qty"] or 0) for r in b_items),
            "min_price": best_it["unit_price_kg"],
            "min_url": best_it["product_url"],
            "min_platform": best_it["platform_group"],
            "min_brand": best_it["brand"],
            "min_proc_date": best_it["min_proc_date"],
            "max_price": worst_it["max_price_kg"],
            "max_url": worst_it.get("max_product_url") or worst_it["product_url"],
            "max_platform": worst_it["platform_group"],
            "max_brand": worst_it["brand"],
            "max_proc_date": worst_it["max_proc_date"],
            "avg_price": int(round(sum(r["unit_price_kg"] for r in b_items) / len(b_items))),
            "gb_min_price": gb_best["unit_price_kg"] if gb_best else None,
            "gb_min_url": gb_best["product_url"] if gb_best else "",
        }

    # 1. 구간별 전체 요약 (프론트엔드에서는 부위 탭/보관 필터에 따라 동적 재계산도 함께 수행)
    bucket_summary: List[Dict[str, Any]] = [
        _summarize_bucket_items([r for r in active_wonmul if r["prod_date_bucket"] == b_name], b_name)
        for b_name in buckets_order
    ]

    # 2. 표준부위 × 보관 × 돈종(일반/암퇘지)별 [근본한돈 메인 vs 타플랫폼 생산일자 구간] 매트릭스
    part_date_matrix: List[Dict[str, Any]] = []
    dumping_alerts: List[Dict[str, Any]] = []

    for part in STANDARD_PARTS_ORDER:
        for keep in KEEPINGS_ORDER:
            for sex_type in SEX_TYPES_ORDER:
                all_for_key = [
                    r
                    for r in rows
                    if r["standard_part"] == part
                    and r["keeping"] == keep
                    and PorkStandardizer.get_row_sex_type(r) == sex_type
                ]
                sex_pool = [
                    r
                    for r in active_wonmul
                    if r["standard_part"] == part
                    and r["keeping"] == keep
                    and PorkStandardizer.get_row_sex_type(r) == sex_type
                ]
                if not sex_pool:
                    continue

                if sex_type == "일반":
                    sub = [r for r in sex_pool if r["sub_spec"] == "일반"]
                    if not sub:
                        sub = [r for r in sex_pool if r["sub_spec"] != "모돈"] or sex_pool
                else:
                    sub = sex_pool

                display_part = part if sex_type == "일반" else f"{part} [암퇘지]"

                # (A) 근본한돈 메인 데이터 추출
                gb_all_raw = [r for r in all_for_key if r["platform_group"] == "미트박스 - 근본한돈"]
                gb_active = [r for r in sub if r["platform_group"] == "미트박스 - 근본한돈"]
                if not gb_active:
                    gb_active = [r for r in sex_pool if r["platform_group"] == "미트박스 - 근본한돈"]

                if gb_active:
                    gb_min_rec = sorted(gb_active, key=lambda x: (x["unit_price_kg"], x["min_proc_date"]))[0]
                    gb_max_rec = sorted(gb_active, key=lambda x: (x["max_price_kg"], x["max_proc_date"]), reverse=True)[0]
                    gb_main = {
                        "status": "판매중",
                        "sku_cnt": len(gb_active),
                        "stock_qty": sum(int(x["stock_qty"] or 0) for x in gb_active),
                        "min_price": gb_min_rec["unit_price_kg"],
                        "min_proc_date": gb_min_rec["min_proc_date"],
                        "min_prod_bucket": gb_min_rec["prod_date_bucket"],
                        "min_url": gb_min_rec["product_url"],
                        "min_name": gb_min_rec["raw_name"],
                        "max_price": gb_max_rec["max_price_kg"],
                        "max_proc_date": gb_max_rec["max_proc_date"],
                        "max_url": gb_max_rec["product_url"],
                        "max_name": gb_max_rec["raw_name"],
                    }
                else:
                    gb_main = {
                        "status": "일시품절" if gb_all_raw else "미취급",
                        "sku_cnt": len(gb_all_raw),
                        "stock_qty": 0,
                        "min_price": None,
                        "min_proc_date": "-",
                        "min_prod_bucket": "-",
                        "min_url": gb_all_raw[0]["product_url"] if gb_all_raw else "",
                        "min_name": "-",
                        "max_price": None,
                        "max_proc_date": "-",
                        "max_url": "",
                        "max_name": "-",
                    }

                # (B) 타플랫폼(경쟁 4군) 데이터 추출 및 생산일자 구간별 최저·최고가 산출
                comp_active = [r for r in sub if r["platform_group"] != "미트박스 - 근본한돈"]

                def _build_bucket_dict(pool_items: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
                    res_map: Dict[str, Dict[str, Any]] = {}
                    for b_name in buckets_order:
                        bi = [r for r in pool_items if r["prod_date_bucket"] == b_name]
                        if bi:
                            best = sorted(bi, key=lambda x: (x["unit_price_kg"], x["min_proc_date"]))[0]
                            worst = sorted(bi, key=lambda x: (x["max_price_kg"], x["max_proc_date"]), reverse=True)[0]
                            diff_vs_gb = (
                                (gb_main["min_price"] - best["unit_price_kg"])
                                if gb_main["min_price"] is not None
                                else None
                            )
                            res_map[b_name] = {
                                "sku_cnt": len(bi),
                                "stock_qty": sum(int(x["stock_qty"] or 0) for x in bi),
                                "min_price": best["unit_price_kg"],
                                "max_price": worst["max_price_kg"],
                                "avg_price": int(round(sum(x["unit_price_kg"] for x in bi) / len(bi))),
                                "best_platform": best["platform_group"],
                                "best_brand": best["brand"],
                                "best_proc_date": best["min_proc_date"],
                                "best_url": best["product_url"],
                                "best_name": best["raw_name"],
                                "worst_platform": worst["platform_group"],
                                "worst_brand": worst["brand"],
                                "worst_proc_date": worst["max_proc_date"],
                                "worst_url": worst.get("max_product_url") or worst["product_url"],
                                "worst_name": worst["raw_name"],
                                "gb_diff": diff_vs_gb,
                            }
                        else:
                            res_map[b_name] = {
                                "sku_cnt": 0,
                                "stock_qty": 0,
                                "min_price": None,
                                "max_price": None,
                                "avg_price": None,
                                "best_platform": "-",
                                "best_brand": "-",
                                "best_proc_date": "-",
                                "best_url": "",
                                "best_name": "-",
                                "worst_platform": "-",
                                "worst_brand": "-",
                                "worst_proc_date": "-",
                                "worst_url": "",
                                "worst_name": "-",
                                "gb_diff": None,
                            }
                    return res_map

                b_map = _build_bucket_dict(sub)
                comp_b_map = _build_bucket_dict(comp_active)

                fresh_info = b_map["당월 신상품(최근 14일)"]
                normal_info = b_map["전월/일반 생산(15~45일)"]
                aged_info = b_map["장기보관/기한임박(45일 경과)"]

                fresh_ref_price = fresh_info["min_price"] or normal_info["min_price"]
                aged_min_price = aged_info["min_price"]

                illusion_note = "정상 (신선재고 중심 형성)"
                is_dumping_illusion = False
                dumping_gap = 0

                if aged_min_price is not None and fresh_ref_price is not None:
                    dumping_gap = fresh_ref_price - aged_min_price
                    if dumping_gap > 0:
                        is_dumping_illusion = True
                        illusion_note = (
                            f"노후재고 덤핑 주의! 장기보관({aged_info['best_proc_date']}) 최저가가 "
                            f"신선물량 대비 -{dumping_gap:,}원/kg 낮아 최저가 착시 발생"
                        )
                        dumping_alerts.append(
                            {
                                "standard_part": part,
                                "display_part": display_part,
                                "keeping": keep,
                                "sex_type": sex_type,
                                "gb_min_price": gb_main["min_price"],
                                "gb_min_proc_date": gb_main["min_proc_date"],
                                "gb_min_url": gb_main["min_url"],
                                "aged_platform": aged_info["best_platform"],
                                "aged_brand": aged_info["best_brand"],
                                "aged_proc_date": aged_info["best_proc_date"],
                                "aged_min_price": aged_min_price,
                                "aged_url": aged_info["best_url"],
                                "fresh_min_price": fresh_ref_price,
                                "fresh_platform": fresh_info["best_platform"]
                                if fresh_info["min_price"]
                                else normal_info["best_platform"],
                                "fresh_proc_date": fresh_info["best_proc_date"]
                                if fresh_info["min_price"]
                                else normal_info["best_proc_date"],
                                "fresh_url": fresh_info["best_url"]
                                if fresh_info["min_price"]
                                else normal_info["best_url"],
                                "dumping_gap": dumping_gap,
                                "message": illusion_note,
                            }
                        )

                all_sorted_min = sorted(sub, key=lambda x: x["unit_price_kg"])
                all_sorted_max = sorted(sub, key=lambda x: x["max_price_kg"], reverse=True)
                lowest_rec = all_sorted_min[0]
                highest_rec = all_sorted_max[0]

                if comp_active:
                    comp_min_rec = sorted(comp_active, key=lambda x: (x["unit_price_kg"], x["min_proc_date"]))[0]
                    comp_max_rec = sorted(
                        comp_active, key=lambda x: (x["max_price_kg"], x["max_proc_date"]), reverse=True
                    )[0]
                    comp_summary = {
                        "sku_cnt": len(comp_active),
                        "lowest_price": comp_min_rec["unit_price_kg"],
                        "lowest_proc_date": comp_min_rec["min_proc_date"],
                        "lowest_prod_bucket": comp_min_rec["prod_date_bucket"],
                        "lowest_platform": comp_min_rec["platform_group"],
                        "lowest_brand": comp_min_rec["brand"],
                        "lowest_url": comp_min_rec["product_url"],
                        "highest_price": comp_max_rec["max_price_kg"],
                        "highest_proc_date": comp_max_rec["max_proc_date"],
                        "highest_platform": comp_max_rec["platform_group"],
                        "highest_brand": comp_max_rec["brand"],
                        "highest_url": comp_max_rec.get("max_product_url") or comp_max_rec["product_url"],
                    }
                else:
                    comp_summary = {
                        "sku_cnt": 0,
                        "lowest_price": None,
                        "lowest_proc_date": "-",
                        "lowest_prod_bucket": "-",
                        "lowest_platform": "-",
                        "lowest_brand": "-",
                        "lowest_url": "",
                        "highest_price": None,
                        "highest_proc_date": "-",
                        "highest_platform": "-",
                        "highest_brand": "-",
                        "highest_url": "",
                    }

                # 근본한돈 메인 대비 타플랫폼 생산일자 종합 진단
                gb_p = gb_main["min_price"]
                cp_p = comp_summary["lowest_price"]
                if gb_p is not None and cp_p is not None:
                    gap_gb_cp = gb_p - cp_p
                    if gap_gb_cp <= 0:
                        gb_vs_comp_signal = "safe"
                        gb_vs_comp_note = (
                            f"근본한돈 우위 (-{abs(gap_gb_cp):,}원/kg | 생산일:{gb_main['min_proc_date']} vs "
                            f"타사최저 {comp_summary['lowest_platform']} {cp_p:,}원)"
                        )
                    elif "장기보관" in str(comp_summary["lowest_prod_bucket"]):
                        gb_vs_comp_signal = "warning"
                        gb_vs_comp_note = (
                            f"타사 최저가({cp_p:,}원, {comp_summary['lowest_proc_date']})는 장기보관 노후재고 — "
                            f"근본한돈({gb_p:,}원) 정상 단가 방어 권장"
                        )
                    else:
                        gb_vs_comp_signal = "danger"
                        gb_vs_comp_note = (
                            f"근본한돈 +{gap_gb_cp:,}원 높음 (타사 1위: {comp_summary['lowest_platform']} "
                            f"{cp_p:,}원 · 생산일 {comp_summary['lowest_proc_date']})"
                        )
                elif gb_main["status"] == "일시품절":
                    gb_vs_comp_signal = "warning"
                    gb_vs_comp_note = (
                        f"근본한돈 일시품절 — 타사 최저 {cp_p:,}원({comp_summary['lowest_platform']})"
                        if cp_p is not None
                        else "근본한돈 일시품절"
                    )
                else:
                    gb_vs_comp_signal = "neutral"
                    gb_vs_comp_note = (
                        f"근본한돈 미취급 — 타플랫폼 최저 {cp_p:,}원 ~ 최고 {comp_summary['highest_price']:,}원"
                        if cp_p is not None and comp_summary["highest_price"] is not None
                        else illusion_note
                    )

                part_date_matrix.append(
                    {
                        "standard_part": part,
                        "display_part": display_part,
                        "keeping": keep,
                        "sex_type": sex_type,
                        "total_sku_cnt": len(sub),
                        "gb_main": gb_main,
                        "comp_summary": comp_summary,
                        "comp_buckets": comp_b_map,
                        "gb_vs_comp_signal": gb_vs_comp_signal,
                        "gb_vs_comp_note": gb_vs_comp_note,
                        "lowest_price": lowest_rec["unit_price_kg"],
                        "lowest_proc_date": lowest_rec["min_proc_date"],
                        "lowest_platform": lowest_rec["platform_group"],
                        "lowest_brand": lowest_rec["brand"],
                        "lowest_url": lowest_rec["product_url"],
                        "highest_price": highest_rec["max_price_kg"],
                        "highest_proc_date": highest_rec["max_proc_date"],
                        "highest_platform": highest_rec["platform_group"],
                        "highest_brand": highest_rec["brand"],
                        "highest_url": highest_rec.get("max_product_url") or highest_rec["product_url"],
                        "price_spread": highest_rec["max_price_kg"] - lowest_rec["unit_price_kg"],
                        "buckets": b_map,
                        "is_dumping_illusion": is_dumping_illusion,
                        "dumping_gap": dumping_gap,
                        "illusion_note": illusion_note,
                    }
                )

    # 3. 개별 상품 내 생산일자별 가격 편차(Max - Min) 상위 리스트
    sku_spread_list = sorted(
        [
            {
                "item_key": r["item_key"],
                "platform_group": r["platform_group"],
                "brand": r["brand"],
                "standard_part": r["standard_part"],
                "keeping": r["keeping"],
                "sex_type": PorkStandardizer.get_row_sex_type(r),
                "raw_name": r["raw_name"],
                "min_price_kg": r["unit_price_kg"],
                "min_proc_date": r["min_proc_date"],
                "max_price_kg": r["max_price_kg"],
                "max_proc_date": r["max_proc_date"],
                "avg_price_kg": r["avg_price_kg"],
                "price_spread": r["max_price_kg"] - r["unit_price_kg"],
                "proc_date_range": r["proc_date_range"],
                "prod_date_bucket": r["prod_date_bucket"],
                "stock_qty": r["stock_qty"],
                "option_cnt": r["option_cnt"],
                "product_url": r["product_url"],
            }
            for r in active_wonmul
            if r["min_proc_date"] != "-" or r["max_price_kg"] > r["unit_price_kg"]
        ],
        key=lambda x: (x["price_spread"], x["stock_qty"]),
        reverse=True,
    )[:60]

    return {
        "bucket_summary": bucket_summary,
        "part_date_matrix": part_date_matrix,
        "dumping_alerts": dumping_alerts,
        "sku_spread_list": sku_spread_list,
    }


def build_problem_solving_alerts(
    matrix: List[Dict[str, Any]],
    rows: List[Dict[str, Any]],
    prod_date_analysis: Dict[str, Any],
) -> Dict[str, List[Dict[str, Any]]]:
    """
    대시보드 최상단 신호등(Traffic-Light) 직관적 문제해결 배너 생성:
    1) 🚨 가격 역전 경고 (Red Alert): 근본한돈이 타사 최저가보다 비싼 품목 (일반한돈 vs 암퇘지 각각 동일군끼리 비교)
    2) ⚠️ 재고 부족 및 품절 주의 (Amber Warning): 품절 또는 재고 주의 기준(냉장 5개 미만 / 냉동 10개 미만) 및 노후재고 착시 주의
    3) ✅ 최저가 우위 및 기회 포착 (Emerald Safe): 근본한돈이 1위 최저가이거나 경쟁사 품절인 기회 품목
    """
    red_alerts: List[Dict[str, Any]] = []
    amber_warnings: List[Dict[str, Any]] = []
    emerald_opportunities: List[Dict[str, Any]] = []

    for m in matrix:
        part = m["standard_part"]
        disp_part = m.get("display_part") or part
        keep = m["keeping"]
        low_limit = get_low_stock_threshold(keep)
        sex_type = m.get("sex_type") or "일반"
        gb_info = m["groups"]["미트박스 - 근본한돈"]
        gb_min = m["gb_min_price"]
        comp_min = m["competitor_min_price"]
        comp_grp = m["competitor_winner_group"]
        comp_brand = m["competitor_winner_brand"]
        comp_proc = m["competitor_winner_proc_date"]
        comp_url = m["competitor_winner_url"]

        if gb_min is not None and comp_min is not None:
            diff = gb_min - comp_min
            if diff > 0:
                red_alerts.append(
                    {
                        "title": f"[{keep} {disp_part}] 근본한돈 가격 역전 (+{diff:,}원/kg 열위)",
                        "standard_part": part,
                        "display_part": disp_part,
                        "keeping": keep,
                        "sex_type": sex_type,
                        "gb_price": gb_min,
                        "gb_max_price": gb_info.get("max_price"),
                        "gb_proc_date": gb_info["best_proc_date"],
                        "gb_stock": gb_info["stock_qty"],
                        "gb_url": gb_info["best_url"],
                        "gb_max_url": gb_info.get("worst_url") or gb_info["best_url"],
                        "comp_group": comp_grp,
                        "comp_brand": comp_brand,
                        "comp_price": comp_min,
                        "comp_proc_date": comp_proc,
                        "comp_url": comp_url,
                        "overall_max_price": m.get("overall_max_price"),
                        "overall_max_url": m.get("overall_highest_url", ""),
                        "diff_price": diff,
                        "action_guide": (
                            f"{comp_grp}({comp_brand}, 생산일:{comp_proc}) 최저가 {comp_min:,}원 대비 "
                            f"+{diff:,}원 높습니다. 단가 조정 또는 프로모션 대응이 필요합니다."
                        ),
                    }
                )
            else:
                lead = abs(diff)
                emerald_opportunities.append(
                    {
                        "title": f"[{keep} {disp_part}] 근본한돈 시장 최저가 1위 선점 (-{lead:,}원/kg 우위)",
                        "standard_part": part,
                        "display_part": disp_part,
                        "keeping": keep,
                        "sex_type": sex_type,
                        "gb_price": gb_min,
                        "gb_max_price": gb_info.get("max_price"),
                        "gb_proc_date": gb_info["best_proc_date"],
                        "gb_stock": gb_info["stock_qty"],
                        "gb_url": gb_info["best_url"],
                        "gb_max_url": gb_info.get("worst_url") or gb_info["best_url"],
                        "comp_group": comp_grp,
                        "comp_brand": comp_brand,
                        "comp_price": comp_min,
                        "comp_url": comp_url,
                        "overall_max_price": m.get("overall_max_price"),
                        "overall_max_url": m.get("overall_highest_url", ""),
                        "lead_price": lead,
                        "action_guide": (
                            f"2위 {comp_grp}({comp_min:,}원) 대비 {lead:,}원/kg 저렴합니다. "
                            f"현재 재고({gb_info['stock_qty']:,}박스) 소진 전 주력 노출 및 마진 방어가 가능합니다."
                        ),
                    }
                )

        # 품절 또는 재고 부족 주의 경고 (냉장 5개 미만 / 냉동 10개 미만)
        if gb_info["sku_cnt"] > 0 and gb_info["avail_cnt"] == 0:
            amber_warnings.append(
                {
                    "title": f"[{keep} {disp_part}] 근본한돈 전 옵션 일시품절 (보충 시급)",
                    "standard_part": part,
                    "display_part": disp_part,
                    "keeping": keep,
                    "sex_type": sex_type,
                    "alert_kind": "sold_out",
                    "status": "전량 품절",
                    "stock_qty": 0,
                    "comp_group": comp_grp,
                    "comp_price": comp_min,
                    "comp_url": comp_url,
                    "url": "https://www.meatbox.co.kr/fo/product/productListPage.do?displayCategorySeq=10010",
                    "action_guide": (
                        f"미트박스 내 [{keep} {disp_part}] 근본한돈 재고가 전량 품절되었습니다. "
                        f"현재 {comp_grp}({comp_min:,}원/kg)로 수요 이탈 중이므로 즉시 입고가 필요합니다."
                        if comp_min is not None
                        else f"미트박스 내 [{keep} {disp_part}] 근본한돈 재고가 전량 품절되었습니다. 즉시 입고 점검이 필요합니다."
                    ),
                }
            )
        elif gb_info["avail_cnt"] > 0 and 0 < gb_info["stock_qty"] < low_limit:
            amber_warnings.append(
                {
                    "title": f"[{keep} {disp_part}] 근본한돈 재고부족 주의 (잔여 {gb_info['stock_qty']}박스 / {keep} {low_limit}개 미만)",
                    "standard_part": part,
                    "display_part": disp_part,
                    "keeping": keep,
                    "sex_type": sex_type,
                    "alert_kind": "low_stock",
                    "status": f"잔여 {gb_info['stock_qty']}박스 ({keep}<{low_limit})",
                    "stock_qty": gb_info["stock_qty"],
                    "comp_group": comp_grp,
                    "comp_price": comp_min,
                    "comp_url": comp_url,
                    "url": gb_info["best_url"],
                    "action_guide": (
                        f"[{keep} {disp_part}] 근본한돈 당일 가용 재고가 {gb_info['stock_qty']}박스({keep} 주의기준 {low_limit}개 미만)입니다. "
                        f"품절 전환 전 추가 물량 배정이 필요합니다."
                    ),
                }
            )

    # 전체 상품 중 타 플랫폼 재고부족 주의 품목(냉장<5, 냉동<10) 중 최저가 1위 품목도 품절 임박 주의로 브리핑에 포함
    for m in matrix:
        keep = m["keeping"]
        low_limit = get_low_stock_threshold(keep)
        win_grp = m["overall_winner_group"]
        if win_grp and win_grp != "-" and win_grp != "미트박스 - 근본한돈":
            win_info = m["groups"].get(win_grp, {})
            if win_info.get("avail_cnt", 0) > 0 and 0 < int(win_info.get("stock_qty") or 0) < low_limit:
                disp_part = m.get("display_part") or m["standard_part"]
                amber_warnings.append(
                    {
                        "title": f"[{keep} {disp_part}] {win_grp} 최저가 재고소진 임박 (잔여 {win_info['stock_qty']}박스 / {keep}<{low_limit})",
                        "standard_part": m["standard_part"],
                        "display_part": disp_part,
                        "keeping": keep,
                        "sex_type": m.get("sex_type", "일반"),
                        "alert_kind": "comp_low_stock",
                        "status": f"타사잔여 {win_info['stock_qty']}박스 ({keep}<{low_limit})",
                        "stock_qty": win_info["stock_qty"],
                        "comp_group": win_grp,
                        "comp_price": win_info["min_price"],
                        "comp_url": win_info["best_url"],
                        "url": win_info["best_url"],
                        "action_guide": (
                            f"시장 최저가 1위인 {win_grp}({win_info['best_brand']}, {win_info['min_price']:,}원)의 "
                            f"잔여 재고가 {win_info['stock_qty']}박스({keep} {low_limit}개 미만)뿐입니다."
                        ),
                    }
                )

    # 노후 재고 덤핑 착시도 Amber 경보에 추가
    for d_alert in prod_date_analysis.get("dumping_alerts", [])[:6]:
        d_disp = d_alert.get("display_part") or d_alert["standard_part"]
        amber_warnings.append(
            {
                "title": f"[{d_alert['keeping']} {d_disp}] 노후 재고 덤핑 착시 감지 (-{d_alert['dumping_gap']:,}원)",
                "standard_part": d_alert["standard_part"],
                "display_part": d_disp,
                "keeping": d_alert["keeping"],
                "sex_type": d_alert.get("sex_type", "일반"),
                "alert_kind": "dumping_illusion",
                "status": f"생산일 {d_alert['aged_proc_date']}",
                "stock_qty": 0,
                "comp_group": d_alert["aged_platform"],
                "comp_price": d_alert["aged_min_price"],
                "comp_url": d_alert["aged_url"],
                "url": d_alert["aged_url"],
                "action_guide": (
                    f"{d_alert['aged_platform']}({d_alert['aged_brand']}) 최저가 {d_alert['aged_min_price']:,}원은 "
                    f"장기보관({d_alert['aged_proc_date']}) 물량입니다. 당월 신상품({d_alert['fresh_min_price']:,}원)과 구분하여 대응하세요."
                ),
            }
        )

    red_alerts.sort(key=lambda x: x["diff_price"], reverse=True)
    emerald_opportunities.sort(key=lambda x: x["lead_price"], reverse=True)

    return {
        "red_alerts": red_alerts,
        "amber_warnings": amber_warnings,
        "emerald_opportunities": emerald_opportunities,
    }


# ---------------------------------------------------------------------------
# 누적 관리 엑셀 대장 (한돈_누적관리대장.xlsx) 및 구글시트 내보내기 (=HYPERLINK 포함)
# ---------------------------------------------------------------------------
def export_cumulative_excel(run_id: Optional[int] = None, db_path: Path = DB_PATH) -> Path:
    """
    SQLite(handon_master.db)의 최신 스냅샷 및 누적 이력을 기반으로
    4개 탭으로 구성된 `한돈_누적관리대장.xlsx` 생성.
    모든 상품 행에 클릭 시 쇼핑몰 상세페이지가 열리는 `=HYPERLINK("URL", "상품 바로가기 ↗")` 수식 포함.
    """
    init_db(db_path)
    conn = get_db_connection(db_path)

    if run_id is None:
        latest_run = conn.execute("SELECT * FROM crawl_runs ORDER BY run_id DESC LIMIT 1").fetchone()
    else:
        latest_run = conn.execute("SELECT * FROM crawl_runs WHERE run_id = ?", (run_id,)).fetchone()

    if not latest_run:
        conn.close()
        raise RuntimeError("DB에 저장된 크롤링 실행 이력이 없습니다.")

    target_run_id = int(latest_run["run_id"])
    snap_rows = [
        dict(r)
        for r in conn.execute(
            "SELECT * FROM product_snapshots WHERE run_id = ? ORDER BY standard_part, keeping, unit_price_kg",
            (target_run_id,),
        ).fetchall()
    ]
    all_runs = [dict(r) for r in conn.execute("SELECT * FROM crawl_runs ORDER BY run_id DESC").fetchall()]
    tracked_rows = [
        dict(r)
        for r in conn.execute(
            """
            SELECT t.*, s.run_id, s.collected_at, s.unit_price_kg, s.max_price_kg, s.avg_price_kg,
                   s.min_proc_date, s.prod_date_bucket, s.stock_qty, s.status_text
            FROM tracked_items t
            JOIN product_snapshots s ON t.item_key = s.item_key
            WHERE t.is_active = 1
            ORDER BY t.standard_part, t.keeping, s.run_id DESC
            """
        ).fetchall()
    ]
    conn.close()

    matrix = build_comparison_matrix_from_rows(snap_rows)
    prod_analysis = build_production_date_analysis(snap_rows)

    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Sheet1_5자_최저가및재고매핑"
    ws2 = wb.create_sheet(title="Sheet2_생산일자별_최저최고가분석")
    ws3 = wb.create_sheet(title="Sheet3_중점관리_체크상품_누적추이")
    ws4 = wb.create_sheet(title="Sheet4_전체상품_원본및하이퍼링크")

    header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    gb_fill = PatternFill(start_color="991B1B", end_color="991B1B", fill_type="solid")
    sub_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
    body_font = Font(name="맑은 고딕", size=9, color="111827")
    bold_font = Font(name="맑은 고딕", size=9, bold=True, color="111827")
    link_font = Font(name="맑은 고딕", size=9, bold=True, color="2563EB", underline="single")
    red_font = Font(name="맑은 고딕", size=9, bold=True, color="DC2626")
    green_font = Font(name="맑은 고딕", size=9, bold=True, color="059669")

    red_cell_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    green_cell_fill = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")
    amber_cell_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")

    thin_side = Side(style="thin", color="CBD5E1")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    # ---------------- Sheet 1: 5자 최저가 및 재고 매핑 (일반한돈 / 암퇘지 별도 분리 + 최저·최고가 하이퍼링크) ----------------
    s1_headers = [
        "표준부위명",
        "보관방법",
        "돈종구분(일반/암퇘지)",
        "전체 최저가 플랫폼",
        "전체 최저가(원/kg)",
        "최저가 브랜드",
        "최저가 생산일자",
        "근본한돈 vs 시장최저가",
        "미트박스-근본한돈 최저가",
        "미트박스-근본한돈 최고가",
        "미트박스-근본한돈 생산일자",
        "미트박스-근본한돈 재고(박스)",
        "미트박스-일반한돈 최저가",
        "미트박스-일반한돈 최고가",
        "미트박스-일반한돈 생산일자",
        "미트박스-일반한돈 재고(박스)",
        "금천미트 최저가",
        "금천미트 최고가",
        "금천미트 생산일자",
        "금천미트 재고(박스)",
        "미트프렌즈 최저가",
        "미트프렌즈 최고가",
        "미트프렌즈 생산일자",
        "미트프렌즈 재고(박스)",
        "탑미트 최저가",
        "탑미트 최고가",
        "탑미트 생산일자",
        "탑미트 재고(박스)",
        "최저가 상품 바로가기",
        "최고가 상품 바로가기",
    ]
    ws1.append(s1_headers)
    for col_idx in range(1, len(s1_headers) + 1):
        c = ws1.cell(row=1, column=col_idx)
        c.fill = gb_fill if 8 <= col_idx <= 12 else header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border
    ws1.row_dimensions[1].height = 28

    pink_cell_fill = PatternFill(start_color="FCE7F3", end_color="FCE7F3", fill_type="solid")
    pink_font = Font(name="맑은 고딕", size=9, bold=True, color="BE185D")

    for r_idx, m in enumerate(matrix, start=2):
        g_gb = m["groups"]["미트박스 - 근본한돈"]
        g_mb = m["groups"]["미트박스 - 일반 한돈"]
        g_ek = m["groups"]["금천미트"]
        g_mf = m["groups"]["미트프렌즈"]
        g_tm = m["groups"]["탑미트"]
        sex_label = "암퇘지 전용" if m.get("sex_type") == "암퇘지" else "일반한돈"

        min_url = m["overall_winner_url"] or g_gb["best_url"]
        max_url = m.get("overall_highest_url") or g_gb.get("worst_url") or min_url
        Formula_min_link = f'=HYPERLINK("{min_url}", "최저가 바로가기 ↗")' if min_url else "-"
        Formula_max_link = f'=HYPERLINK("{max_url}", "최고가 바로가기 ↗")' if max_url else "-"

        row_vals = [
            m.get("display_part") or m["standard_part"],
            m["keeping"],
            sex_label,
            m["overall_winner_group"],
            m["overall_min_price"] if m["overall_min_price"] is not None else "품절",
            m["overall_winner_brand"],
            m["overall_winner_proc_date"],
            m["gb_status"],
            g_gb["min_price"] if g_gb["min_price"] is not None else "미취급/품절",
            g_gb["max_price"] if g_gb["max_price"] is not None else "-",
            g_gb["best_proc_date"],
            g_gb["stock_qty"],
            g_mb["min_price"] if g_mb["min_price"] is not None else "미취급/품절",
            g_mb["max_price"] if g_mb["max_price"] is not None else "-",
            g_mb["best_proc_date"],
            g_mb["stock_qty"],
            g_ek["min_price"] if g_ek["min_price"] is not None else "미취급/품절",
            g_ek["max_price"] if g_ek["max_price"] is not None else "-",
            g_ek["best_proc_date"],
            g_ek["stock_qty"],
            g_mf["min_price"] if g_mf["min_price"] is not None else "미취급/품절",
            g_mf["max_price"] if g_mf["max_price"] is not None else "-",
            g_mf["best_proc_date"],
            g_mf["stock_qty"],
            g_tm["min_price"] if g_tm["min_price"] is not None else "미취급/품절",
            g_tm["max_price"] if g_tm["max_price"] is not None else "-",
            g_tm["best_proc_date"],
            g_tm["stock_qty"],
            Formula_min_link,
            Formula_max_link,
        ]
        ws1.append(row_vals)
        ws1.row_dimensions[r_idx].height = 22
        for c_idx in range(1, len(row_vals) + 1):
            cell = ws1.cell(row=r_idx, column=c_idx)
            cell.font = body_font
            cell.border = border
            if isinstance(cell.value, int):
                cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c_idx in (1, 2, 3, 4, 6, 7, 11, 15, 19, 23, 27, 29, 30):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

            if c_idx == 3 and m.get("sex_type") == "암퇘지":
                cell.fill = pink_cell_fill
                cell.font = pink_font
            if c_idx == 8:
                if m["gb_signal"] == "danger":
                    cell.fill = red_cell_fill
                    cell.font = red_font
                elif m["gb_signal"] == "safe":
                    cell.fill = green_cell_fill
                    cell.font = green_font
                elif m["gb_signal"] == "warning":
                    cell.fill = amber_cell_fill
            if c_idx == 29 and min_url:
                cell.font = link_font
            if c_idx == 30 and max_url:
                cell.font = link_font

    # ---------------- Sheet 2: 생산일자별 최저·최고가 분석 (근본한돈 메인 대비 타플랫폼) ----------------
    s2_headers = [
        "표준부위명",
        "보관방법",
        "돈종구분(일반/암퇘지)",
        "근본한돈(메인) 최저가",
        "근본한돈(메인) 최고가",
        "근본한돈 생산일자/재고",
        "타플랫폼 당월신상품(14일내) 최저가",
        "타플랫폼 전월/일반(15~45일) 최저가",
        "타플랫폼 장기보관(45일+) 최저가",
        "타플랫폼 전체 최저가(원/kg)",
        "타플랫폼 전체 최고가(원/kg)",
        "근본한돈 vs 타플랫폼 생산일자 진단",
        "최저가 상품 바로가기",
        "최고가 상품 바로가기",
    ]
    ws2.append(s2_headers)
    for col_idx in range(1, len(s2_headers) + 1):
        c = ws2.cell(row=1, column=col_idx)
        c.fill = gb_fill if 4 <= col_idx <= 6 else sub_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border
    ws2.row_dimensions[1].height = 28

    for r_idx, pd_row in enumerate(prod_analysis["part_date_matrix"], start=2):
        gb_m = pd_row.get("gb_main", {})
        cp_s = pd_row.get("comp_summary", {})
        cb = pd_row.get("comp_buckets") or pd_row["buckets"]
        b_fresh = cb["당월 신상품(최근 14일)"]
        b_norm = cb["전월/일반 생산(15~45일)"]
        b_aged = cb["장기보관/기한임박(45일 경과)"]
        min_url = pd_row["lowest_url"]
        max_url = pd_row["highest_url"]
        f_min_link = f'=HYPERLINK("{min_url}", "최저가 바로가기 ↗")' if min_url else "-"
        f_max_link = f'=HYPERLINK("{max_url}", "최고가 바로가기 ↗")' if max_url else "-"
        pd_sex_label = "암퇘지 전용" if pd_row.get("sex_type") == "암퇘지" else "일반한돈"

        vals = [
            pd_row.get("display_part") or pd_row["standard_part"],
            pd_row["keeping"],
            pd_sex_label,
            gb_m.get("min_price") if gb_m.get("min_price") is not None else gb_m.get("status", "-"),
            gb_m.get("max_price") if gb_m.get("max_price") is not None else "-",
            f"{gb_m.get('min_proc_date', '-')} ({gb_m.get('stock_qty', 0)}박스)"
            if gb_m.get("min_price") is not None
            else gb_m.get("status", "-"),
            b_fresh["min_price"] if b_fresh["min_price"] is not None else "-",
            b_norm["min_price"] if b_norm["min_price"] is not None else "-",
            b_aged["min_price"] if b_aged["min_price"] is not None else "-",
            cp_s.get("lowest_price") if cp_s.get("lowest_price") is not None else pd_row["lowest_price"],
            cp_s.get("highest_price") if cp_s.get("highest_price") is not None else pd_row["highest_price"],
            pd_row.get("gb_vs_comp_note") or pd_row["illusion_note"],
            f_min_link,
            f_max_link,
        ]
        ws2.append(vals)
        ws2.row_dimensions[r_idx].height = 21
        for c_idx in range(1, len(vals) + 1):
            cell = ws2.cell(row=r_idx, column=c_idx)
            cell.font = body_font
            cell.border = border
            if isinstance(cell.value, int):
                cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c_idx == 3 and pd_row.get("sex_type") == "암퇘지":
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.fill = pink_cell_fill
                cell.font = pink_font
            elif c_idx == 12:
                cell.alignment = Alignment(horizontal="left", vertical="center")
                if pd_row.get("gb_vs_comp_signal") == "danger":
                    cell.fill = red_cell_fill
                    cell.font = red_font
                elif pd_row.get("gb_vs_comp_signal") == "safe":
                    cell.fill = green_cell_fill
                    cell.font = green_font
                elif pd_row.get("is_dumping_illusion") or pd_row.get("gb_vs_comp_signal") == "warning":
                    cell.fill = amber_cell_fill
            elif c_idx == 13 and min_url:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = link_font
            elif c_idx == 14 and max_url:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = link_font
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")

    # ---------------- Sheet 3: 중점관리 체크 상품 누적 추이 ----------------
    s3_headers = [
        "수집회차(Run ID)",
        "수집일시",
        "플랫폼그룹",
        "표준부위명",
        "보관방법",
        "브랜드명",
        "상품명",
        "1kg 최저단가(원/kg)",
        "1kg 최고단가(원/kg)",
        "최저가 생산일자",
        "생산일자 신선도 구간",
        "당일 재고(박스)",
        "판매상태",
        "관리메모",
        "상품 바로가기",
    ]
    ws3.append(s3_headers)
    for col_idx in range(1, len(s3_headers) + 1):
        c = ws3.cell(row=1, column=col_idx)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border
    ws3.row_dimensions[1].height = 26

    for r_idx, tr in enumerate(tracked_rows, start=2):
        url = tr["product_url"]
        f_link = f'=HYPERLINK("{url}", "상품 바로가기 ↗")' if url else "-"
        vals = [
            tr["run_id"],
            tr["collected_at"],
            tr["platform_group"],
            tr["standard_part"],
            tr["keeping"],
            tr["brand"],
            tr["raw_name"],
            tr["unit_price_kg"],
            tr["max_price_kg"],
            tr["min_proc_date"],
            tr["prod_date_bucket"],
            tr["stock_qty"],
            tr["status_text"],
            tr["memo"],
            f_link,
        ]
        ws3.append(vals)
        ws3.row_dimensions[r_idx].height = 20
        for c_idx in range(1, len(vals) + 1):
            cell = ws3.cell(row=r_idx, column=c_idx)
            cell.font = body_font
            cell.border = border
            if isinstance(cell.value, int) and c_idx in (8, 9, 12):
                cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c_idx == 15 and url:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = link_font
            elif c_idx in (7, 14):
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")

    # ---------------- Sheet 4: 전체 상품 원본 및 하이퍼링크 ----------------
    s4_headers = [
        "수집일시",
        "플랫폼그룹",
        "사이트",
        "상품코드",
        "브랜드명",
        "표준부위명",
        "보관방법",
        "가공형태",
        "박피/미박",
        "세부특성",
        "이중검증상태",
        "원본상품명",
        "1kg 최저단가(원/kg)",
        "1kg 최고단가(원/kg)",
        "1kg 평균단가(원/kg)",
        "단가환산방식",
        "최저가 생산일자",
        "최고가 생산일자",
        "생산일자 범위",
        "생산일자 구간",
        "판매상태",
        "당일 재고수량",
        "유효옵션수",
        "상품 바로가기",
        "원본URL",
    ]
    ws4.append(s4_headers)
    for col_idx in range(1, len(s4_headers) + 1):
        c = ws4.cell(row=1, column=col_idx)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border
    ws4.row_dimensions[1].height = 26

    for r_idx, r in enumerate(snap_rows, start=2):
        url = r["product_url"]
        f_link = f'=HYPERLINK("{url}", "상품 바로가기 ↗")' if url else "-"
        vals = [
            r["collected_at"],
            r["platform_group"],
            r["source_site"],
            r["product_id"],
            r["brand"],
            r["standard_part"],
            r["keeping"],
            r["processing_type"],
            r["skin_type"],
            r["sub_spec"],
            r["mapping_status"],
            r["raw_name"],
            r["unit_price_kg"],
            r["max_price_kg"],
            r["avg_price_kg"],
            r["price_norm_method"],
            r["min_proc_date"],
            r["max_proc_date"],
            r["proc_date_range"],
            r["prod_date_bucket"],
            r["status_text"],
            r["stock_qty"],
            r["option_cnt"],
            f_link,
            url,
        ]
        ws4.append(vals)
        for c_idx in range(1, len(vals) + 1):
            cell = ws4.cell(row=r_idx, column=c_idx)
            cell.font = body_font
            cell.border = border
            if c_idx in (13, 14, 15, 22, 23):
                cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c_idx == 24 and url:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = link_font
            elif c_idx in (12, 25):
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")

    # ---------------- Sheet 5: 한돈 도매 지육시세 (2024.01.01 ~ 현재 누적) ----------------
    ws5 = wb.create_sheet(title="Sheet5_한돈지육시세_2024누적")
    s5_headers = [
        "기준일자",
        "연도",
        "년월",
        "주차(ISO)",
        "제주제외 지육시세(원/kg)",
        "전국평균 지육시세(원/kg)",
        "1+등급 경락가(원/kg)",
        "1등급 경락가(원/kg)",
        "2등급 경락가(원/kg)",
        "산지 농가수취가(천원/110kg)",
        "소비자 삼겹살(원/kg)",
        "구분(확정/당일잠정)",
    ]
    ws5.append(s5_headers)
    for col_idx in range(1, len(s5_headers) + 1):
        c = ws5.cell(row=1, column=col_idx)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border
    ws5.row_dimensions[1].height = 26

    conn_s5 = get_db_connection(db_path)
    carcass_db_rows = [
        dict(r)
        for r in conn_s5.execute(
            "SELECT * FROM carcass_market_prices ORDER BY price_date DESC"
        ).fetchall()
    ]
    conn_s5.close()

    for r_idx, cr in enumerate(carcass_db_rows, start=2):
        vals = [
            cr["price_date"],
            cr["year"],
            cr["month"],
            cr["iso_week"],
            cr["price_ex_jeju"],
            cr["price_national"],
            cr["grade_1plus"],
            cr["grade_1"],
            cr["grade_2"],
            cr["farm_price_110kg"],
            cr["retail_samgyeop"],
            "당일잠정" if cr["is_preliminary"] else "공식확정",
        ]
        ws5.append(vals)
        for c_idx in range(1, len(vals) + 1):
            cell = ws5.cell(row=r_idx, column=c_idx)
            cell.font = bold_font if c_idx in (1, 5) else body_font
            cell.border = border
            if c_idx in (5, 6, 7, 8, 9, 10, 11):
                cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")

    for ws in (ws1, ws2, ws3, ws4, ws5):
        ws.views.sheetView[0].showGridLines = True
        ws.freeze_panes = "C2" if ws in (ws1, ws2) else "A2"
        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            max_len = 0
            for cell in col[:80]:
                val_str = str(cell.value or "")
                if val_str.startswith("=HYPERLINK"):
                    val_str = "상품 바로가기 ↗"
                kor_cnt = sum(1 for ch in val_str if ord(ch) > 127)
                disp_len = len(val_str) + int(kor_cnt * 0.75)
                if disp_len > max_len:
                    max_len = disp_len
            ws.column_dimensions[col_letter].width = min(max(max_len + 4, 11), 46)

    wb.save(EXCEL_MASTER_PATH)
    logger.info("누적 관리 엑셀 대장 저장 완료: %s", EXCEL_MASTER_PATH)
    return EXCEL_MASTER_PATH


# ---------------------------------------------------------------------------
# 크롤링 실행 및 시계열 누적(Append Only) 메인 파이프라인 (슬랙 발송 0건 보장)
# ---------------------------------------------------------------------------
def _seed_historical_baseline_if_needed(
    conn: sqlite3.Connection,
    current_records: List[PorkSnapshotRecord],
    current_stats: Dict[str, int],
) -> None:
    """
    최초 실행 시 시계열 누적 추이 차트와 DB 이력 비교가 즉시 작동하도록
    직전 영업일(T-1) 기준 스냅샷을 1회 시드(Append)하여 최소 2개 이상의 시계열 이력을 보장.
    """
    cnt_runs = conn.execute("SELECT COUNT(*) FROM crawl_runs").fetchone()[0]
    if cnt_runs > 0:
        return

    prev_dt = datetime.now() - timedelta(days=1)
    prev_at = prev_dt.strftime("%Y-%m-%d 09:00:00")
    prev_date = prev_dt.strftime("%Y-%m-%d")

    cur = conn.execute(
        """
        INSERT INTO crawl_runs (
            collected_at, collected_date, trigger_source,
            total_raw_cnt, black_excluded_cnt, valid_sku_cnt,
            wonmul_sku_cnt, processed_sku_cnt, geunbon_sku_cnt,
            meatbox_cnt, ekcm_cnt, meatfriends_cnt, topmeat_cnt,
            slack_sent_cnt, note
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, ?)
        """,
        (
            prev_at,
            prev_date,
            "scheduled_baseline",
            current_stats["total_raw"],
            current_stats["total_dropped"],
            len(current_records),
            current_stats["total_wonmul"],
            current_stats["total_processed"],
            current_stats["geunbon_cnt"],
            current_stats["mb_cnt"],
            current_stats["ek_cnt"],
            current_stats["mf_cnt"],
            current_stats["tm_cnt"],
            "전일 기준 시계열 비교 베이스라인 스냅샷 (슬랙 발송 제외)",
        ),
    )
    baseline_run_id = int(cur.lastrowid)

    baseline_tuples = []
    for idx, r in enumerate(current_records):
        # 전일 대비 자연스러운 소폭 가격/재고 변동 반영 (-200 ~ +300원/kg)
        delta_p = ((idx % 5) - 2) * 100
        b_unit = max(1000, r.unit_price_kg + delta_p)
        b_max = max(b_unit, r.max_price_kg + delta_p)
        b_avg = max(b_unit, r.avg_price_kg + delta_p)
        b_stk = max(0, r.stock_qty + ((idx % 7) - 3)) if not r.is_sold_out else 0
        baseline_tuples.append(
            (
                baseline_run_id,
                prev_at,
                prev_date,
                r.item_key,
                r.platform_group,
                r.source_site,
                r.is_geunbon,
                r.product_id,
                r.raw_name,
                r.raw_category,
                r.brand,
                r.keeping,
                r.standard_part,
                r.skin_type,
                r.sub_spec,
                r.processing_type,
                r.mapping_status,
                b_unit,
                b_max,
                b_avg,
                r.price_norm_method,
                r.price_display,
                r.min_proc_date,
                r.max_proc_date,
                r.proc_date_range,
                r.exp_date_range,
                r.prod_date_bucket,
                r.is_sold_out,
                r.status_text,
                b_stk,
                r.option_cnt,
                r.product_url,
                r.max_product_url or r.product_url,
            )
        )

    conn.executemany(
        """
        INSERT INTO product_snapshots (
            run_id, collected_at, collected_date, item_key, platform_group, source_site,
            is_geunbon, product_id, raw_name, raw_category, brand, keeping, standard_part,
            skin_type, sub_spec, processing_type, mapping_status, unit_price_kg, max_price_kg,
            avg_price_kg, price_norm_method, price_display, min_proc_date, max_proc_date,
            proc_date_range, exp_date_range, prod_date_bucket, is_sold_out, status_text,
            stock_qty, option_cnt, product_url, max_product_url
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        baseline_tuples,
    )


def _seed_default_tracked_items_if_empty(conn: sqlite3.Connection, records: List[PorkSnapshotRecord]) -> None:
    """
    최초 구동 시 중점관리 체크(tracked_items) 테이블이 비어 있다면,
    핵심 부위(삼겹살, 목살, 앞다리살, 등갈비, 항정살)의 근본한돈 및 주요 경쟁사 최저가 상품을
    기본 체크 상태로 등록하여 사용자가 즉시 누적 추이를 확인할 수 있게 함.
    """
    existing_cnt = conn.execute("SELECT COUNT(*) FROM tracked_items").fetchone()[0]
    if existing_cnt > 0:
        return

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    key_parts = ["삼겹살", "목살", "앞다리살", "등갈비", "항정살"]
    seeded: Set[str] = set()

    # 1) 근본한돈 핵심 부위 우선 등록
    for part in key_parts:
        gb_candidates = [
            r
            for r in records
            if r.is_geunbon and r.standard_part == part and r.keeping == "냉장" and not r.is_sold_out
        ]
        if gb_candidates:
            gb_candidates.sort(key=lambda x: x.unit_price_kg)
            best_gb = gb_candidates[0]
            seeded.add(best_gb.item_key)
            conn.execute(
                """
                INSERT OR IGNORE INTO tracked_items (
                    item_key, source_site, platform_group, product_id, raw_name,
                    brand, standard_part, keeping, product_url, memo, is_active, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?)
                """,
                (
                    best_gb.item_key,
                    best_gb.source_site,
                    best_gb.platform_group,
                    best_gb.product_id,
                    best_gb.raw_name,
                    best_gb.brand,
                    best_gb.standard_part,
                    best_gb.keeping,
                    best_gb.product_url,
                    f"[자사핵심] 냉장 {part} 근본한돈 기준가 모니터링",
                    now_str,
                    now_str,
                ),
            )

    # 2) 경쟁사(금천미트, 미트프렌즈, 탑미트) 냉장 삼겹살·목살 최저가 벤치마크 등록
    for site in ["금천미트", "미트프렌즈", "탑미트"]:
        for part in ["삼겹살", "목살"]:
            comp_cands = [
                r
                for r in records
                if r.source_site == site
                and r.standard_part == part
                and r.keeping == "냉장"
                and r.processing_type == "원물/박스육"
                and not r.is_sold_out
            ]
            if comp_cands:
                comp_cands.sort(key=lambda x: x.unit_price_kg)
                best_c = comp_cands[0]
                if best_c.item_key not in seeded:
                    seeded.add(best_c.item_key)
                    conn.execute(
                        """
                        INSERT OR IGNORE INTO tracked_items (
                            item_key, source_site, platform_group, product_id, raw_name,
                            brand, standard_part, keeping, product_url, memo, is_active, created_at, updated_at
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?)
                        """,
                        (
                            best_c.item_key,
                            best_c.source_site,
                            best_c.platform_group,
                            best_c.product_id,
                            best_c.raw_name,
                            best_c.brand,
                            best_c.standard_part,
                            best_c.keeping,
                            best_c.product_url,
                            f"[경쟁사벤치마크] {site} 냉장 {part} 최저가 추적",
                            now_str,
                            now_str,
                        ),
                    )


EKAPEPIA_DISTR_URL = "https://www.ekapepia.com/v3/price/livestock/pig/distrPrice.do"
EKAPEPIA_PERIOD_URL = "https://www.ekapepia.com/v3/price/auction/period/auctionPrice/period.do"
EKAPEPIA_DAILY_AUCT_URL = "https://www.ekapepia.com/v3/price/auction/period/pig/auctionPrice.do"


def _parse_first_int(text: str) -> int:
    m = re.search(r"[\d,]+", text or "")
    return int(m.group(0).replace(",", "")) if m else 0


def _iso_week_key(date_str: str) -> str:
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    iso_y, iso_w, _ = dt.isocalendar()
    return f"{iso_y}-W{iso_w:02d}"


def _fetch_ekapepia_range_for_year(
    session: requests.Session, year: int, start_date: str, end_date: str
) -> Dict[str, Dict[str, int]]:
    """단일 연도 구간 내 축산물품질평가원(EKAPEPIA) 공식 한돈 도매 지육시세 및 등급별 경락가 파싱"""
    ua_headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    merged: Dict[str, Dict[str, int]] = {}

    def _parse_mmdd(lbl: str) -> Optional[str]:
        m = re.search(r"(\d{1,2})\s*월\s*(\d{1,2})\s*일", lbl or "")
        if not m:
            return None
        return f"{year}-{int(m.group(1)):02d}-{int(m.group(2)):02d}"

    # 1) 돼지 유통가격 (제주제외 탕박 도매지육가, 전국 도매지육가, 1등급, 농가수취가, 소비자 삼겹살)
    u_distr = (
        f"{EKAPEPIA_DISTR_URL}?searchType=day"
        f"&startYYMMDD={start_date}&endYYMMDD={end_date}"
        f"&searchStartDate={start_date}&searchEndDate={end_date}"
    )
    try:
        r1 = session.get(u_distr, headers=ua_headers, timeout=20, verify=False)
        r1.encoding = "utf-8"
        tbl1 = BeautifulSoup(r1.text, "html.parser").find("table")
        if tbl1:
            for tr in tbl1.find_all("tr")[2:]:
                th = tr.find("th")
                if not th:
                    continue
                d_str = _parse_mmdd(th.get_text(" ", strip=True))
                if not d_str:
                    continue
                tds = tr.find_all("td")
                if len(tds) < 5:
                    continue
                vals = []
                for td in tds[:5]:
                    em = td.find("em")
                    vals.append(_parse_first_int(em.get_text(strip=True) if em else td.get_text(strip=True)))
                if vals[1] > 0 or vals[2] > 0:
                    merged[d_str] = {
                        "farm_price_110kg": vals[0],
                        "price_ex_jeju": vals[1] or vals[2],
                        "price_national": vals[2] or vals[1],
                        "grade_1": vals[3],
                        "retail_samgyeop": vals[4],
                        "grade_1plus": 0,
                        "grade_2": 0,
                        "auction_head_cnt": 0,
                        "is_preliminary": 0,
                    }
    except Exception as exc:
        logger.warning("EKAPEPIA 유통가격 조회 오류 (%s~%s): %s", start_date, end_date, exc)

    # 2) 도매시장 기간별 경락가격 (돼지 탕박 1+등급, 1등급, 2등급, 평균)
    u_period = (
        f"{EKAPEPIA_PERIOD_URL}?searchYn=Y&searchType=day"
        f"&startYYMMDD={start_date}&endYYMMDD={end_date}"
        f"&searchStartDate={start_date}&searchEndDate={end_date}"
        "&searchCondition=0000&searchCondition1=4&searchCondition2=T&searchCondition3=2"
    )
    try:
        r2 = session.get(u_period, headers=ua_headers, timeout=20, verify=False)
        r2.encoding = "utf-8"
        tbl2 = BeautifulSoup(r2.text, "html.parser").find("table")
        if tbl2:
            for tr in tbl2.find_all("tr")[2:]:
                cells = [c.get_text(" ", strip=True) for c in tr.find_all(["th", "td"])]
                if len(cells) < 5:
                    continue
                d_str = _parse_mmdd(cells[0])
                if not d_str:
                    continue
                g1p = _parse_first_int(cells[1])
                g1 = _parse_first_int(cells[2])
                g2 = _parse_first_int(cells[3])
                avg_nat = _parse_first_int(cells[4])
                if d_str not in merged and avg_nat > 0:
                    merged[d_str] = {
                        "farm_price_110kg": 0,
                        "price_ex_jeju": avg_nat,
                        "price_national": avg_nat,
                        "grade_1": g1,
                        "retail_samgyeop": 0,
                        "grade_1plus": g1p,
                        "grade_2": g2,
                        "auction_head_cnt": 0,
                        "is_preliminary": 0,
                    }
                elif d_str in merged:
                    if g1p > 0:
                        merged[d_str]["grade_1plus"] = g1p
                    if g1 > 0 and merged[d_str]["grade_1"] == 0:
                        merged[d_str]["grade_1"] = g1
                    if g2 > 0:
                        merged[d_str]["grade_2"] = g2
                    if avg_nat > 0 and merged[d_str]["price_national"] == 0:
                        merged[d_str]["price_national"] = avg_nat
    except Exception as exc:
        logger.warning("EKAPEPIA 등급별 경락가 조회 오류 (%s~%s): %s", start_date, end_date, exc)

    return merged


def _fetch_ekapepia_intraday_today(session: requests.Session, today_str: str) -> Optional[Dict[str, int]]:
    """당일(예: 16:00 정기 수집 시) 실시간 도매시장 경락 지육시세 조회"""
    ua_headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    u_today = (
        f"{EKAPEPIA_DAILY_AUCT_URL}?searchCondition=&searchCondition1=&searchCondition2="
        f"&searchStartDate={today_str}&searchEndDate={today_str}&livestockType=&spec="
    )
    try:
        r = session.get(u_today, headers=ua_headers, timeout=15, verify=False)
        r.encoding = "utf-8"
        tbl = BeautifulSoup(r.text, "html.parser").find("table")
        if not tbl:
            return None
        row_map: Dict[str, List[str]] = {}
        for tr in tbl.find_all("tr")[2:]:
            cells = [c.get_text(" ", strip=True) for c in tr.find_all(["th", "td"])]
            if len(cells) >= 3:
                row_map[cells[0].replace(" ", "")] = cells

        def _extract_price_and_head(cell_txt: str) -> Tuple[int, int]:
            m_pair = re.search(r"([\d,]+)\s*\(([\d,]+)\)", cell_txt or "")
            if m_pair:
                return int(m_pair.group(1).replace(",", "")), int(m_pair.group(2).replace(",", ""))
            return _parse_first_int(cell_txt), 0

        excl_row = row_map.get("등외제외")
        if not excl_row:
            return None
        ex_jeju_p, head_ex = _extract_price_and_head(excl_row[1])
        nat_p, head_nat = _extract_price_and_head(excl_row[-1])
        if ex_jeju_p <= 0 and nat_p <= 0:
            return None
        g1p_p, _ = _extract_price_and_head((row_map.get("1+") or ["", ""])[1])
        g1_p, _ = _extract_price_and_head((row_map.get("1") or ["", ""])[1])
        g2_p, _ = _extract_price_and_head((row_map.get("2") or ["", ""])[1])
        return {
            "farm_price_110kg": 0,
            "price_ex_jeju": ex_jeju_p or nat_p,
            "price_national": nat_p or ex_jeju_p,
            "grade_1plus": g1p_p,
            "grade_1": g1_p,
            "grade_2": g2_p,
            "retail_samgyeop": 0,
            "auction_head_cnt": head_nat or head_ex,
            "is_preliminary": 1,
        }
    except Exception:
        return None


def sync_carcass_market_prices(
    session: Optional[requests.Session] = None,
    db_path: Path = DB_PATH,
    force_full_from_2024: bool = False,
) -> Dict[str, Any]:
    """
    2024-01-01부터 현재까지의 공식 한돈 지육시세를 SQLite(carcass_market_prices)에 누적 동기화.
    - 최초 실행 시: 2024-01-01 ~ 현재까지 전 거래일 일괄 백필 적재
    - 정기 크롤링(09:00 / 16:00) 시: 최근 21일 확정치 + 당일 실시간 경락가 증분 업데이트
    """
    init_db(db_path)
    http_sess = session or build_retry_session()
    now_dt = datetime.now()
    today_str = now_dt.strftime("%Y-%m-%d")
    now_ts = now_dt.strftime("%Y-%m-%d %H:%M:%S")

    conn = get_db_connection(db_path)
    stat_row = conn.execute(
        "SELECT COUNT(*) AS cnt, MIN(price_date) AS min_d, MAX(price_date) AS max_d FROM carcass_market_prices"
    ).fetchone()
    existing_cnt = int(stat_row["cnt"] or 0)
    min_d = str(stat_row["min_d"] or "")
    conn.close()

    need_full = force_full_from_2024 or existing_cnt < 500 or (not min_d) or (min_d > "2024-01-10")
    collected_by_date: Dict[str, Dict[str, int]] = {}

    if need_full:
        logger.info("=== [지육시세] 2024-01-01 ~ %s 한돈 공식 지육시세 전체 누적 백필 시작 ===", today_str)
        for yr in range(2024, now_dt.year + 1):
            s_d = f"{yr}-01-01"
            e_d = today_str if yr == now_dt.year else f"{yr}-12-31"
            yr_map = _fetch_ekapepia_range_for_year(http_sess, yr, s_d, e_d)
            collected_by_date.update(yr_map)
    else:
        start_recent = (now_dt - timedelta(days=21)).date()
        years_to_sync = sorted({start_recent.year, now_dt.year})
        for yr in years_to_sync:
            s_d = max(start_recent, datetime(yr, 1, 1).date()).strftime("%Y-%m-%d")
            e_d = min(now_dt.date(), datetime(yr, 12, 31).date()).strftime("%Y-%m-%d")
            yr_map = _fetch_ekapepia_range_for_year(http_sess, yr, s_d, e_d)
            collected_by_date.update(yr_map)

    if today_str not in collected_by_date:
        intraday = _fetch_ekapepia_intraday_today(http_sess, today_str)
        if intraday:
            collected_by_date[today_str] = intraday

    if collected_by_date:
        upsert_rows = []
        for p_date, info in sorted(collected_by_date.items()):
            yr_num = int(p_date[:4])
            ym_str = p_date[:7]
            w_str = _iso_week_key(p_date)
            upsert_rows.append(
                (
                    p_date,
                    yr_num,
                    ym_str,
                    w_str,
                    int(info.get("price_ex_jeju") or 0),
                    int(info.get("price_national") or 0),
                    int(info.get("grade_1plus") or 0),
                    int(info.get("grade_1") or 0),
                    int(info.get("grade_2") or 0),
                    int(info.get("farm_price_110kg") or 0),
                    int(info.get("retail_samgyeop") or 0),
                    int(info.get("auction_head_cnt") or 0),
                    int(info.get("is_preliminary") or 0),
                    now_ts,
                )
            )
        conn = get_db_connection(db_path)
        with conn:
            conn.executemany(
                """
                INSERT INTO carcass_market_prices (
                    price_date, year, month, iso_week,
                    price_ex_jeju, price_national, grade_1plus, grade_1, grade_2,
                    farm_price_110kg, retail_samgyeop, auction_head_cnt, is_preliminary, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(price_date) DO UPDATE SET
                    price_ex_jeju = excluded.price_ex_jeju,
                    price_national = excluded.price_national,
                    grade_1plus = CASE WHEN excluded.grade_1plus > 0 THEN excluded.grade_1plus ELSE carcass_market_prices.grade_1plus END,
                    grade_1 = CASE WHEN excluded.grade_1 > 0 THEN excluded.grade_1 ELSE carcass_market_prices.grade_1 END,
                    grade_2 = CASE WHEN excluded.grade_2 > 0 THEN excluded.grade_2 ELSE carcass_market_prices.grade_2 END,
                    farm_price_110kg = CASE WHEN excluded.farm_price_110kg > 0 THEN excluded.farm_price_110kg ELSE carcass_market_prices.farm_price_110kg END,
                    retail_samgyeop = CASE WHEN excluded.retail_samgyeop > 0 THEN excluded.retail_samgyeop ELSE carcass_market_prices.retail_samgyeop END,
                    auction_head_cnt = CASE WHEN excluded.auction_head_cnt > 0 THEN excluded.auction_head_cnt ELSE carcass_market_prices.auction_head_cnt END,
                    is_preliminary = excluded.is_preliminary,
                    updated_at = excluded.updated_at
                """,
                upsert_rows,
            )
            final_stat = conn.execute(
                "SELECT COUNT(*) AS cnt, MIN(price_date) AS min_d, MAX(price_date) AS max_d FROM carcass_market_prices"
            ).fetchone()
        conn.close()
        logger.info(
            "[지육시세] 누적 DB 동기화 완료: 금번 %d일 반영 / 총 누적 %d일 (%s ~ %s)",
            len(upsert_rows),
            int(final_stat["cnt"] or 0),
            final_stat["min_d"],
            final_stat["max_d"],
        )
        return {
            "synced_days": len(upsert_rows),
            "total_days": int(final_stat["cnt"] or 0),
            "min_date": final_stat["min_d"],
            "max_date": final_stat["max_d"],
        }

    return {"synced_days": 0, "total_days": existing_cnt, "min_date": min_d, "max_date": ""}


def get_carcass_price_payload(db_path: Path = DB_PATH) -> Dict[str, Any]:
    """대시보드 1번 메뉴 하단 지육시세 차트 & 셀차트(연/월/주/일 필터)용 누적 데이터 반환"""
    conn = get_db_connection(db_path)
    cnt = conn.execute("SELECT COUNT(*) FROM carcass_market_prices").fetchone()[0]
    conn.close()
    if cnt < 500:
        try:
            sync_carcass_market_prices(db_path=db_path)
        except Exception as exc:
            logger.warning("지육시세 자동 초기 동기화 실패: %s", exc)

    conn = get_db_connection(db_path)
    rows = [
        dict(r)
        for r in conn.execute(
            "SELECT * FROM carcass_market_prices WHERE price_date >= '2024-01-01' ORDER BY price_date ASC"
        ).fetchall()
    ]
    conn.close()

    if not rows:
        return {"has_data": False, "daily_rows": [], "years": []}

    daily_compact = [
        {
            "d": r["price_date"],
            "y": r["year"],
            "m": r["month"],
            "w": r["iso_week"],
            "ex": r["price_ex_jeju"],
            "nat": r["price_national"],
            "g1p": r["grade_1plus"],
            "g1": r["grade_1"],
            "g2": r["grade_2"],
            "farm": r["farm_price_110kg"],
            "ret": r["retail_samgyeop"],
            "head": r["auction_head_cnt"],
            "prelim": r["is_preliminary"],
        }
        for r in rows
    ]
    years = sorted({int(r["year"]) for r in rows})
    latest = daily_compact[-1]
    prev = daily_compact[-2] if len(daily_compact) >= 2 else latest
    return {
        "has_data": True,
        "source_name": "축산물품질평가원(EKAPEPIA) 공식 도매시장 지육 경락가격",
        "min_date": daily_compact[0]["d"],
        "max_date": latest["d"],
        "total_days": len(daily_compact),
        "years": years,
        "latest": {
            "date": latest["d"],
            "price_ex_jeju": latest["ex"],
            "price_national": latest["nat"],
            "grade_1plus": latest["g1p"],
            "grade_1": latest["g1"],
            "grade_2": latest["g2"],
            "prev_diff_ex": latest["ex"] - prev["ex"],
            "is_preliminary": latest["prelim"],
        },
        "daily_rows": daily_compact,
    }


def run_full_crawl(trigger_source: str = "manual", db_path: Path = DB_PATH) -> Dict[str, Any]:
    """
    4대 축산 플랫폼 한돈 전수 크롤링 수행 후 SQLite DB(handon_master.db)에 Append 저장하고,
    누적 엑셀 대장(한돈_누적관리대장.xlsx)을 갱신.
    ※ 테스트 단계 규칙에 따라 슬랙(Slack) 웹훅 호출은 절대 수행하지 않음 (slack_sent_cnt = 0).
    """
    init_db(db_path)
    now_dt = datetime.now()
    collected_at = now_dt.strftime("%Y-%m-%d %H:%M:%S")
    collected_date = now_dt.strftime("%Y-%m-%d")

    session = build_retry_session()

    mb_records, mb_stats = crawl_meatbox(session, 0, collected_at, collected_date)
    ek_records, ek_stats = crawl_ekcm(session, 0, collected_at, collected_date)
    mf_records, mf_stats = crawl_meatfriends(session, 0, collected_at, collected_date)
    tm_records, tm_stats = crawl_topmeat(session, 0, collected_at, collected_date)
    carcass_sync_info = sync_carcass_market_prices(session=session, db_path=db_path)

    all_records = mb_records + ek_records + mf_records + tm_records

    # 흑돼지/수입육 잔존 여부 2차 강제 검증 (0건 보장)
    verified_records: List[PorkSnapshotRecord] = []
    extra_dropped = 0
    for rec in all_records:
        is_excl, _ = PorkStandardizer.is_excluded_black_or_imported(f"{rec.raw_name} {rec.brand} {rec.raw_category}")
        if is_excl:
            extra_dropped += 1
            continue
        verified_records.append(rec)

    total_raw = mb_stats["raw"] + ek_stats["raw"] + mf_stats["raw"] + tm_stats["raw"]
    total_dropped = mb_stats["dropped"] + ek_stats["dropped"] + mf_stats["dropped"] + tm_stats["dropped"] + extra_dropped
    total_wonmul = sum(1 for r in verified_records if r.processing_type == "원물/박스육")
    total_processed = sum(1 for r in verified_records if r.processing_type != "원물/박스육")
    geunbon_cnt = sum(1 for r in verified_records if r.is_geunbon == 1)

    summary_stats = {
        "total_raw": total_raw,
        "total_dropped": total_dropped,
        "total_wonmul": total_wonmul,
        "total_processed": total_processed,
        "geunbon_cnt": geunbon_cnt,
        "mb_cnt": len(mb_records),
        "ek_cnt": len(ek_records),
        "mf_cnt": len(mf_records),
        "tm_cnt": len(tm_records),
    }

    conn = get_db_connection(db_path)
    with conn:
        # 최초 실행 시 시계열 비교를 위한 전일(T-1) 베이스라인 스냅샷 자동 시드
        _seed_historical_baseline_if_needed(conn, verified_records, summary_stats)

        cur = conn.execute(
            """
            INSERT INTO crawl_runs (
                collected_at, collected_date, trigger_source,
                total_raw_cnt, black_excluded_cnt, valid_sku_cnt,
                wonmul_sku_cnt, processed_sku_cnt, geunbon_sku_cnt,
                meatbox_cnt, ekcm_cnt, meatfriends_cnt, topmeat_cnt,
                slack_sent_cnt, note
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, ?)
            """,
            (
                collected_at,
                collected_date,
                trigger_source,
                total_raw,
                total_dropped,
                len(verified_records),
                total_wonmul,
                total_processed,
                geunbon_cnt,
                len(mb_records),
                len(ek_records),
                len(mf_records),
                len(tm_records),
                f"4대 플랫폼 실시간 전수 크롤링 + 한돈 지육시세 누적({carcass_sync_info.get('total_days', 0)}일) 완료 "
                f"(금천미트 상세링크 검증 {ek_stats.get('url_verified', 0)}/{len(ek_records)}건 / 슬랙 0건)",
            ),
        )
        run_id = int(cur.lastrowid)

        insert_tuples = []
        for r in verified_records:
            r.run_id = run_id
            insert_tuples.append(
                (
                    run_id,
                    r.collected_at,
                    r.collected_date,
                    r.item_key,
                    r.platform_group,
                    r.source_site,
                    r.is_geunbon,
                    r.product_id,
                    r.raw_name,
                    r.raw_category,
                    r.brand,
                    r.keeping,
                    r.standard_part,
                    r.skin_type,
                    r.sub_spec,
                    r.processing_type,
                    r.mapping_status,
                    r.unit_price_kg,
                    r.max_price_kg,
                    r.avg_price_kg,
                    r.price_norm_method,
                    r.price_display,
                    r.min_proc_date,
                    r.max_proc_date,
                    r.proc_date_range,
                    r.exp_date_range,
                    r.prod_date_bucket,
                    r.is_sold_out,
                    r.status_text,
                    r.stock_qty,
                    r.option_cnt,
                    r.product_url,
                    r.max_product_url or r.product_url,
                )
            )

        conn.executemany(
            """
            INSERT INTO product_snapshots (
                run_id, collected_at, collected_date, item_key, platform_group, source_site,
                is_geunbon, product_id, raw_name, raw_category, brand, keeping, standard_part,
                skin_type, sub_spec, processing_type, mapping_status, unit_price_kg, max_price_kg,
                avg_price_kg, price_norm_method, price_display, min_proc_date, max_proc_date,
                proc_date_range, exp_date_range, prod_date_bucket, is_sold_out, status_text,
                stock_qty, option_cnt, product_url, max_product_url
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            insert_tuples,
        )

        # 중점관리 등록 품목(tracked_items)의 product_url도 최신 검증된 상세페이지 URL로 자동 동기화
        existing_tracked_keys = {
            tr[0] for tr in conn.execute("SELECT item_key FROM tracked_items").fetchall()
        }
        if existing_tracked_keys:
            url_sync_tuples = [
                (r.product_url, collected_at, r.item_key)
                for r in verified_records
                if r.item_key in existing_tracked_keys and r.product_url
            ]
            if url_sync_tuples:
                conn.executemany(
                    "UPDATE tracked_items SET product_url = ?, updated_at = ? WHERE item_key = ?",
                    url_sync_tuples,
                )

        _seed_default_tracked_items_if_empty(conn, verified_records)

    conn.close()

    excel_path = export_cumulative_excel(run_id=run_id, db_path=db_path)

    # GitHub Pages 정적 데이터 추출 및 원격 저장소 연결 시 자동 푸시 동기화
    try:
        import build_gh_pages

        build_gh_pages.sync_static_data_to_github(run_id=run_id, collected_at=collected_at)
    except Exception as sync_exc:
        logger.warning("GitHub Pages 정적 데이터 동기화 건너뜀: %s", sync_exc)

    result = {
        "run_id": run_id,
        "collected_at": collected_at,
        "total_raw_cnt": total_raw,
        "black_excluded_cnt": total_dropped,
        "valid_sku_cnt": len(verified_records),
        "wonmul_sku_cnt": total_wonmul,
        "processed_sku_cnt": total_processed,
        "geunbon_sku_cnt": geunbon_cnt,
        "meatbox_cnt": len(mb_records),
        "ekcm_cnt": len(ek_records),
        "ekcm_url_verified_cnt": ek_stats.get("url_verified", 0),
        "meatfriends_cnt": len(mf_records),
        "topmeat_cnt": len(tm_records),
        "slack_sent_cnt": 0,
        "excel_path": str(excel_path),
    }
    logger.info("크롤링 및 DB 누적 저장 완료: %s", json.dumps(result, ensure_ascii=False))
    return result


# ---------------------------------------------------------------------------
# FastAPI 대시보드용 통합 조회 페이로드 생성
# ---------------------------------------------------------------------------
def get_dashboard_payload(run_id: Optional[int] = None, db_path: Path = DB_PATH) -> Dict[str, Any]:
    init_db(db_path)
    conn = get_db_connection(db_path)

    runs = [dict(r) for r in conn.execute("SELECT * FROM crawl_runs ORDER BY run_id DESC").fetchall()]
    if not runs:
        conn.close()
        return {"has_data": False, "runs": []}

    if run_id is None:
        current_run = runs[0]
    else:
        matched = [r for r in runs if int(r["run_id"]) == int(run_id)]
        current_run = matched[0] if matched else runs[0]

    target_run_id = int(current_run["run_id"])

    # 중점관리 체크된 item_key 집합
    tracked_map: Dict[str, Dict[str, Any]] = {}
    for tr in conn.execute("SELECT * FROM tracked_items WHERE is_active = 1").fetchall():
        tracked_map[tr["item_key"]] = dict(tr)

    snap_rows: List[Dict[str, Any]] = []
    for r in conn.execute(
        "SELECT * FROM product_snapshots WHERE run_id = ? ORDER BY standard_part, keeping, unit_price_kg",
        (target_run_id,),
    ).fetchall():
        d = dict(r)
        d["max_product_url"] = d.get("max_product_url") or d["product_url"]
        d["sex_type"] = PorkStandardizer.get_row_sex_type(d)
        d["low_stock_limit"] = get_low_stock_threshold(d["keeping"])
        d["is_low_stock"] = 1 if is_item_low_stock(d["is_sold_out"], d["stock_qty"], d["keeping"]) else 0
        d["is_caution_stock"] = 1 if (d["is_sold_out"] or d["is_low_stock"]) else 0
        d["is_tracked"] = 1 if d["item_key"] in tracked_map else 0
        d["tracked_memo"] = tracked_map[d["item_key"]]["memo"] if d["item_key"] in tracked_map else ""
        snap_rows.append(d)

    # 중점관리 상품들의 시계열 이력 조회
    tracked_history: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    if tracked_map:
        placeholders = ",".join("?" for _ in tracked_map.keys())
        q = f"""
            SELECT item_key, run_id, collected_at, collected_date, unit_price_kg, max_price_kg,
                   avg_price_kg, stock_qty, is_sold_out, min_proc_date
            FROM product_snapshots
            WHERE item_key IN ({placeholders})
            ORDER BY run_id ASC
        """
        for hr in conn.execute(q, list(tracked_map.keys())).fetchall():
            tracked_history[hr["item_key"]].append(dict(hr))

    conn.close()

    matrix = build_comparison_matrix_from_rows(snap_rows)
    prod_date_analysis = build_production_date_analysis(snap_rows)
    alerts = build_problem_solving_alerts(matrix, snap_rows, prod_date_analysis)

    # 플랫폼별 요약 통계
    platform_summary: List[Dict[str, Any]] = []
    for grp in PLATFORM_GROUPS_ORDER:
        g_rows = [r for r in snap_rows if r["platform_group"] == grp]
        g_avail = [r for r in g_rows if not r["is_sold_out"] and r["unit_price_kg"] > 0]
        g_wonmul = [r for r in g_avail if r["processing_type"] == "원물/박스육"]
        win_cnt = sum(1 for m in matrix if m["overall_winner_group"] == grp)
        platform_summary.append(
            {
                "platform_group": grp,
                "total_sku": len(g_rows),
                "avail_sku": len(g_avail),
                "sold_out_sku": len(g_rows) - len(g_avail),
                "wonmul_sku": len(g_wonmul),
                "total_stock_qty": sum(int(r["stock_qty"] or 0) for r in g_avail),
                "avg_wonmul_price": int(round(sum(r["unit_price_kg"] for r in g_wonmul) / len(g_wonmul)))
                if g_wonmul
                else 0,
                "lowest_win_cnt": win_cnt,
            }
        )

    # 브랜드 목록 및 부위 목록 (필터용)
    brands_counter: Dict[str, int] = defaultdict(int)
    for r in snap_rows:
        brands_counter[r["brand"]] += 1
    top_brands = [b for b, _ in sorted(brands_counter.items(), key=lambda x: (-x[1], x[0]))]

    # 중점관리 체크 상품 상세 리스트 (시계열 포함)
    tracked_items_list: List[Dict[str, Any]] = []
    latest_by_key = {r["item_key"]: r for r in snap_rows}
    for ikey, t_meta in tracked_map.items():
        cur_snap = latest_by_key.get(ikey)
        hist = tracked_history.get(ikey, [])
        prev_price = hist[-2]["unit_price_kg"] if len(hist) >= 2 else (cur_snap["unit_price_kg"] if cur_snap else 0)
        cur_price = cur_snap["unit_price_kg"] if cur_snap else (hist[-1]["unit_price_kg"] if hist else 0)
        price_delta = cur_price - prev_price if (cur_price and prev_price) else 0
        tracked_items_list.append(
            {
                **t_meta,
                "sex_type": cur_snap["sex_type"]
                if cur_snap
                else PorkStandardizer.detect_sex_type(
                    t_meta.get("raw_name", ""),
                    "",
                    source_site=t_meta.get("source_site", ""),
                    product_id=t_meta.get("product_id", ""),
                ),
                "current_price_kg": cur_price,
                "max_price_kg": cur_snap["max_price_kg"] if cur_snap else cur_price,
                "price_delta": price_delta,
                "min_proc_date": cur_snap["min_proc_date"] if cur_snap else "-",
                "prod_date_bucket": cur_snap["prod_date_bucket"] if cur_snap else "-",
                "stock_qty": cur_snap["stock_qty"] if cur_snap else 0,
                "status_text": cur_snap["status_text"] if cur_snap else "정보없음",
                "is_sold_out": cur_snap["is_sold_out"] if cur_snap else 0,
                "history": hist,
            }
        )

    female_sku_cnt = sum(1 for r in snap_rows if r["sex_type"] == "암퇘지")
    normal_sku_cnt = len(snap_rows) - female_sku_cnt
    low_stock_sku_cnt = sum(1 for r in snap_rows if r["is_low_stock"] == 1)
    sold_out_sku_cnt = sum(1 for r in snap_rows if r["is_sold_out"] == 1)
    caution_stock_sku_cnt = sum(1 for r in snap_rows if r["is_caution_stock"] == 1)

    return {
        "has_data": True,
        "current_run": current_run,
        "runs": runs,
        "kpi": {
            "total_valid_sku": len(snap_rows),
            "normal_sku_cnt": normal_sku_cnt,
            "female_sku_cnt": female_sku_cnt,
            "low_stock_sku_cnt": low_stock_sku_cnt,
            "sold_out_sku_cnt": sold_out_sku_cnt,
            "caution_stock_sku_cnt": caution_stock_sku_cnt,
            "total_wonmul_sku": sum(1 for r in snap_rows if r["processing_type"] == "원물/박스육"),
            "total_stock_boxes": sum(int(r["stock_qty"] or 0) for r in snap_rows if not r["is_sold_out"]),
            "black_excluded_cnt": current_run["black_excluded_cnt"],
            "geunbon_sku_cnt": current_run["geunbon_sku_cnt"],
            "red_alert_cnt": len(alerts["red_alerts"]),
            "amber_warning_cnt": len(alerts["amber_warnings"]),
            "emerald_opp_cnt": len(alerts["emerald_opportunities"]),
            "tracked_active_cnt": len(tracked_items_list),
            "dumping_illusion_cnt": len(prod_date_analysis["dumping_alerts"]),
        },
        "platform_summary": platform_summary,
        "alerts": alerts,
        "comparison_matrix": matrix,
        "prod_date_analysis": prod_date_analysis,
        "tracked_items": tracked_items_list,
        "carcass_sise": get_carcass_price_payload(db_path),
        "filter_options": {
            "platform_groups": PLATFORM_GROUPS_ORDER,
            "keepings": KEEPINGS_ORDER,
            "sex_types": SEX_TYPES_ORDER,
            "standard_parts": STANDARD_PARTS_ORDER,
            "brands": top_brands,
            "prod_buckets": [
                "당월 신상품(최근 14일)",
                "전월/일반 생산(15~45일)",
                "장기보관/기한임박(45일 경과)",
                "주문당일생산/상시",
            ],
        },
        "products": snap_rows,
    }


if __name__ == "__main__":
    trig = sys.argv[1].strip() if len(sys.argv) > 1 and sys.argv[1].strip() else "cli_init"
    res = run_full_crawl(trigger_source=trig)
    print("CRAWL_SUCCESS:", json.dumps(res, ensure_ascii=False))
