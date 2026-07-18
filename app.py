import streamlit as st
import pandas as pd
from PIL import Image
import io
import json
import os
import math
import traceback

# 1. 일러스트/PDF 변환 모듈 확인
try:
    import fitz
    HAS_FITZ = True
except ImportError:
    HAS_FITZ = False

# 1-1. 구글 시트 연동 모듈 확인 (업체 데이터를 모든 사용자와 공유하기 위함)
try:
    import gspread
    from google.oauth2.service_account import Credentials
    HAS_GSPREAD = True
except ImportError:
    HAS_GSPREAD = False

# 2. 파일 I/O 및 로그 관리 함수
DATA_FILE = "vendor_data.json"
CONFIG_FILE = "config.json"
MASTER_OPT_FILE = "master_options.json"
LOG_FILE = "error_log.txt"
GSHEET_WORKSHEET_NAME = "vendors"

def load_json(file_path, default_data):
    if os.path.exists(file_path):
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            log_error(f"{file_path} 로드 실패: {str(e)}")
            return default_data
    return default_data

def save_json(file_path, data):
    try:
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        return True
    except Exception as e:
        log_error(f"{file_path} 저장 실패: {str(e)}")
        return False

def log_error(error_msg):
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"[오류 로그] {error_msg}\n")

# 2-1. 구글 시트 연동 (업체 데이터 공유 저장소)
# secrets.toml에 gcp_service_account, SHEET_URL이 설정되어 있지 않으면
# 자동으로 로컬 JSON 파일(vendor_data.json) 저장 방식으로 대체된다.
@st.cache_resource(show_spinner=False)
def get_gsheet_worksheet():
    if not HAS_GSPREAD:
        return None
    try:
        if "gcp_service_account" not in st.secrets or "SHEET_URL" not in st.secrets:
            return None
        creds = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]),
            scopes=["https://www.googleapis.com/auth/spreadsheets"],
        )
        client = gspread.authorize(creds)
        sh = client.open_by_url(st.secrets["SHEET_URL"])
        try:
            ws = sh.worksheet(GSHEET_WORKSHEET_NAME)
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title=GSHEET_WORKSHEET_NAME, rows=1000, cols=3)
            ws.append_row(["카테고리", "업체명", "데이터"])
        return ws
    except st.errors.StreamlitSecretNotFoundError:
        return None
    except Exception as e:
        log_error(f"구글 시트 연결 실패: {str(e)}")
        return None

def _write_vendors_to_sheet(ws, data):
    rows = [["카테고리", "업체명", "데이터"]]
    for cat, vendor_list in data.items():
        for v in vendor_list:
            rows.append([cat, v.get("업체명", ""), json.dumps(v, ensure_ascii=False)])
    ws.clear()
    ws.update(rows)

def load_vendors():
    ws = get_gsheet_worksheet()
    if ws is not None:
        try:
            records = ws.get_all_records()
            data = {}
            for r in records:
                cat = str(r.get("카테고리", "")).strip()
                raw = r.get("데이터", "")
                if not cat or not raw:
                    continue
                try:
                    data.setdefault(cat, []).append(json.loads(raw))
                except Exception:
                    continue
            if not data:
                _write_vendors_to_sheet(ws, DEFAULT_VENDORS)
                return DEFAULT_VENDORS
            return data
        except Exception as e:
            log_error(f"구글 시트 업체 데이터 로드 실패: {str(e)}")
            st.warning("구글 시트 연결에 실패해 로컬 데이터를 임시로 표시한다. 오류 로그를 확인하렴.")
    return load_json(DATA_FILE, DEFAULT_VENDORS)

def save_vendors(data):
    ws = get_gsheet_worksheet()
    if ws is not None:
        try:
            _write_vendors_to_sheet(ws, data)
            return True
        except Exception as e:
            log_error(f"구글 시트 업체 데이터 저장 실패: {str(e)}")
            st.error("구글 시트 저장에 실패했다. 오류 로그를 확인하렴.")
            return False
    return save_json(DATA_FILE, data)

# 기본 마스터 옵션 풀 (스티커, 엽서, 마스킹 테이프 공용 목록)
DEFAULT_MASTER_OPTIONS = {
    "용지": ["일반아트지", "고급유포지", "투명데드롱", "모조지 220g", "크라프트지", "랑데부 240g", "특수홀로그램"],
    "접착": ["일반영구접착", "리무버블(재박리)", "강접착", "해당없음"],
    "후지": ["백색후지", "황색후지", "투명후지", "해당없음"],
    "코팅": ["무코팅", "유광코팅", "무광코팅", "스파클코팅", "단면유광"],
    "엽서용지": ["모조지 220g", "랑데부 240g", "몽블랑 240g", "아르떼 310g", "틴토레토 250g", "반포드 250g"],
    "인쇄방식": ["토너 인쇄", "인디고 인쇄", "디지털 인쇄", "옵셋 인쇄"],
    "인쇄도수": ["단면 4도", "양면 8도", "단면 1도 (흑백)", "양면 2도 (흑백)"],
    "고정사이즈": ["100x148 (표준 엽서)", "105x148 (A6)", "150x150 (정사각)", "100x100 (소형 정사각)", "140x200 (대형 엽서)"],
    "마테타입": ["일반 종이 마테", "박 마테 (금박)", "박 마테 (은박)", "다이컷 마테", "투명 PET 마테"],
    "마테가로": ["5m", "7m", "10m"],
    "마테세로": ["15mm", "20mm", "25mm", "30mm", "40mm", "50mm"],
    "마테도수": ["단면 4도", "단면 1도 (별색)", "무동판 인쇄"],
    "마테포장": ["수축 튜브 포장", "라벨스티커 스포 포장", "개별 종이갑 포장"],
    "아크릴굿즈종류": [
        "아크릴키링", "코롯토", "아크릴스티커", "스마트톡", "스탠드/디오라마",
        "포카홀더/포토액자", "아크릴쉐이커", "아크릴카라비너", "NFC/전자태그",
        "자석/뱃지/코스터/참", "문구류(집게, 볼펜 등)", "아크릴 재단"
    ]
}

# 기본 업체 데이터 (스티커, 엽서, 마스킹테이프 기본 탑재)
DEFAULT_VENDORS = {
    "스티커": [
        {
            "업체명": "우주 프린팅",
            "과금방식": "1판 자유 배치",
            "단가결정방식": "기준단가 + 옵션 추가금 합산",
            "기준단가": 50000,
            "판가로": 1000,
            "판세로": 500,
            "화이트인쇄": "지원 가능",
            "색상프로필": "CMYK + RGB 겸용",
            "반칼과금유형": "기본가에 포함",
            "반칼추가금": 0,
            "완칼과금유형": "별도 필수 과금",
            "완칼추가금": 5000,
            "최소재단기준": "가로/세로 각각 기준",
            "최소재단가로": 10,
            "최소재단세로": 10,
            "최소재단합계": 20,
            "반칼간거리": 2.0, "완칼간거리_동일색": 3.0, "완칼간거리_다른색": 5.0,
            "이미지반칼거리": 1.5, "완칼반칼거리": 2.0, "반칼최소가로": 5.0, "반칼최소세로": 5.0,
            "반칼색상": "#FF00FF", "완칼색상": "#00FFFF", "재단선마크": "+자 형", "칼선굵기": 0.25,
            "배송비": 3000, "무료배송액": 50000,
            "제공용지": ["일반아트지", "고급유포지", "투명데드롱"],
            "제공접착": ["일반영구접착", "리무버블(재박리)"],
            "제공후지": ["백색후지", "황색후지"],
            "제공코팅": ["무코팅", "유광코팅", "무광코팅"],
            "조합단가표": [
                {"용지": "일반아트지", "접착": "일반영구접착", "후지": "백색후지", "코팅": "무코팅", "최소수량": 1, "최대수량": 5, "적용값": 0},
                {"용지": "고급유포지", "접착": "리무버블(재박리)", "후지": "백색후지", "코팅": "유광코팅", "최소수량": 1, "최대수량": 10, "적용값": 8000}
            ]
        }
    ],
    "엽서": [
        {
            "업체명": "갈리프레이 인쇄",
            "상품명": "고품격 아트 엽서",
            "과금기준": "장수 제작 (수량 기준)",
            "단가결정방식": "옵션 조합별 단가 직접 설정",
            "기준단가": 0,
            "제공인쇄방식": ["인디고 인쇄", "디지털 인쇄"],
            "제공인쇄도수": ["단면 4도", "양면 8도"],
            "제공용지": ["랑데부 240g", "모조지 220g", "몽블랑 240g"],
            "사이즈모드": "고정 + 자유 겸용",
            "제공고정사이즈": ["100x148 (표준 엽서)", "105x148 (A6)"],
            "최소가로": 80, "최소세로": 80, "최대가로": 210, "최대세로": 297,
            "후가공목록": {
                "귀돌이(라운딩)": {"기본": 3000, "수량당": 10},
                "금박(유광)": {"기본": 15000, "수량당": 50},
                "형압(압인)": {"기본": 20000, "수량당": 0}
            },
            "재단비유형": "인쇄 가격에 포함",
            "재단비별도": 0,
            "편집여백플러스": 3.0,
            "안전여백마이너스": 3.0,
            "색상프로필": "CMYK 전용",
            "화이트인쇄": "지원 불가",
            "배송비": 3000,
            "무료배송액": 30000,
            "조합단가표": [
                {"용지": "랑데부 240g", "인쇄방식": "인디고 인쇄", "인쇄도수": "단면 4도", "최소수량": 50, "최대수량": 500, "적용값": 120},
                {"용지": "랑데부 240g", "인쇄방식": "인디고 인쇄", "인쇄도수": "양면 8도", "최소수량": 50, "최대수량": 500, "적용값": 180}
            ]
        }
    ],
    "마스킹 테이프": [
        {
            "업체명": "타디스 마테공방",
            "상품명": "프리미엄 한지 마테",
            "제공타입": ["일반 종이 마테", "박 마테 (금박)"],
            "제공가로": ["5m", "10m"],
            "제공세로": ["15mm", "20mm"],
            "제공도수": ["단면 4도"],
            "제공포장": ["수축 튜브 포장", "라벨스티커 스포 포장"],
            "세로편집추가": 1.5,
            "라벨포장여부": "지원 가능",
            "타입별라벨차등": "아니오",
            "라벨편집가로": 40.0, "라벨편집세로": 40.0,
            "라벨재단가로": 38.0, "라벨재단세로": 38.0,
            "라벨안전가로": 36.0, "라벨안전세로": 36.0,
            "타입별라벨설정": {},
            "배송비": 3000,
            "무료배송액": 50000,
            "조합단가표": [
                {"포장방법": "수축 튜브 포장", "최소수량": 10, "최대수량": 100, "적용값": 1200},
                {"포장방법": "라벨스티커 스포 포장", "최소수량": 10, "최대수량": 100, "적용값": 1500}
            ]
        }
    ]
}

DEFAULT_CONFIG = {
    "primary_color": "#1E90FF",
    "button_style": "default",
}

# 업데이트 노트 — 새 변경사항은 위쪽(리스트 맨 앞)에 추가한다.
UPDATE_NOTES = [
    {"date": "2026-07-18", "note": "환경 설정 화면 분리, 마법사 단계에 톱니바퀴 아이콘 추가, 사용 방법·업데이트 노트 화면 신설"},
    {"date": "2026-07-17", "note": "업체 데이터 구글 시트 연동 — 여러 사용자가 등록한 업체 정보를 실시간으로 공유"},
    {"date": "2026-07-16", "note": "업체 찾기 화면을 진입화면 + 1→2→3단계 마법사 UI로 전면 개편"},
    {"date": "2026-07-15", "note": "인쇄 업체 최저가 비교 프로그램 최초 배포"},
]

if "vendors" not in st.session_state:
    st.session_state.vendors = load_vendors()
if "config" not in st.session_state:
    st.session_state.config = load_json(CONFIG_FILE, DEFAULT_CONFIG)
if "master_opts" not in st.session_state:
    st.session_state.master_opts = load_json(MASTER_OPT_FILE, DEFAULT_MASTER_OPTIONS)
if "page" not in st.session_state:
    st.session_state.page = "landing"       # landing | match | settings
if "match_step" not in st.session_state:
    st.session_state.match_step = 1
if "match_product" not in st.session_state:
    st.session_state.match_product = None

_WIDE_PAGES = ("settings", "app_settings")
st.set_page_config(page_title="주문 업체 매칭", layout="wide" if st.session_state.page in _WIDE_PAGES else "centered")

_ACCENT = st.session_state.config["primary_color"]
_RADIUS = "15px" if st.session_state.config["button_style"] == "round" else "4px"
_MAX_WIDTH = "1100px" if st.session_state.page in _WIDE_PAGES else "640px"

st.html(f"""
    <link rel="stylesheet" as="style" crossorigin
          href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/variable/pretendardvariable.css">
    <style>
    :root {{
        --ink:#0a0a0a; --paper:#fff; --line:#111; --hair:#e2e2e2;
        --mute:#8a8a8a; --tint:#f7f7f7; --radius:{_RADIUS}; --accent:{_ACCENT};
    }}
    html, body, [class*="css"] {{ font-family:'Pretendard Variable',Pretendard,system-ui,sans-serif; }}
    .block-container {{ max-width:{_MAX_WIDTH}; padding-top:3.5rem; padding-bottom:5rem; }}
    h1, h2, h3 {{ letter-spacing:-0.02em; }}

    /* ── 스텝 인디케이터 ── */
    .steps {{ display:flex; align-items:center; gap:0; margin-bottom:28px; font-variant-numeric:tabular-nums; }}
    .step {{ display:flex; align-items:center; gap:10px; flex:0 0 auto; }}
    .step .num {{
        width:26px; height:26px; border-radius:50%; border:1.5px solid var(--hair);
        display:flex; align-items:center; justify-content:center;
        font-size:13px; font-weight:700; color:var(--mute);
    }}
    .step .lbl {{ font-size:13px; font-weight:600; color:var(--mute); white-space:nowrap; }}
    .step.active .num {{ border-color:var(--ink); background:var(--ink); color:#fff; }}
    .step.active .lbl {{ color:var(--ink); }}
    .step.done .num {{ border-color:var(--ink); background:#fff; color:var(--ink); }}
    .step.done .lbl {{ color:var(--ink); }}
    .step-bar {{ flex:1; height:1.5px; background:var(--hair); margin:0 12px; min-width:14px; }}
    .step-bar.filled {{ background:var(--ink); }}

    .step-title {{ font-size:20px; font-weight:800; letter-spacing:-0.02em; margin-bottom:2px; }}

    /* ── 결과 카드 ── */
    .result-summary {{ font-size:13px; color:var(--mute); margin:4px 0 18px; padding-bottom:14px; border-bottom:1px solid var(--hair); }}
    .result-summary b {{ color:var(--ink); font-weight:700; }}
    .vendor {{ border:1px solid var(--hair); border-radius:var(--radius); padding:18px 20px; margin-bottom:12px; display:flex; gap:16px; align-items:flex-start; }}
    .vendor .rank {{ flex:0 0 auto; width:30px; height:30px; border-radius:50%; background:var(--ink); color:#fff; display:flex; align-items:center; justify-content:center; font-size:14px; font-weight:800; font-variant-numeric:tabular-nums; }}
    .vendor .rank.r2 {{ background:#555; }}
    .vendor .rank.r3 {{ background:#888; }}
    .vendor .body {{ flex:1; min-width:0; }}
    .vendor .vtop {{ display:flex; align-items:baseline; justify-content:space-between; gap:10px; margin-bottom:8px; }}
    .vendor .vname {{ font-size:16px; font-weight:700; }}
    .vendor .score {{ font-size:15px; font-weight:800; font-variant-numeric:tabular-nums; white-space:nowrap; color:var(--accent); }}
    .vendor .badges {{ display:flex; flex-wrap:wrap; gap:6px; margin-bottom:6px; }}
    .vendor .badge {{ font-size:11.5px; font-weight:600; padding:3px 8px; border-radius:3px; background:var(--tint); color:#444; border:1px solid var(--hair); }}
    .vendor .note {{ font-size:13px; color:var(--mute); line-height:1.55; }}

    /* ── 진입 화면 ── */
    .landing-brand {{ font-size:26px; font-weight:800; letter-spacing:-0.03em; margin-bottom:8px; }}
    .landing-desc {{ color:var(--mute); font-size:14px; margin-bottom:8px; line-height:1.6; }}
    .ec-num {{ font-size:12px; font-weight:700; color:var(--mute); letter-spacing:0.06em; }}
    .ec-title {{ font-size:18px; font-weight:800; letter-spacing:-0.02em; margin:4px 0 6px; }}
    .ec-desc {{ font-size:13px; color:var(--mute); line-height:1.55; margin-bottom:4px; }}

    /* ── 버튼 공통 ── */
    div[data-testid="stButton"] button {{ border-radius:var(--radius); font-weight:700; font-family:inherit; }}
    div[data-testid="stButton"] button[kind="primary"] {{ background:var(--accent); border-color:var(--accent); }}

    /* 제품 선택 카드 그리드 */
    [class*="st-key-prod_card_"] button {{ height:64px; text-align:center; }}

    /* 옵션 선택 chip (pills) */
    div[data-testid="stPills"] button {{ border-radius:var(--radius) !important; font-weight:600 !important; }}
    div[data-testid="stPills"] button[aria-checked="true"] {{
        background:var(--ink) !important; border-color:var(--ink) !important; color:#fff !important;
    }}

    /* 상단 바 아이콘 (좌: 사용 방법, 우: 환경 설정) */
    [class*="st-key-help_wrap"] button, [class*="st-key-gear_wrap"] button {{
        border-color:var(--hair); background:var(--paper); padding:0; width:38px; height:38px;
        font-size:16px; margin:0 auto;
    }}
    [class*="st-key-help_wrap"] button:hover, [class*="st-key-gear_wrap"] button:hover {{ border-color:var(--ink); }}
    [class*="st-key-gear_wrap"], [class*="st-key-help_wrap"] {{ display:flex; justify-content:center; }}

    /* ── 탭 · 익스팬더 · 표 : 모던 톤으로 통일 ── */
    button[data-baseweb="tab"] {{ font-weight:700; font-size:14px; }}
    div[data-baseweb="tab-highlight"] {{ background-color:var(--ink) !important; height:2px !important; }}
    div[data-baseweb="tab-border"] {{ background-color:var(--hair) !important; }}
    [data-testid="stExpander"] {{ border:1px solid var(--hair); border-radius:var(--radius); }}
    [data-testid="stExpander"] summary {{ font-weight:600; }}
    [data-testid="stDataFrame"], [data-testid="stDataEditor"] {{ border:1px solid var(--hair); border-radius:var(--radius); overflow:hidden; }}
    [data-testid="stMetric"], [data-testid="stVerticalBlockBorderWrapper"] {{ border-radius:var(--radius); }}
    hr {{ border-color:var(--hair); }}
    </style>
""")

def esc(s):
    return str(s if s is not None else "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def load_image(uploaded_file):
    try:
        file_name = uploaded_file.name.lower()
        if file_name.endswith((".png", ".jpg", ".jpeg")):
            return Image.open(uploaded_file)
        elif file_name.endswith((".ai", ".pdf")):
            if not HAS_FITZ:
                st.error("AI 또는 PDF 파일을 화면에 보려면 터미널에서 'python -m pip install pymupdf'를 설치해야 한다.")
                return None
            file_bytes = uploaded_file.read()
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            page = doc.load_page(0)
            pix = page.get_pixmap(dpi=150)
            return Image.open(io.BytesIO(pix.tobytes("png")))
    except Exception as e:
        log_error(f"도안 변환 오류: {str(e)}\n{traceback.format_exc()}")
        st.error("도안을 읽는 중 문제가 발생. 오류 로그를 확인.")
    return None

# ==========================================
# [화면 1] 설정 및 관리 콘솔 (Settings View)
# ==========================================
if st.session_state.page == "settings":
    col_top1, col_top2 = st.columns([8, 2])
    with col_top1:
        st.title("업체 등록 · 관리 콘솔")
    with col_top2:
        if st.button("← 처음 화면으로", use_container_width=True):
            st.session_state.page = "landing"
            st.rerun()

    if get_gsheet_worksheet() is not None:
        st.caption("🔗 구글 시트 연동 중 — 여기서 등록/수정한 업체 데이터는 모든 사용자에게 공유된다.")
    else:
        st.caption("⚠️ 구글 시트 미연동 — 업체 데이터가 이 서버의 로컬 파일에만 저장되어 다른 사용자와 공유되지 않는다. (README 참고)")

    st.markdown("---")

    tab2 = st.container()

    with tab2:
        st.subheader("상품별 인쇄 업체 맞춤 설정 및 마스터 관리")
        st.write("엽서, 스티커, 마스킹테이프 등 각 상품 군에 맞는 정밀 공정을 정의하고 업체 데이터를 제어하는 콘솔이다.")

        if get_gsheet_worksheet() is not None:
            if st.button("🔄 구글 시트에서 최신 업체 데이터 새로고침 (다른 사용자가 등록한 내용 반영)"):
                st.session_state.vendors = load_vendors()
                st.success("구글 시트의 최신 업체 데이터를 반영했다.")
                st.rerun()

        with st.expander("전체 공용 마스터 재료/공정 풀 추가 (카테고리별 목록 확장)", expanded=False):
            m_col1, m_col2 = st.columns([1, 2])
            with m_col1:
                target_cat = st.selectbox("추가할 옵션 카테고리 선택", ["마테타입", "마테가로", "마테세로", "마테도수", "마테포장", "엽서용지", "인쇄방식", "인쇄도수", "고정사이즈", "용지", "접착", "후지", "코팅"])
            with m_col2:
                new_val = st.text_input(f"신규 [{target_cat}] 항목 입력")
                if st.button(f"[{target_cat}] 마스터 목록에 추가"):
                    if new_val and new_val not in st.session_state.master_opts.get(target_cat, []):
                        if target_cat not in st.session_state.master_opts:
                            st.session_state.master_opts[target_cat] = []
                        st.session_state.master_opts[target_cat].append(new_val)
                        save_json(MASTER_OPT_FILE, st.session_state.master_opts)
                        st.success(f"[{new_val}] 항목이 성공적으로 등록되었다.")
                        st.rerun()
                    else:
                        st.warning("이미 존재하거나 입력값이 비어 있다.")
                        
            st.write("현재 등록된 마스터 풀 현황:")
            for k, v_list in st.session_state.master_opts.items():
                st.write(f"- **{k}**: {', '.join(v_list)}")

        st.markdown("---")
        
        prod_list = list(st.session_state.vendors.keys())
        p_col1, p_col2 = st.columns([3, 1])
        with p_col1:
            target_prod = st.selectbox("관리할 상품 카테고리 선택", prod_list)
        with p_col2:
            new_prod_name = st.text_input("새 상품군 추가")
            if st.button("상품군 추가") and new_prod_name:
                if new_prod_name not in st.session_state.vendors:
                    st.session_state.vendors[new_prod_name] = []
                    save_vendors(st.session_state.vendors)
                    st.rerun()

        current_vendors = st.session_state.vendors.get(target_prod, [])
        vendor_names = [v.get("업체명", f"업체_{idx}") for idx, v in enumerate(current_vendors)]
        vendor_names.append("+ 신규 업체 등록")
        
        selected_v_name = st.selectbox("설정하거나 수정할 인쇄 업체 선택", vendor_names)
        is_new = (selected_v_name == "+ 신규 업체 등록")

        # ==========================================================
        # [분기 1] 마스킹 테이프 설정 콘솔 (새로 구현된 부분)
        # ==========================================================
        if target_prod == "마스킹 테이프":
            if is_new:
                v_data = {
                    "업체명": "", "상품명": "",
                    "제공타입": [st.session_state.master_opts["마테타입"][0]],
                    "제공가로": [st.session_state.master_opts["마테가로"][0]],
                    "제공세로": [st.session_state.master_opts["마테세로"][0]],
                    "제공도수": [st.session_state.master_opts["마테도수"][0]],
                    "제공포장": [st.session_state.master_opts["마테포장"][0]],
                    "세로편집추가": 1.5,
                    "라벨포장여부": "지원 불가",
                    "타입별라벨차등": "아니오",
                    "라벨편집가로": 40.0, "라벨편집세로": 40.0,
                    "라벨재단가로": 38.0, "라벨재단세로": 38.0,
                    "라벨안전가로": 36.0, "라벨안전세로": 36.0,
                    "타입별라벨설정": {},
                    "배송비": 3000, "무료배송액": 50000, "조합단가표": []
                }
                v_index = len(current_vendors)
            else:
                v_index = vendor_names.index(selected_v_name)
                v_data = current_vendors[v_index]

            st.markdown("### 1. 마스킹 테이프 업체 및 핵심 사양 설정")
            mtc1, mtc2, mtc3 = st.columns(3)
            with mtc1:
                edit_name = st.text_input("업체명", value=v_data.get("업체명", ""))
                edit_prod = st.text_input("상품명", value=v_data.get("상품명", ""))
            with mtc2:
                edit_bleed_height_plus = st.number_input("세로 편집 여백 추가 규격 (+mm, 재단사이즈 기준)", min_value=0.0, value=v_data.get("세로편집추가", 1.5), step=0.1)
                st.caption("가로 편집 사이즈는 고정 규격으로 산출됩니다.")
            with mtc3:
                edit_ship = st.number_input("기본 배송비 (원)", min_value=0, value=v_data.get("배송비", 3000))
                edit_free_ship = st.number_input("무료 배송 조건 (원, 0=없음)", min_value=0, value=v_data.get("무료배송액", 50000))

            mtc4, mtc5 = st.columns(2)
            with mtc4:
                edit_lead_time = st.text_input("제작 기간 (예: 영업일 기준 3~5일)", value=v_data.get("제작기간", ""))
            with mtc5:
                edit_fast_ship = st.radio("빠른 배송 가능 여부", ["가능", "불가능"],
                                          index=0 if v_data.get("빠른배송가능", "불가능") == "가능" else 1, horizontal=True)

            st.markdown("### 2. 마테 타입, 규격, 인쇄도수 및 포장법 다중 지정")
            mt_sec1, mt_sec2, mt_sec3 = st.columns(3)
            with mt_sec1:
                sel_types = st.multiselect("제공 마스킹 테이프 타입 선택/추가 (복수)", st.session_state.master_opts["마테타입"], default=v_data.get("제공타입", [st.session_state.master_opts["마테타입"][0]]))
                sel_colors = st.multiselect("제공 인쇄 도수 선택/추가 (복수)", st.session_state.master_opts["마테도수"], default=v_data.get("제공도수", [st.session_state.master_opts["마테도수"][0]]))
            with mt_sec2:
                sel_widths = st.multiselect("제공 가로 규격 선택/추가 (m 단위, 복수)", st.session_state.master_opts["마테가로"], default=v_data.get("제공가로", [st.session_state.master_opts["마테가로"][0]]))
                sel_heights = st.multiselect("제공 세로 규격 선택/추가 (mm 단위, 복수)", st.session_state.master_opts["마테세로"], default=v_data.get("제공세로", [st.session_state.master_opts["마테세로"][0]]))
            with mt_sec3:
                sel_packs = st.multiselect("제공 포장 방법 선택/추가 (복수)", st.session_state.master_opts["마테포장"], default=v_data.get("제공포장", [st.session_state.master_opts["마테포장"][0]]))

            st.markdown("### 3. 포장 및 원통 라벨 스티커 정밀 설정")
            lbl_avail = st.radio("라벨 포장 지원 여부", ["지원 불가", "지원 가능"], index=0 if v_data.get("라벨포장여부", "지원 불가") == "지원 불가" else 1)
            
            type_diff = "아니오"
            lbl_edit_w, lbl_edit_h = 40.0, 40.0
            lbl_cut_w, lbl_cut_h = 38.0, 38.0
            lbl_safe_w, lbl_safe_h = 36.0, 36.0
            type_lbl_config = v_data.get("타입별라벨설정", {})

            if lbl_avail == "지원 가능":
                type_diff = st.radio("마스킹 테이프 타입에 따라 라벨 사이즈가 다릅니까?", ["아니오", "예"], index=0 if v_data.get("타입별라벨차등", "아니오") == "아니오" else 1)
                
                if type_diff == "아니오":
                    lc1, lc2, lc3 = st.columns(3)
                    with lc1:
                        lbl_edit_w = st.number_input("라벨 편집 가로 (mm)", min_value=1.0, value=v_data.get("라벨편집가로", 40.0))
                        lbl_edit_h = st.number_input("라벨 편집 세로 (mm)", min_value=1.0, value=v_data.get("라벨편집세로", 40.0))
                    with lc2:
                        lbl_cut_w = st.number_input("라벨 재단 가로 (mm)", min_value=1.0, value=v_data.get("라벨재단가로", 38.0))
                        lbl_cut_h = st.number_input("라벨 재단 세로 (mm)", min_value=1.0, value=v_data.get("라벨재단세로", 38.0))
                    with lc3:
                        lbl_safe_w = st.number_input("라벨 안전 가로 (mm)", min_value=1.0, value=v_data.get("라벨안전가로", 36.0))
                        lbl_safe_h = st.number_input("라벨 안전 세로 (mm)", min_value=1.0, value=v_data.get("라벨안전세로", 36.0))
                else:
                    st.write("각 타입별 라벨 규격을 개별로 제어합니다.")
                    for t in sel_types:
                        st.markdown(f"**[{t}] 규격 설정**")
                        exist_t = type_lbl_config.get(t, {"lbl_edit_w": 40.0, "lbl_edit_h": 40.0, "lbl_cut_w": 38.0, "lbl_cut_h": 38.0, "lbl_safe_w": 36.0, "lbl_safe_h": 36.0})
                        tc1, tc2, tc3 = st.columns(3)
                        with tc1:
                            exist_t["lbl_edit_w"] = st.number_input(f"{t} 편집 가로 (mm)", min_value=1.0, value=float(exist_t.get("lbl_edit_w", 40.0)), key=f"edit_w_{t}")
                            exist_t["lbl_edit_h"] = st.number_input(f"{t} 편집 세로 (mm)", min_value=1.0, value=float(exist_t.get("lbl_edit_h", 40.0)), key=f"edit_h_{t}")
                        with tc2:
                            exist_t["lbl_cut_w"] = st.number_input(f"{t} 재단 가로 (mm)", min_value=1.0, value=float(exist_t.get("lbl_cut_w", 38.0)), key=f"cut_w_{t}")
                            exist_t["lbl_cut_h"] = st.number_input(f"{t} 재단 세로 (mm)", min_value=1.0, value=float(exist_t.get("lbl_cut_h", 38.0)), key=f"cut_h_{t}")
                        with tc3:
                            exist_t["lbl_safe_w"] = st.number_input(f"{t} 안전 가로 (mm)", min_value=1.0, value=float(exist_t.get("lbl_safe_w", 36.0)), key=f"safe_w_{t}")
                            exist_t["lbl_safe_h"] = st.number_input(f"{t} 안전 세로 (mm)", min_value=1.0, value=float(exist_t.get("lbl_safe_h", 36.0)), key=f"safe_h_{t}")
                        type_lbl_config[t] = exist_t

            st.markdown("### 4. 포장 방식별 최소/최대 수량 단가 스프레드시트 매트릭스")
            st.write("해당 업체가 제공하는 포장 방식별 구간 요금 구조입니다. 추가, 수정, 삭제가 자유롭습니다.")
            
            matrix_data = v_data.get("조합단가표", [])
            val_col_name = "단가(원)"
            if not matrix_data and sel_packs:
                matrix_data = [{"포장방법": sel_packs[0], "최소수량": 10, "최대수량": 100, val_col_name: 1200}]
            else:
                for row in matrix_data:
                    if "적용값" in row:
                        row[val_col_name] = row.pop("적용값")
                    else:
                        for old_k in ["조합적용단가(원)", "조합추가금(원)", "옵션추가금(원)", "단가(원)"]:
                            if old_k in row and val_col_name != old_k:
                                row[val_col_name] = row.pop(old_k)

            df_matrix = pd.DataFrame(matrix_data)
            req_cols = ["포장방법", "최소수량", "최대수량", val_col_name]
            for c in req_cols:
                if c not in df_matrix.columns:
                    df_matrix[c] = 0 if "수량" in c or "원" in c else ""
            df_matrix = df_matrix[req_cols]

            col_config = {
                "포장방법": st.column_config.SelectboxColumn("포장방법 선택", options=sel_packs, required=True),
                "최소수량": st.column_config.NumberColumn("최소수량(개)", min_value=1, step=1, required=True),
                "최대수량": st.column_config.NumberColumn("최대수량(개)", min_value=1, step=1, required=True),
                val_col_name: st.column_config.NumberColumn("개당 단가(원)", min_value=0, step=10, required=True)
            }

            edited_matrix = st.data_editor(df_matrix, column_config=col_config, num_rows="dynamic", use_container_width=True, key=f"mt_matrix_editor_{selected_v_name}")

            st.markdown("---")
            btn_col1, btn_col2 = st.columns([1, 4])
            with btn_col1:
                if not is_new:
                    if st.button("이 마테 업체 삭제"):
                        del st.session_state.vendors[target_prod][v_index]
                        save_vendors(st.session_state.vendors)
                        st.success("업체가 성공적으로 삭제되었다.")
                        st.rerun()
            with btn_col2:
                if st.button("마스킹 테이프 설정 및 단가 매트릭스 최종 저장"):
                    if not edit_name.strip() or not edit_prod.strip():
                        st.warning("업체명과 상품명을 확실히 기입하렴.")
                    else:
                        records = edited_matrix.to_dict(orient="records")
                        clean_matrix = []
                        for r in records:
                            clean_matrix.append({
                                "포장방법": str(r.get("포장방법", "")),
                                "최소수량": int(r.get("최소수량", 1)),
                                "최대수량": int(r.get("최대수량", 1)),
                                "적용값": int(r.get(val_col_name, 0))
                            })

                        updated_v = {
                            "업체명": edit_name.strip(),
                            "상품명": edit_prod.strip(),
                            "제공타입": sel_types,
                            "제공가로": sel_widths,
                            "제공세로": sel_heights,
                            "제공도수": sel_colors,
                            "제공포장": sel_packs,
                            "세로편집추가": edit_bleed_height_plus,
                            "라벨포장여부": lbl_avail,
                            "타입별라벨차등": type_diff,
                            "라벨편집가로": lbl_edit_w, "라벨편집세로": lbl_edit_h,
                            "라벨재단가로": lbl_cut_w, "라벨재단세로": lbl_cut_h,
                            "라벨안전가로": lbl_safe_w, "라벨안전세로": lbl_safe_h,
                            "타입별라벨설정": type_lbl_config,
                            "배송비": edit_ship,
                            "무료배송액": edit_free_ship,
                            "제작기간": edit_lead_time.strip(),
                            "빠른배송가능": edit_fast_ship,
                            "조합단가표": clean_matrix
                        }

                        if is_new:
                            st.session_state.vendors[target_prod].append(updated_v)
                        else:
                            st.session_state.vendors[target_prod][v_index] = updated_v

                        if save_vendors(st.session_state.vendors):
                            st.success(f"[{edit_name} - {edit_prod}] 마스킹 테이프 시스템 정보가 완벽하게 저장되어 마스터 표에 등재되었다.")
                            st.rerun()

            st.markdown("---")
            st.markdown("### 5. 등록된 마스킹 테이프 업체 마스터 리스트 표")
            st.write("등록 완료되거나 새로 추가된 모든 마스킹 테이프 업체의 핵심 마스터 리스트가 한눈에 정렬된다.")
            
            mt_summary = []
            for v in current_vendors:
                lbl_desc = "지원 불가"
                if v.get("라벨포장여부", "지원 불가") == "지원 가능":
                    if v.get("타입별라벨차등", "아니오") == "예":
                        lbl_desc = "타입별 라벨 사이즈 다름"
                    else:
                        lbl_desc = f"동일 라벨 (재단 {v.get('라벨재단가로')}x{v.get('라벨재단세로')}mm)"

                mt_summary.append({
                    "업체명": v.get("업체명", ""),
                    "상품명": v.get("상품명", ""),
                    "제공타입수": f"{len(v.get('제공타입', []))}종",
                    "규격 가로": ", ".join(v.get("제공가로", [])),
                    "규격 세로": ", ".join(v.get("제공세로", [])),
                    "세로 편집 오차": f"+{v.get('세로편집추가', 1.5)}mm",
                    "제공도수": ", ".join(v.get("제공도수", [])),
                    "원통 라벨 포장": lbl_desc,
                    "배송비": f"{v.get('배송비', 0):,} 원",
                    "제작기간": v.get("제작기간", "-"), "빠른배송": v.get("빠른배송가능", "불가능")
                })
            if mt_summary:
                st.dataframe(pd.DataFrame(mt_summary), use_container_width=True)
            else:
                st.info("등록된 마스킹 테이프 업체가 없다.")

        # ==========================================================
        # [분기 2] 엽서 정밀 설정 콘솔
        # ==========================================================
        elif target_prod == "엽서":
            if is_new:
                v_data = {
                    "업체명": "", "상품명": "표준 아트 엽서", "과금기준": "장수 제작 (수량 기준)",
                    "단가결정방식": "옵션 조합별 단가 직접 설정", "기준단가": 0,
                    "제공인쇄방식": [st.session_state.master_opts["인쇄방식"][0]],
                    "제공인쇄도수": [st.session_state.master_opts["인쇄도수"][0]],
                    "제공용지": [st.session_state.master_opts["엽서용지"][0]],
                    "사이즈모드": "고정 + 자유 겸용",
                    "제공고정사이즈": [st.session_state.master_opts["고정사이즈"][0]],
                    "최소가로": 80, "최소세로": 80, "최대가로": 210, "최대세로": 297,
                    "후가공목록": {"귀돌이(라운딩)": {"기본": 3000, "수량당": 10}, "금박(유광)": {"기본": 15000, "수량당": 50}},
                    "재단비유형": "인쇄 가격에 포함", "재단비별도": 0,
                    "편집여백플러스": 3.0, "안전여백마이너스": 3.0,
                    "색상프로필": "CMYK 전용", "화이트인쇄": "지원 불가",
                    "배송비": 3000, "무료배송액": 30000, "조합단가표": []
                }
                v_index = len(current_vendors)
            else:
                v_index = vendor_names.index(selected_v_name)
                v_data = current_vendors[v_index]

            st.markdown("### 1. 엽서 업체 기본 정보 및 과금/단가 결정 기준")
            pc1, pc2, pc3, pc4 = st.columns(4)
            with pc1:
                edit_name = st.text_input("업체명", value=v_data.get("업체명", ""))
                edit_prod = st.text_input("상품명", value=v_data.get("상품명", "표준 아트 엽서"))
            with pc2:
                edit_basis = st.radio("과금 기준 선택", ["장수 제작 (수량 기준)", "판 제작 (사이즈별 수량 변동)"],
                                      index=0 if v_data.get("과금기준", "장수 제작 (수량 기준)")=="장수 제작 (수량 기준)" else 1)
                edit_pricing = st.radio("단가 결정 방식", ["옵션 조합별 단가 직접 설정", "기준단가 + 옵션 추가금 합산"],
                                        index=0 if v_data.get("단가결정방식", "옵션 조합별 단가 직접 설정")=="옵션 조합별 단가 직접 설정" else 1)
            with pc3:
                edit_base_p = st.number_input("기준 단가 (원, 합산 방식일 때 사용)", min_value=0, value=v_data.get("기준단가", 0))
                edit_ship = st.number_input("기본 배송비 (원)", min_value=0, value=v_data.get("배송비", 3000))
            with pc4:
                edit_free_ship = st.number_input("무료 배송 조건 (원, 0=없음)", min_value=0, value=v_data.get("무료배송액", 30000))
                edit_profile = st.selectbox("지원 색상 프로필", ["CMYK 전용", "CMYK + RGB 겸용"], index=0 if v_data.get("색상프로필")=="CMYK 전용" else 1)
                edit_white = st.selectbox("화이트 인쇄 호환 여부", ["지원 불가", "지원 가능"], index=0 if v_data.get("화이트인쇄")=="지원 불가" else 1)

            pc5, pc6 = st.columns(2)
            with pc5:
                edit_lead_time = st.text_input("제작 기간 (예: 영업일 기준 3~5일)", value=v_data.get("제작기간", ""))
            with pc6:
                edit_fast_ship = st.radio("빠른 배송 가능 여부", ["가능", "불가능"],
                                          index=0 if v_data.get("빠른배송가능", "불가능") == "가능" else 1, horizontal=True)

            st.markdown("### 2. 인쇄 방식, 도수 및 엽서 용지 다중 선택")
            ic1, ic2, ic3 = st.columns(3)
            with ic1:
                sel_methods = st.multiselect("제공 인쇄 방식 (복수 선택 가능)", st.session_state.master_opts["인쇄방식"], default=v_data.get("제공인쇄방식", [st.session_state.master_opts["인쇄방식"][0]]))
            with ic2:
                sel_colors = st.multiselect("제공 인쇄 도수 (복수 선택 가능)", st.session_state.master_opts["인쇄도수"], default=v_data.get("제공인쇄도수", [st.session_state.master_opts["인쇄도수"][0]]))
            with ic3:
                sel_papers = st.multiselect("제공 엽서 용지 및 g수 (복수 선택 가능)", st.session_state.master_opts["엽서용지"], default=v_data.get("제공용지", [st.session_state.master_opts["엽서용지"][0]]))

            st.markdown("### 3. 사이즈 규격 모드 및 재단 여백 공차 설정")
            sc1, sc2, sc3, sc4 = st.columns(4)
            with sc1:
                edit_size_mode = st.selectbox("사이즈 지원 모드", ["고정 사이즈 전용", "자유 사이즈 전용", "고정 + 자유 겸용"],
                                              index=["고정 사이즈 전용", "자유 사이즈 전용", "고정 + 자유 겸용"].index(v_data.get("사이즈모드", "고정 + 자유 겸용")))
            with sc2:
                edit_bleed = st.number_input("편집 사이즈 여백 (+mm, 재단사이즈 기준)", min_value=0.0, value=v_data.get("편집여백플러스", 3.0), step=0.5)
            with sc3:
                edit_safe = st.number_input("안전 여백 규격 (-mm, 재단사이즈 기준)", min_value=0.0, value=v_data.get("안전여백마이너스", 3.0), step=0.5)
            with sc4:
                edit_cut_type = st.selectbox("재단 가격 과금 방식", ["인쇄 가격에 포함", "별도 설정"], index=0 if v_data.get("재단비유형")=="인쇄 가격에 포함" else 1)
                edit_cut_fee = st.number_input("별도 재단비 (원)", min_value=0, value=v_data.get("재단비별도", 0))

            if edit_size_mode in ["고정 사이즈 전용", "고정 + 자유 겸용"]:
                sel_fixed_sizes = st.multiselect("제공 고정 사이즈 목록 선택", st.session_state.master_opts["고정사이즈"], default=v_data.get("제공고정사이즈", [st.session_state.master_opts["고정사이즈"][0]]))
            else:
                sel_fixed_sizes = []

            if edit_size_mode in ["자유 사이즈 전용", "고정 + 자유 겸용"]:
                fc1, fc2, fc3, fc4 = st.columns(4)
                with fc1: edit_min_w = st.number_input("최소 가로 규격 (mm)", min_value=10, value=v_data.get("최소가로", 80))
                with fc2: edit_min_h = st.number_input("최소 세로 규격 (mm)", min_value=10, value=v_data.get("최소세로", 80))
                with fc3: edit_max_w = st.number_input("최대 가로 규격 (mm)", min_value=10, value=v_data.get("최대가로", 210))
                with fc4: edit_max_h = st.number_input("최대 세로 규격 (mm)", min_value=10, value=v_data.get("최대세로", 297))
            else:
                edit_min_w, edit_min_h, edit_max_w, edit_max_h = 0, 0, 0, 0

            st.markdown("### 4. 엽서 후가공 옵션 및 복합 추가금 관리 (기본 세팅비 + 수량당 추가금)")
            st.caption("후가공명:기본추가금:수량당추가금 형태로 쉼표로 구분하여 입력한다. 예: 귀돌이(라운딩):3000:10, 금박(유광):15000:50, 형압:20000:0")
            
            post_dict = v_data.get("후가공목록", {"귀돌이(라운딩)": {"기본": 3000, "수량당": 10}})
            post_str_list = []
            for k, val in post_dict.items():
                if isinstance(val, dict):
                    post_str_list.append(f"{k}:{val.get('기본', 0)}:{val.get('수량당', 0)}")
                else:
                    post_str_list.append(f"{k}:{val}:0")
            post_str_init = ", ".join(post_str_list)
            
            edit_post_str = st.text_input("후가공 항목 및 2단 과금 설정", value=post_str_init)
            
            def parse_post_kv(text):
                res = {}
                if not text.strip(): return res
                for item in text.split(","):
                    parts = item.split(":")
                    if len(parts) >= 3:
                        try: res[parts[0].strip()] = {"기본": int(float(parts[1].strip())), "수량당": int(float(parts[2].strip()))}
                        except: res[parts[0].strip()] = {"기본": 0, "수량당": 0}
                    elif len(parts) == 2:
                        try: res[parts[0].strip()] = {"기본": int(float(parts[1].strip())), "수량당": 0}
                        except: res[parts[0].strip()] = {"기본": 0, "수량당": 0}
                    elif len(parts) == 1 and parts[0].strip():
                        res[parts[0].strip()] = {"기본": 0, "수량당": 0}
                return res

            st.markdown("### 5. 엽서 구간별 조합 요금 스프레드시트 매트릭스")
            val_col_name = "조합적용단가(원)" if edit_pricing=="옵션 조합별 단가 직접 설정" else "조합추가금(원)"
            
            matrix_data = v_data.get("조합단가표", [])
            if not matrix_data and sel_papers and sel_methods and sel_colors:
                matrix_data = [{
                    "용지": sel_papers[0], "인쇄방식": sel_methods[0], "인쇄도수": sel_colors[0],
                    "최소수량": 50, "최대수량": 500, val_col_name: 100
                }]
            else:
                for row in matrix_data:
                    if "적용값" in row: row[val_col_name] = row.pop("적용값")
                    else:
                        for old_k in ["조합적용단가(원)", "조합추가금(원)", "옵션추가금(원)"]:
                            if old_k in row and val_col_name != old_k:
                                row[val_col_name] = row.pop(old_k)

            df_matrix = pd.DataFrame(matrix_data)
            req_cols = ["용지", "인쇄방식", "인쇄도수", "최소수량", "최대수량", val_col_name]
            for c in req_cols:
                if c not in df_matrix.columns:
                    df_matrix[c] = 0 if "수량" in c or "원" in c else ""
            df_matrix = df_matrix[req_cols]

            col_config = {
                "용지": st.column_config.SelectboxColumn("용지 선택", options=sel_papers, required=True),
                "인쇄방식": st.column_config.SelectboxColumn("인쇄방식 선택", options=sel_methods, required=True),
                "인쇄도수": st.column_config.SelectboxColumn("인쇄도수 선택", options=sel_colors, required=True),
                "최소수량": st.column_config.NumberColumn("최소수량(장/판)", min_value=1, step=10, required=True),
                "최대수량": st.column_config.NumberColumn("최대수량(장/판)", min_value=1, step=10, required=True),
                val_col_name: st.column_config.NumberColumn(val_col_name, min_value=0, step=10, required=True)
            }

            edited_matrix = st.data_editor(df_matrix, column_config=col_config, num_rows="dynamic", use_container_width=True, key=f"post_matrix_{selected_v_name}")

            st.markdown("---")
            btn_col1, btn_col2 = st.columns([1, 4])
            with btn_col1:
                if not is_new:
                    if st.button("이 엽서 업체 삭제"):
                        del st.session_state.vendors[target_prod][v_index]
                        save_vendors(st.session_state.vendors)
                        st.success("엽서 업체가 삭제되었다.")
                        st.rerun()
            with btn_col2:
                if st.button("엽서 설정 최종 저장 및 표에 등록"):
                    if not edit_name.strip() or not edit_prod.strip():
                        st.warning("업체명과 상품명을 모두 입력해야 한다.")
                    else:
                        records = edited_matrix.to_dict(orient="records")
                        clean_matrix = []
                        for r in records:
                            clean_matrix.append({
                                "용지": str(r.get("용지", "")), "인쇄방식": str(r.get("인쇄방식", "")),
                                "인쇄도수": str(r.get("인쇄도수", "")), "최소수량": int(r.get("최소수량", 1)),
                                "최대수량": int(r.get("최대수량", 1)), "적용값": int(r.get(val_col_name, 0))
                            })

                        updated_v = {
                            "업체명": edit_name.strip(), "상품명": edit_prod.strip(), "과금기준": edit_basis,
                            "단가결정방식": edit_pricing, "기준단가": edit_base_p, "제공인쇄방식": sel_methods,
                            "제공인쇄도수": sel_colors, "제공용지": sel_papers, "사이즈모드": edit_size_mode,
                            "제공고정사이즈": sel_fixed_sizes, "최소가로": edit_min_w, "최소세로": edit_min_h,
                            "최대가로": edit_max_w, "최대세로": edit_max_h, "후가공목록": parse_post_kv(edit_post_str),
                            "재단비유형": edit_cut_type, "재단비별도": edit_cut_fee, "편집여백플러스": edit_bleed,
                            "안전여백마이너스": edit_safe, "색상프로필": edit_profile, "화이트인쇄": edit_white,
                            "배송비": edit_ship, "무료배송액": edit_free_ship,
                            "제작기간": edit_lead_time.strip(), "빠른배송가능": edit_fast_ship,
                            "조합단가표": clean_matrix
                        }

                        if is_new: st.session_state.vendors[target_prod].append(updated_v)
                        else: st.session_state.vendors[target_prod][v_index] = updated_v

                        if save_vendors(st.session_state.vendors):
                            st.success(f"[{edit_name} - {edit_prod}] 엽서 정책이 완벽하게 등록되어 하단 표에 적용되었다.")
                            st.rerun()

            st.markdown("---")
            st.markdown("### 6. 등록된 엽서 업체 마스터 리스트 표")
            st.write("설정된 엽서 업체의 인쇄 방식, 도수, 여백 공차 규격 및 과금 기준이 한눈에 정렬된다.")
            
            post_summary = []
            for v in current_vendors:
                post_summary.append({
                    "업체명": v.get("업체명", ""), "상품명": v.get("상품명", ""), "과금기준": v.get("과금기준", ""),
                    "인쇄방식": ", ".join(v.get("제공인쇄방식", [])), "인쇄도수": ", ".join(v.get("제공인쇄도수", [])),
                    "사이즈모드": v.get("사이즈모드", ""), "여백(편집+/안전-)": f"+{v.get('편집여백플러스',0)} / -{v.get('안전여백마이너스',0)} mm",
                    "재단비": v.get("재단비유형", "") if v.get("재단비유형")=="인쇄 가격에 포함" else f"별도({v.get('재단비별도',0):,}원)",
                    "후가공수": f"{len(v.get('후가공목록', {}))}개 종류", "배송비": f"{v.get('배송비', 0):,} 원",
                    "제작기간": v.get("제작기간", "-"), "빠른배송": v.get("빠른배송가능", "불가능")
                })
            if post_summary: st.dataframe(pd.DataFrame(post_summary), use_container_width=True)
            else: st.info("등록된 엽서 업체가 없다.")

        # ==========================================================
        # [분기 3] 아크릴 굿즈 설정 콘솔
        # ==========================================================
        elif target_prod == "아크릴":
            if is_new:
                v_data = {
                    "업체명": "", "상품명": "",
                    "제공굿즈종류": [st.session_state.master_opts["아크릴굿즈종류"][0]],
                    "단가결정방식": "옵션 조합별 단가 직접 설정", "기준단가": 0,
                    "배송비": 3000, "무료배송액": 50000,
                    "제작기간": "", "빠른배송가능": "불가능",
                    "조합단가표": []
                }
                v_index = len(current_vendors)
            else:
                v_index = vendor_names.index(selected_v_name)
                v_data = current_vendors[v_index]

            st.markdown("### 1. 아크릴 업체 기본 정보")
            ac1, ac2, ac3 = st.columns(3)
            with ac1:
                edit_name = st.text_input("업체명", value=v_data.get("업체명", ""))
                edit_prod = st.text_input("상품명", value=v_data.get("상품명", ""))
            with ac2:
                edit_pricing_rule = st.radio("단가 결정 방식", ["옵션 조합별 단가 직접 설정", "기준단가 + 옵션 추가금 합산"],
                                             index=0 if v_data.get("단가결정방식", "옵션 조합별 단가 직접 설정") == "옵션 조합별 단가 직접 설정" else 1)
                edit_base_p = st.number_input("기준 단가 (원, 합산 방식일 때 사용)", min_value=0, value=v_data.get("기준단가", 0))
            with ac3:
                edit_ship = st.number_input("기본 배송비 (원)", min_value=0, value=v_data.get("배송비", 3000))
                edit_free_ship = st.number_input("무료 배송 조건 (원, 0=없음)", min_value=0, value=v_data.get("무료배송액", 50000))

            ac4, ac5 = st.columns(2)
            with ac4:
                edit_lead_time = st.text_input("제작 기간 (예: 영업일 기준 3~5일)", value=v_data.get("제작기간", ""))
            with ac5:
                edit_fast_ship = st.radio("빠른 배송 가능 여부", ["가능", "불가능"],
                                          index=0 if v_data.get("빠른배송가능", "불가능") == "가능" else 1, horizontal=True)

            st.markdown("### 2. 제작 가능한 굿즈 종류 선택 (복수 선택 가능)")
            sel_goods = st.multiselect("제작 가능 굿즈 종류", st.session_state.master_opts["아크릴굿즈종류"],
                                       default=v_data.get("제공굿즈종류", [st.session_state.master_opts["아크릴굿즈종류"][0]]))

            st.markdown("### 3. 굿즈 종류별 구간 요금 스프레드시트 매트릭스")
            val_col_name = "조합적용단가(원)" if edit_pricing_rule == "옵션 조합별 단가 직접 설정" else "옵션추가금(원)"

            matrix_data = v_data.get("조합단가표", [])
            if not matrix_data and sel_goods:
                matrix_data = [{"굿즈종류": sel_goods[0], "최소수량": 1, "최대수량": 50, val_col_name: 500}]
            else:
                for row in matrix_data:
                    if "적용값" in row: row[val_col_name] = row.pop("적용값")
                    else:
                        for old_k in ["조합적용단가(원)", "옵션추가금(원)"]:
                            if old_k in row and val_col_name != old_k: row[val_col_name] = row.pop(old_k)

            df_matrix = pd.DataFrame(matrix_data)
            req_cols = ["굿즈종류", "최소수량", "최대수량", val_col_name]
            for c in req_cols:
                if c not in df_matrix.columns: df_matrix[c] = 0 if "수량" in c or "원" in c else ""
            df_matrix = df_matrix[req_cols]

            col_config = {
                "굿즈종류": st.column_config.SelectboxColumn("굿즈 종류 선택", options=sel_goods, required=True),
                "최소수량": st.column_config.NumberColumn("최소수량(개)", min_value=1, step=1, required=True),
                "최대수량": st.column_config.NumberColumn("최대수량(개)", min_value=1, step=1, required=True),
                val_col_name: st.column_config.NumberColumn(val_col_name, min_value=0, step=10, required=True)
            }

            edited_matrix = st.data_editor(df_matrix, column_config=col_config, num_rows="dynamic", use_container_width=True, key=f"acrylic_matrix_{selected_v_name}")

            st.markdown("---")
            btn_col1, btn_col2 = st.columns([1, 4])
            with btn_col1:
                if not is_new:
                    if st.button("이 아크릴 업체 삭제"):
                        del st.session_state.vendors[target_prod][v_index]
                        save_vendors(st.session_state.vendors)
                        st.success("업체가 삭제되었다.")
                        st.rerun()
            with btn_col2:
                if st.button("아크릴 업체 설정 및 단가 매트릭스 최종 저장"):
                    if not edit_name.strip():
                        st.warning("업체 이름을 반드시 입력해야 한다.")
                    else:
                        records = edited_matrix.to_dict(orient="records")
                        clean_matrix = []
                        for r in records:
                            clean_matrix.append({
                                "굿즈종류": str(r.get("굿즈종류", "")),
                                "최소수량": int(r.get("최소수량", 1)),
                                "최대수량": int(r.get("최대수량", 1)),
                                "적용값": int(r.get(val_col_name, 0))
                            })

                        updated_v = {
                            "업체명": edit_name.strip(), "상품명": edit_prod.strip(),
                            "단가결정방식": edit_pricing_rule, "기준단가": edit_base_p,
                            "제공굿즈종류": sel_goods,
                            "배송비": edit_ship, "무료배송액": edit_free_ship,
                            "제작기간": edit_lead_time.strip(), "빠른배송가능": edit_fast_ship,
                            "조합단가표": clean_matrix
                        }

                        if is_new: st.session_state.vendors[target_prod].append(updated_v)
                        else: st.session_state.vendors[target_prod][v_index] = updated_v

                        if save_vendors(st.session_state.vendors):
                            st.success(f"[{edit_name}] 아크릴 업체 설정이 저장되었다.")
                            st.rerun()

            st.markdown("---")
            st.markdown("### 4. 등록된 아크릴 업체 전체 마스터 요약 표")
            ac_summary = []
            for v in current_vendors:
                ac_summary.append({
                    "업체명": v.get("업체명", ""), "상품명": v.get("상품명", ""),
                    "제작 가능 굿즈": ", ".join(v.get("제공굿즈종류", [])),
                    "단가결정방식": v.get("단가결정방식", ""),
                    "제작기간": v.get("제작기간", "-"),
                    "빠른배송": v.get("빠른배송가능", "불가능"),
                    "배송비": f"{v.get('배송비', 0):,} 원",
                    "조합규칙수": f"{len(v.get('조합단가표', []))}개 구간"
                })
            if ac_summary: st.dataframe(pd.DataFrame(ac_summary), use_container_width=True)
            else: st.info("등록된 아크릴 업체가 없다.")

        # ==========================================================
        # [분기 4] 스티커 및 기타 일반 상품 설정 콘솔
        # ==========================================================
        else:
            if is_new:
                v_data = {
                    "업체명": "", "과금방식": "수량별 장당 단가", "단가결정방식": "기준단가 + 옵션 추가금 합산",
                    "기준단가": 100, "판가로": 1000, "판세로": 500, "화이트인쇄": "지원 불가", "색상프로필": "CMYK 전용",
                    "반칼과금유형": "기본가에 포함", "반칼추가금": 0, "완칼과금유형": "기본가에 포함", "완칼추가금": 0,
                    "최소재단기준": "가로/세로 각각 기준", "최소재단가로": 10, "최소재단세로": 10, "최소재단합계": 20,
                    "반칼간거리": 2.0, "완칼간거리_동일색": 3.0, "완칼간거리_다른색": 5.0,
                    "이미지반칼거리": 1.5, "완칼반칼거리": 2.0, "반칼최소가로": 5.0, "반칼최소세로": 5.0,
                    "반칼색상": "#FF00FF", "완칼색상": "#00FFFF", "재단선마크": "+자 형", "칼선굵기": 0.25,
                    "배송비": 3000, "무료배송액": 50000,
                    "제공용지": [st.session_state.master_opts["용지"][0]], "제공접착": [st.session_state.master_opts["접착"][0]],
                    "제공후지": [st.session_state.master_opts["후지"][0]], "제공코팅": [st.session_state.master_opts["코팅"][0]],
                    "조합단가표": []
                }
                v_index = len(current_vendors)
            else:
                v_index = vendor_names.index(selected_v_name)
                v_data = current_vendors[v_index]

            st.markdown("### 1. 업체 기본 정보, 과금 구조 및 단가 결정 방식 선택")
            col_v1, col_v2, col_v3, col_v4 = st.columns(4)
            with col_v1:
                edit_name = st.text_input("업체 이름", value=v_data.get("업체명", ""))
                edit_mode = st.radio("과금 방식", ["1판 자유 배치", "수량별 장당 단가"], index=0 if v_data.get("과금방식")=="1판 자유 배치" else 1)
            with col_v2:
                edit_pricing_rule = st.radio("단가 결정 방식 선택", ["기준단가 + 옵션 추가금 합산", "옵션 조합별 단가 직접 설정"],
                                             index=0 if v_data.get("단가결정방식", "기준단가 + 옵션 추가금 합산")=="기준단가 + 옵션 추가금 합산" else 1)
                edit_base_p = st.number_input("기준 단가 (원)", min_value=0, value=v_data.get("기준단가", 100)) if edit_pricing_rule=="기준단가 + 옵션 추가금 합산" else 0
            with col_v3:
                edit_white = st.selectbox("화이트 인쇄 가능 여부", ["지원 가능", "지원 불가"], index=0 if v_data.get("화이트인쇄")=="지원 가능" else 1)
                edit_profile = st.selectbox("지원 색상 프로필", ["CMYK 전용", "CMYK + RGB 겸용"], index=0 if v_data.get("색상프로필")=="CMYK 전용" else 1)
            with col_v4:
                edit_ship = st.number_input("기본 배송비 (원)", min_value=0, value=v_data.get("배송비", 3000))
                edit_free_ship = st.number_input("무료 배송 조건 (원, 0=없음)", min_value=0, value=v_data.get("무료배송액", 50000))

            col_v5, col_v6 = st.columns(2)
            with col_v5:
                edit_lead_time = st.text_input("제작 기간 (예: 영업일 기준 3~5일)", value=v_data.get("제작기간", ""))
            with col_v6:
                edit_fast_ship = st.radio("빠른 배송 가능 여부", ["가능", "불가능"],
                                          index=0 if v_data.get("빠른배송가능", "불가능") == "가능" else 1, horizontal=True)

            if edit_mode == "1판 자유 배치":
                sc1, sc2 = st.columns(2)
                with sc1: edit_sw = st.number_input("1판 가로 규격 (mm)", min_value=100, value=v_data.get("판가로", 1000))
                with sc2: edit_sh = st.number_input("1판 세로 규격 (mm)", min_value=100, value=v_data.get("판세로", 500))
            else: edit_sw, edit_sh = 0, 0

            st.markdown("### 2. 칼선 추가금 과금 규칙")
            kc1, kc2, kc3, kc4 = st.columns(4)
            with kc1: edit_half_rule = st.selectbox("반칼과금유형", ["기본가에 포함", "별도 필수 과금", "선택 옵션 추가금"], index=["기본가에 포함", "별도 필수 과금", "선택 옵션 추가금"].index(v_data.get("반칼과금유형", "기본가에 포함")))
            with kc2: edit_half_price = st.number_input("반칼추가금 (원)", min_value=0, value=v_data.get("반칼추가금", 0))
            with kc3: edit_full_rule = st.selectbox("완칼과금유형", ["기본가에 포함", "별도 필수 과금", "선택 옵션 추가금"], index=["기본가에 포함", "별도 필수 과금", "선택 옵션 추가금"].index(v_data.get("완칼과금유형", "기본가에 포함")))
            with kc4: edit_full_price = st.number_input("완칼추가금 (원)", min_value=0, value=v_data.get("완칼추가금", 0))

            st.markdown("### 3. 이 업체가 제공하는 재료 목록 선택")
            oc1, oc2, oc3, oc4 = st.columns(4)
            with oc1: sel_papers = st.multiselect("제공 용지 선택", st.session_state.master_opts["용지"], default=v_data.get("제공용지", [st.session_state.master_opts["용지"][0]]))
            with oc2: sel_glues = st.multiselect("제공 접착 선택", st.session_state.master_opts["접착"], default=v_data.get("제공접착", [st.session_state.master_opts["접착"][0]]))
            with oc3: sel_backs = st.multiselect("제공 후지 선택", st.session_state.master_opts["후지"], default=v_data.get("제공후지", [st.session_state.master_opts["후지"][0]]))
            with oc4: sel_coats = st.multiselect("제공 코팅 선택", st.session_state.master_opts["코팅"], default=v_data.get("제공코팅", [st.session_state.master_opts["코팅"][0]]))

            st.markdown("### 4. 구간별 조합 요금 스프레드시트 매트릭스")
            val_col_name = "옵션추가금(원)" if edit_pricing_rule == "기준단가 + 옵션 추가금 합산" else "조합적용단가(원)"

            matrix_data = v_data.get("조합단가표", [])
            if not matrix_data and sel_papers and sel_glues and sel_backs and sel_coats:
                matrix_data = [{"용지": sel_papers[0], "접착": sel_glues[0], "후지": sel_backs[0], "코팅": sel_coats[0], "최소수량": 1, "최대수량": 100, val_col_name: 0}]
            else:
                for row in matrix_data:
                    if "적용값" in row: row[val_col_name] = row.pop("적용값")
                    else:
                        for old_k in ["옵션추가금(원)", "조합적용단가(원)", "구간판작업추가금(원)", "구간장당적용단가(원)"]:
                            if old_k in row and val_col_name != old_k: row[val_col_name] = row.pop(old_k)

            df_matrix = pd.DataFrame(matrix_data)
            req_cols = ["용지", "접착", "후지", "코팅", "최소수량", "최대수량", val_col_name]
            for c in req_cols:
                if c not in df_matrix.columns: df_matrix[c] = 0 if "수량" in c or "원" in c else ""
            df_matrix = df_matrix[req_cols]

            col_config = {
                "용지": st.column_config.SelectboxColumn("용지 선택", options=sel_papers, required=True),
                "접착": st.column_config.SelectboxColumn("접착 선택", options=sel_glues, required=True),
                "후지": st.column_config.SelectboxColumn("후지 선택", options=sel_backs, required=True),
                "코팅": st.column_config.SelectboxColumn("코팅 선택", options=sel_coats, required=True),
                "최소수량": st.column_config.NumberColumn("최소수량(판/장)", min_value=1, step=1, required=True),
                "최대수량": st.column_config.NumberColumn("최대수량(판/장)", min_value=1, step=1, required=True),
                val_col_name: st.column_config.NumberColumn(val_col_name, min_value=0, step=10, required=True)
            }

            edited_matrix = st.data_editor(df_matrix, column_config=col_config, num_rows="dynamic", use_container_width=True, key=f"matrix_editor_{selected_v_name}")

            st.markdown("---")
            btn_col1, btn_col2 = st.columns([1, 4])
            with btn_col1:
                if not is_new:
                    if st.button("이 업체 삭제"):
                        del st.session_state.vendors[target_prod][v_index]
                        save_vendors(st.session_state.vendors)
                        st.success("업체가 삭제되었다.")
                        st.rerun()
            with btn_col2:
                if st.button("업체 설정 및 조합 단가 매트릭스 최종 저장"):
                    if not edit_name.strip(): st.warning("업체 이름을 반드시 입력해야 한다.")
                    else:
                        records = edited_matrix.to_dict(orient="records")
                        clean_matrix = []
                        for r in records:
                            clean_matrix.append({
                                "용지": str(r.get("용지", "")), "접착": str(r.get("접착", "")),
                                "후지": str(r.get("후지", "")), "코팅": str(r.get("코팅", "")),
                                "최소수량": int(r.get("최소수량", 1)), "최대수량": int(r.get("최대수량", 1)),
                                "적용값": int(r.get(val_col_name, 0))
                            })

                        updated_v = {
                            "업체명": edit_name.strip(), "과금방식": edit_mode, "단가결정방식": edit_pricing_rule,
                            "기준단가": edit_base_p, "판가로": edit_sw, "판세로": edit_sh, "화이트인쇄": edit_white,
                            "색상프로필": edit_profile, "반칼과금유형": edit_half_rule, "반칼추가금": edit_half_price,
                            "완칼과금유형": edit_full_rule, "완칼추가금": edit_full_price,
                            "최소재단기준": v_data.get("최소재단기준", "가로/세로 각각 기준"),
                            "최소재단가로": v_data.get("최소재단가로", 10), "최소재단세로": v_data.get("최소재단세로", 10),
                            "최소재단합계": v_data.get("최소재단합계", 20), "반칼간거리": v_data.get("반칼간거리", 2.0),
                            "완칼간거리_동일색": v_data.get("완칼간거리_동일색", 3.0), "완칼간거리_다른색": v_data.get("완칼간거리_다른색", 5.0),
                            "이미지반칼거리": v_data.get("이미지반칼거리", 1.5), "완칼반칼거리": v_data.get("완칼반칼거리", 2.0),
                            "반칼최소가로": v_data.get("반칼최소가로", 5.0), "반칼최소세로": v_data.get("반칼최소세로", 5.0),
                            "반칼색상": v_data.get("반칼색상", "#FF00FF"), "완칼색상": v_data.get("완칼색상", "#00FFFF"),
                            "재단선마크": v_data.get("재단선마크", "+자 형"), "칼선굵기": v_data.get("칼선굵기", 0.25),
                            "배송비": edit_ship, "무료배송액": edit_free_ship,
                            "제작기간": edit_lead_time.strip(), "빠른배송가능": edit_fast_ship,
                            "제공용지": sel_papers,
                            "제공접착": sel_glues, "제공후지": sel_backs, "제공코팅": sel_coats, "조합단가표": clean_matrix
                        }

                        if is_new: st.session_state.vendors[target_prod].append(updated_v)
                        else: st.session_state.vendors[target_prod][v_index] = updated_v

                        if save_vendors(st.session_state.vendors):
                            st.success(f"[{edit_name}] 업체 설정이 완벽하게 저장되었다.")
                            st.rerun()

            st.markdown("---")
            st.markdown(f"### 5. [{target_prod}] 등록 업체 전체 마스터 요약 표")
            summary_rows = []
            for v in current_vendors:
                summary_rows.append({
                    "업체명": v.get("업체명", ""), "과금방식": v.get("과금방식", ""),
                    "단가결정방식": v.get("단가결정방식", "기준단가 + 옵션 추가금 합산"),
                    "기준단가": f"{v.get('기준단가', 0):,} 원" if v.get("단가결정방식")=="기준단가 + 옵션 추가금 합산" else "직접 설정",
                    "판규격(mm)": f"{v.get('판가로',0)}x{v.get('판세로',0)}" if v.get("과금방식")=="1판 자유 배치" else "해당 없음",
                    "화이트인쇄": v.get("화이트인쇄", ""), "색상프로필": v.get("색상프로필", ""),
                    "반칼과금": f"{v.get('반칼과금유형','')} ({v.get('반칼추가금',0):,}원)",
                    "완칼과금": f"{v.get('완칼과금유형','')} ({v.get('완칼추가금',0):,}원)",
                    "배송비": f"{v.get('배송비', 0):,} 원", "조합규칙수": f"{len(v.get('조합단가표', []))}개 구간",
                    "제작기간": v.get("제작기간", "-"), "빠른배송": v.get("빠른배송가능", "불가능")
                })
            if summary_rows: st.dataframe(pd.DataFrame(summary_rows), use_container_width=True)
            else: st.info("등록된 업체가 없다.")

# ==========================================
# [화면 1-2] 환경 설정 (화면 구성 · 오류 로그)
# ==========================================
elif st.session_state.page == "app_settings":
    col_top1, col_top2 = st.columns([8, 2])
    with col_top1:
        st.title("환경 설정")
    with col_top2:
        if st.button("← 돌아가기", use_container_width=True, key="app_settings_back"):
            st.session_state.page = "match"
            st.rerun()

    st.markdown("---")

    tab1, tab2 = st.tabs(["화면 구성", "오류 로그"])

    with tab1:
        st.subheader("화면 색상 및 버튼 스타일 설정")
        st.write("업체 찾기 화면의 강조 색상과 버튼 모양을 조작 및 적용.")

        col_c1, col_c2 = st.columns(2)
        with col_c1:
            new_color = st.color_picker("버튼 및 강조 색상 선택", st.session_state.config["primary_color"])
        with col_c2:
            new_style = st.radio("버튼 모양 선택", ["default (각진 모양)", "round (둥근 모양)"],
                                 index=0 if st.session_state.config["button_style"] == "default" else 1)

        if st.button("화면 설정 저장 및 즉시 적용"):
            st.session_state.config["primary_color"] = new_color
            st.session_state.config["button_style"] = "default" if "default" in new_style else "round"

            if save_json(CONFIG_FILE, st.session_state.config):
                st.success("화면 설정이 완벽하게 저장되고 메인 시스템에 적용되었다.")
                st.rerun()

    with tab2:
        st.subheader("시스템 오류 로그 진단")
        if os.path.exists(LOG_FILE):
            with open(LOG_FILE, "r", encoding="utf-8") as f:
                logs = f.read()
            st.text_area("기록된 오류 내역", logs, height=300)
            if st.button("로그 기록 삭제"):
                os.remove(LOG_FILE)
                st.rerun()
        else:
            st.info("현재 기록된 시스템 오류가 없다.")

# ==========================================
# [화면 1-3] 사용 방법 · 업데이트 노트
# ==========================================
elif st.session_state.page == "info":
    col_top1, col_top2 = st.columns([8, 2])
    with col_top1:
        st.title("사용 방법 · 업데이트 노트")
    with col_top2:
        if st.button("← 돌아가기", use_container_width=True, key="info_back"):
            st.session_state.page = "match" if st.session_state.match_product or st.session_state.match_step != 1 else "landing"
            st.rerun()

    st.markdown("---")

    tab1, tab2 = st.tabs(["사용 방법", "업데이트 노트"])

    with tab1:
        st.markdown("""
##### 1. 업체 찾기
랜딩 화면에서 **업체 찾기**를 누르면 **제품 선택 → 옵션 선택 → 추천 결과** 3단계로 진행됩니다.
원하는 제품과 규격·수량·용지 등 옵션을 고르면, 등록된 업체 중 조건에 맞는 곳을 최종 결제액이
낮은 순으로 보여줍니다.

##### 2. 업체 등록
랜딩 화면에서 **업체 등록**을 누르면 관리 콘솔로 이동합니다. 여기서 상품 카테고리별로 업체와
단가표(용지·접착·후지·코팅 조합, 수량 구간별 가격 등)를 등록·수정할 수 있습니다.

##### 3. 여러 사람과 데이터 공유
구글 시트 연동이 되어 있으면, 관리 콘솔에서 등록·수정한 업체 데이터가 이 앱을 쓰는 모든 사용자에게
공유됩니다. 다른 사람이 등록한 최신 내용을 반영하려면 관리 콘솔 상단의 **새로고침** 버튼을 누르세요.

##### 4. 환경 설정
마법사 화면 우측 상단의 ⚙️ 아이콘을 누르면 버튼 색상 등 화면 설정과 오류 로그를 확인할 수 있습니다.
        """)

    with tab2:
        for entry in UPDATE_NOTES:
            st.markdown(f"**{entry['date']}**  \n{entry['note']}")
            st.markdown("---")

if st.session_state.page in ("settings", "app_settings", "info"):
    st.stop()

# ==========================================
# [화면 2] 업체 찾기 — 진입 화면 + 1→2→3 단계 마법사
# ==========================================
def render_top_bar(steps_current=None):
    col_l, col_c, col_r = st.columns([1, 8, 1])
    with col_l:
        with st.container(key="help_wrap"):
            if st.button("❔", key="nav_help", help="사용 방법 · 업데이트 노트"):
                st.session_state.page = "info"
                st.rerun()
    with col_c:
        if steps_current:
            render_steps(steps_current)
    with col_r:
        with st.container(key="gear_wrap"):
            if st.button("⚙️", key="nav_gear", help="환경 설정"):
                st.session_state.page = "app_settings"
                st.rerun()


def render_landing():
    render_top_bar()
    st.html("<div style='height:8px'></div>")
    st.html('<div class="landing-brand">주문 업체 매칭</div>')
    st.html(
        '<p class="landing-desc">엽서·스티커·마스킹테이프 커스텀 제작 업체를<br>'
        '조건에 맞게 찾거나, 새 업체를 등록하세요.</p>')
    st.html("<div style='height:12px'></div>")

    c1, c2 = st.columns(2)
    with c1:
        with st.container(key="entry_card_match", border=True):
            st.html(
                '<div class="ec-num">01</div>'
                '<div class="ec-title">업체 찾기</div>'
                '<div class="ec-desc">원하는 옵션을 고르면 가장 잘 맞는 업체를 최저가 순으로 보여드려요.</div>')
            if st.button("시작하기 →", key="entry_match", type="primary", use_container_width=True):
                st.session_state.page = "match"
                st.session_state.match_step = 1
                st.session_state.match_product = None
                st.rerun()
    with c2:
        with st.container(key="entry_card_admin", border=True):
            st.html(
                '<div class="ec-num">02</div>'
                '<div class="ec-title">업체 등록</div>'
                '<div class="ec-desc">새 제작 업체의 취급 제품과 옵션, 단가 정보를 등록해요.</div>')
            if st.button("등록하기 →", key="entry_admin", type="primary", use_container_width=True):
                st.session_state.page = "settings"
                st.rerun()


def render_steps(current):
    labels = ["제품 선택", "옵션 선택", "추천 결과"]
    parts = []
    for i, lbl in enumerate(labels, start=1):
        cls = "active" if i == current else ("done" if i < current else "")
        parts.append(f'<div class="step {cls}" data-step="{i}"><span class="num">{i}</span><span class="lbl">{lbl}</span></div>')
        if i < 3:
            bar_cls = "filled" if i < current else ""
            parts.append(f'<div class="step-bar {bar_cls}"></div>')
    st.html(f'<div class="steps">{"".join(parts)}</div>')


def render_step1():
    st.html('<h2 class="step-title">어떤 제품을 만드시나요?</h2>')
    st.caption("제품을 먼저 고르면, 그에 맞는 옵션을 물어봐요.")

    products = list(st.session_state.vendors.keys())
    if not products:
        st.warning("등록된 상품 카테고리가 없다. 업체 등록 화면에서 상품을 먼저 추가하렴.")
        return

    cols = st.columns(min(3, len(products)))
    for i, prod in enumerate(products):
        col = cols[i % len(cols)]
        with col:
            selected = st.session_state.match_product == prod
            if st.button(prod, key=f"prod_card_{i}", use_container_width=True,
                         type="primary" if selected else "secondary"):
                st.session_state.match_product = prod
                st.rerun()

    st.html("<div style='height:20px'></div>")
    if st.button("다음으로", key="nav_to_step2", type="primary", use_container_width=True,
                 disabled=(st.session_state.match_product is None)):
        st.session_state.match_step = 2
        st.rerun()

    st.html('<div class="admin-link">')
    if st.button("← 처음 화면으로", key="nav_landing_1"):
        st.session_state.page = "landing"
        st.rerun()
    st.html('</div>')


def _pills(label, options, key, multi=False):
    opts = options if options else ["옵션 없음"]
    return st.pills(label, opts, selection_mode="multi" if multi else "single",
                     default=([] if multi else opts[0]), key=key)


def render_step2_generic(product):
    vendors = st.session_state.vendors.get(product, [])
    avail_papers, avail_glues, avail_backs, avail_coats = set(), set(), set(), set()
    for v in vendors:
        avail_papers.update(v.get("제공용지", []))
        avail_glues.update(v.get("제공접착", []))
        avail_backs.update(v.get("제공후지", []))
        avail_coats.update(v.get("제공코팅", []))

    st.markdown("**실제 제작 규격**")
    c1, c2 = st.columns(2)
    with c1:
        st.number_input("가로 사이즈 (mm)", min_value=1, value=50, step=1, key="mc_gen_w")
    with c2:
        st.number_input("세로 사이즈 (mm)", min_value=1, value=50, step=1, key="mc_gen_h")
    st.number_input("총 제작 수량 (개/장)", min_value=1, value=500, step=50, key="mc_gen_qty")

    _pills("용지 종류", sorted(avail_papers), "mc_gen_paper")
    _pills("접착 종류", sorted(avail_glues), "mc_gen_glue")
    _pills("후지 종류", sorted(avail_backs), "mc_gen_back")
    _pills("코팅 종류", sorted(avail_coats), "mc_gen_coat")

    st.markdown("**특수 인쇄 공정 및 칼선 옵션**")
    _pills("화이트 인쇄", ["화이트 인쇄 없음", "화이트 인쇄 필요"], "mc_gen_white")
    _pills("작업 색상 프로필", ["CMYK (인쇄 표준)", "RGB (웹/모니터 표준)"], "mc_gen_profile")
    st.pills("칼선 옵션 (복수 선택)", ["반칼선 추가", "완칼선 추가"], selection_mode="multi",
             default=["반칼선 추가"], key="mc_gen_cuts")


def render_step2_postcard():
    vendors = st.session_state.vendors.get("엽서", [])
    avail_fixed, all_methods, all_colors, all_papers, all_posts = set(), set(), set(), set(), set()
    for v in vendors:
        avail_fixed.update(v.get("제공고정사이즈", []))
        all_methods.update(v.get("제공인쇄방식", []))
        all_colors.update(v.get("제공인쇄도수", []))
        all_papers.update(v.get("제공용지", []))
        all_posts.update(v.get("후가공목록", {}).keys())

    st.markdown("**사이즈 및 수량**")
    _pills("사이즈 입력 방식", ["고정 사이즈 선택", "자유 사이즈 직접 입력"], "mc_pc_sizemode")
    if st.session_state.get("mc_pc_sizemode", "고정 사이즈 선택") == "고정 사이즈 선택":
        _pills("고정 규격", sorted(avail_fixed), "mc_pc_fixed")
    else:
        c1, c2 = st.columns(2)
        with c1:
            st.number_input("재단 가로 사이즈 (mm)", min_value=10, value=100, step=1, key="mc_pc_w")
        with c2:
            st.number_input("재단 세로 사이즈 (mm)", min_value=10, value=148, step=1, key="mc_pc_h")
    st.number_input("총 제작 수량 (장)", min_value=1, value=100, step=50, key="mc_pc_qty")

    st.markdown("**인쇄 사양**")
    _pills("인쇄 방식", sorted(all_methods), "mc_pc_method")
    _pills("인쇄 도수", sorted(all_colors), "mc_pc_color")
    _pills("엽서 용지", sorted(all_papers), "mc_pc_paper")

    st.markdown("**특수 인쇄 공정 및 후가공**")
    _pills("화이트 인쇄", ["화이트 인쇄 없음", "화이트 인쇄 필요"], "mc_pc_white")
    _pills("작업 색상 프로필", ["CMYK (인쇄 표준)", "RGB (웹/모니터 표준)"], "mc_pc_profile")
    st.pills("적용할 후가공 (복수 선택)", sorted(all_posts), selection_mode="multi", default=[], key="mc_pc_posts")


def render_step2_tape():
    vendors = st.session_state.vendors.get("마스킹 테이프", [])
    all_types, all_widths, all_heights, all_colors, all_packs = set(), set(), set(), set(), set()
    for v in vendors:
        all_types.update(v.get("제공타입", []))
        all_widths.update(v.get("제공가로", []))
        all_heights.update(v.get("제공세로", []))
        all_colors.update(v.get("제공도수", []))
        all_packs.update(v.get("제공포장", []))

    st.markdown("**규격 및 수량**")
    _pills("마스킹테이프 타입", sorted(all_types), "mc_mt_type")
    c1, c2 = st.columns(2)
    with c1:
        _pills("가로 길이 (m)", sorted(all_widths), "mc_mt_w")
    with c2:
        _pills("세로 폭 (mm)", sorted(all_heights), "mc_mt_h")
    st.number_input("총 주문 수량 (개)", min_value=1, value=10, step=10, key="mc_mt_qty")

    st.markdown("**인쇄 및 포장**")
    _pills("인쇄 도수", sorted(all_colors), "mc_mt_color")
    _pills("포장 방식", sorted(all_packs), "mc_mt_pack")


def render_step2_acrylic():
    vendors = st.session_state.vendors.get("아크릴", [])
    all_goods = set()
    for v in vendors:
        all_goods.update(v.get("제공굿즈종류", []))

    st.markdown("**제작할 굿즈 종류 및 수량**")
    _pills("굿즈 종류", sorted(all_goods), "mc_ac_goods")
    st.number_input("총 제작 수량 (개)", min_value=1, value=50, step=10, key="mc_ac_qty")


def render_step2():
    product = st.session_state.match_product
    if not product:
        st.session_state.match_step = 1
        st.rerun()
        return

    st.html(f'<h2 class="step-title">{esc(product)} 옵션</h2>')
    st.caption("원하는 조건을 골라 주세요. 이 조건과 가장 잘 맞는 업체를 찾아드려요.")

    with st.expander("디자인 도안 미리보기 (선택)"):
        uploaded_file = st.file_uploader(
            "디자인 도안 파일(AI, PDF, PNG, JPG)을 선택하세요.",
            type=["ai", "pdf", "png", "jpg", "jpeg"], key="design_upload",
        )
        if uploaded_file:
            img = load_image(uploaded_file)
            if img:
                st.image(img, caption=uploaded_file.name, use_container_width=True)

    if product == "마스킹 테이프":
        render_step2_tape()
    elif product == "엽서":
        render_step2_postcard()
    elif product == "아크릴":
        render_step2_acrylic()
    else:
        render_step2_generic(product)

    st.html("<div style='height:16px'></div>")
    nc1, nc2 = st.columns(2)
    with nc1:
        if st.button("이전", key="nav_back_2", use_container_width=True):
            st.session_state.match_step = 1
            st.rerun()
    with nc2:
        if st.button("결과보기", key="nav_to_3", type="primary", use_container_width=True):
            st.session_state.match_step = 3
            st.rerun()


# ── 업체별 최종 가격 계산 (기존 계산 로직을 그대로 이식) ──

def _lead_badges(v):
    badges = []
    lead = str(v.get("제작기간", "") or "").strip()
    if lead:
        badges.append(f"제작기간 {lead}")
    if v.get("빠른배송가능", "불가능") == "가능":
        badges.append("⚡ 빠른배송 가능")
    return badges


def calc_generic_results(product, size_w, size_h, req_qty, sel_p, sel_g, sel_b, sel_c,
                          req_white, req_rgb, req_half, req_full):
    vendors = st.session_state.vendors.get(product, [])
    results = []
    for v in vendors:
        if req_white and v.get("화이트인쇄") == "지원 불가":
            continue
        if req_rgb and v.get("색상프로필") == "CMYK 전용":
            continue
        if (sel_p not in v.get("제공용지", []) or sel_g not in v.get("제공접착", [])
                or sel_b not in v.get("제공후지", []) or sel_c not in v.get("제공코팅", [])):
            continue

        cut_rule = v.get("최소재단기준", "가로/세로 각각 기준")
        if cut_rule == "가로/세로 각각 기준":
            if size_w < v.get("최소재단가로", 10) or size_h < v.get("최소재단세로", 10):
                continue
        elif cut_rule == "가로+세로 합계 기준":
            if (size_w + size_h) < v.get("최소재단합계", 20):
                continue

        mode = v.get("과금방식", "수량별 장당 단가")
        pricing_rule = v.get("단가결정방식", "기준단가 + 옵션 추가금 합산")
        base_p = v.get("기준단가", 0) if pricing_rule == "기준단가 + 옵션 추가금 합산" else 0
        matrix = v.get("조합단가표", [])

        half_fee = 0
        h_type = v.get("반칼과금유형", "기본가에 포함")
        if h_type == "별도 필수 과금":
            half_fee = v.get("반칼추가금", 0)
        elif h_type == "선택 옵션 추가금" and req_half:
            half_fee = v.get("반칼추가금", 0)

        full_fee = 0
        f_type = v.get("완칼과금유형", "기본가에 포함")
        if f_type == "별도 필수 과금":
            full_fee = v.get("완칼추가금", 0)
        elif f_type == "선택 옵션 추가금" and req_full:
            full_fee = v.get("완칼추가금", 0)

        if mode == "1판 자유 배치":
            sheet_w = v.get("판가로", 1000)
            sheet_h = v.get("판세로", 500)
            margin = v.get("완칼간거리_동일색", 3.0)
            eff_w = size_w + margin
            eff_h = size_h + margin
            fit_normal = (int(sheet_w // eff_w)) * (int(sheet_h // eff_h))
            fit_rotated = (int(sheet_w // eff_h)) * (int(sheet_h // eff_w))
            per_sheet_fit = max(fit_normal, fit_rotated)
            if per_sheet_fit == 0:
                continue
            sheets_needed = math.ceil(req_qty / per_sheet_fit)

            matched_matrix_val = 0
            matrix_found = False
            for row in matrix:
                if (row.get("용지") == sel_p and row.get("접착") == sel_g
                        and row.get("후지") == sel_b and row.get("코팅") == sel_c):
                    if row.get("최소수량", 1) <= sheets_needed <= row.get("최대수량", 999999):
                        matched_matrix_val = row.get("적용값", 0)
                        matrix_found = True
                        break

            if pricing_rule == "기준단가 + 옵션 추가금 합산":
                total_print_fee = (sheets_needed * base_p) + matched_matrix_val + half_fee + full_fee
                mode_detail = f"1판 {per_sheet_fit}개 배치 (총 {sheets_needed}판 / 기준료+추가금 {matched_matrix_val:,}원)"
            else:
                total_print_fee = (sheets_needed * matched_matrix_val) + half_fee + full_fee
                mode_detail = f"1판 {per_sheet_fit}개 배치 (총 {sheets_needed}판 / 판당 고유단가 {matched_matrix_val:,}원)"
            unit_price_calc = int(total_print_fee / req_qty)
        else:
            matched_matrix_val = 0
            matrix_found = False
            for row in matrix:
                if (row.get("용지") == sel_p and row.get("접착") == sel_g
                        and row.get("후지") == sel_b and row.get("코팅") == sel_c):
                    if row.get("최소수량", 1) <= req_qty <= row.get("최대수량", 999999):
                        matched_matrix_val = row.get("적용값", 0)
                        matrix_found = True
                        break

            if pricing_rule == "기준단가 + 옵션 추가금 합산":
                effective_unit_p = base_p + matched_matrix_val
                mode_detail = f"기준단가({base_p:,}원) + 조합추가금({matched_matrix_val:,}원) 합산"
            else:
                effective_unit_p = matched_matrix_val if matrix_found else 0
                mode_detail = f"조합 구간 고유 단가 {effective_unit_p:,}원 직접 적용"
            total_print_fee = (req_qty * effective_unit_p) + half_fee + full_fee
            unit_price_calc = effective_unit_p

        ship_fee = v.get("배송비", 0)
        if v.get("무료배송액", 0) > 0 and total_print_fee >= v.get("무료배송액", 0):
            ship_fee = 0
        final_total = total_print_fee + ship_fee

        badges = [f"장당 환산 {unit_price_calc:,}원", "무료배송" if ship_fee == 0 else f"배송비 {ship_fee:,}원"]
        if h_type != "기본가에 포함" or half_fee:
            badges.append(f"반칼 {half_fee:,}원")
        if f_type != "기본가에 포함" or full_fee:
            badges.append(f"완칼 {full_fee:,}원")
        badges.extend(_lead_badges(v))

        results.append({
            "업체명": v.get("업체명", ""),
            "최종총가격": final_total,
            "badges": badges,
            "note": f"[{pricing_rule}] {mode_detail}",
        })
    return sorted(results, key=lambda x: x["최종총가격"])


def calc_postcard_results(size_choice_mode, sel_fixed_str, target_w, target_h, total_qty,
                           sel_m, sel_c, sel_p, req_white, req_rgb, sel_post_items):
    vendors = st.session_state.vendors.get("엽서", [])
    results = []
    for v in vendors:
        if req_white and v.get("화이트인쇄") == "지원 불가":
            continue
        if req_rgb and v.get("색상프로필") == "CMYK 전용":
            continue
        if (sel_m not in v.get("제공인쇄방식", []) or sel_c not in v.get("제공인쇄도수", [])
                or sel_p not in v.get("제공용지", [])):
            continue

        s_mode = v.get("사이즈모드", "고정 + 자유 겸용")
        if size_choice_mode == "고정 사이즈 선택":
            if s_mode == "자유 사이즈 전용" or sel_fixed_str not in v.get("제공고정사이즈", []):
                continue
        else:
            if s_mode == "고정 사이즈 전용":
                continue
            if target_w < v.get("최소가로", 80) or target_h < v.get("최소세로", 80):
                continue
            if target_w > v.get("최대가로", 210) or target_h > v.get("최대세로", 297):
                continue

        pricing_rule = v.get("단가결정방식", "옵션 조합별 단가 직접 설정")
        base_p = v.get("기준단가", 0) if pricing_rule == "기준단가 + 옵션 추가금 합산" else 0
        matrix = v.get("조합단가표", [])

        matched_val = 0
        found = False
        for row in matrix:
            if row.get("용지") == sel_p and row.get("인쇄방식") == sel_m and row.get("인쇄도수") == sel_c:
                if row.get("최소수량", 1) <= total_qty <= row.get("최대수량", 999999):
                    matched_val = row.get("적용값", 0)
                    found = True
                    break

        if pricing_rule == "기준단가 + 옵션 추가금 합산":
            eff_unit = base_p + matched_val
            calc_note = f"기준료({base_p:,}원) + 조합추가금({matched_val:,}원)"
        else:
            eff_unit = matched_val if found else 0
            calc_note = f"조합 고유 단가 {eff_unit:,}원 적용"

        post_fee_total = 0
        v_posts = v.get("후가공목록", {})
        for p_item in sel_post_items:
            val = v_posts.get(p_item, {"기본": 0, "수량당": 0})
            if isinstance(val, dict):
                post_fee_total += val.get("기본", 0) + (total_qty * val.get("수량당", 0))
            else:
                post_fee_total += val

        cut_fee = v.get("재단비별도", 0) if v.get("재단비유형") == "별도 설정" else 0
        print_total = (total_qty * eff_unit) + post_fee_total + cut_fee

        ship_fee = v.get("배송비", 0)
        if v.get("무료배송액", 0) > 0 and print_total >= v.get("무료배송액", 0):
            ship_fee = 0
        final_total = print_total + ship_fee

        bleed_w = target_w + (v.get("편집여백플러스", 3.0) * 2)
        bleed_h = target_h + (v.get("편집여백플러스", 3.0) * 2)
        safe_w = target_w - (v.get("안전여백마이너스", 3.0) * 2)
        safe_h = target_h - (v.get("안전여백마이너스", 3.0) * 2)

        badges = [
            f"장당 {int(print_total / total_qty):,}원",
            "무료배송" if ship_fee == 0 else f"배송비 {ship_fee:,}원",
            f"편집 {bleed_w:.0f}x{bleed_h:.0f}mm / 안전 {safe_w:.0f}x{safe_h:.0f}mm",
        ]
        if post_fee_total:
            badges.append(f"후가공 {post_fee_total:,}원")
        if cut_fee:
            badges.append(f"재단비 {cut_fee:,}원")
        badges.extend(_lead_badges(v))

        results.append({
            "업체명": f"{v.get('업체명', '')} ({v.get('상품명', '')})",
            "최종총가격": final_total,
            "badges": badges,
            "note": f"[{v.get('과금기준', '장수 제작')}] {calc_note}",
        })
    return sorted(results, key=lambda x: x["최종총가격"])


def calc_tape_results(user_type, user_w, user_h, total_qty, user_color, user_pack):
    vendors = st.session_state.vendors.get("마스킹 테이프", [])
    results = []
    for v in vendors:
        if (user_type not in v.get("제공타입", []) or user_w not in v.get("제공가로", [])
                or user_h not in v.get("제공세로", []) or user_color not in v.get("제공도수", [])
                or user_pack not in v.get("제공포장", [])):
            continue

        matrix = v.get("조합단가표", [])
        matched_val = 0
        found = False
        for row in matrix:
            if row.get("포장방법") == user_pack:
                if row.get("최소수량", 1) <= total_qty <= row.get("최대수량", 999999):
                    matched_val = row.get("적용값", 0)
                    found = True
                    break
        if not found:
            continue

        raw_h_num = float(str(user_h).replace("mm", "").strip())
        bleed_h = raw_h_num + v.get("세로_편집_추가_mm", v.get("세로편집추가", 1.5))

        lbl_info = None
        if v.get("라벨포장여부", "지원 불가") == "지원 가능" and "라벨" in user_pack:
            if v.get("타입별라벨차등", "아니오") == "예":
                t_cfg = v.get("타입별라벨설정", {}).get(user_type, {})
                lbl_info = f"라벨 재단 {t_cfg.get('lbl_cut_w', 38.0)}x{t_cfg.get('lbl_cut_h', 38.0)}mm"
            else:
                lbl_info = (
                    f"라벨 재단 {v.get('lbl_cut_w', v.get('라벨재단가로', 38.0))}"
                    f"x{v.get('lbl_cut_h', v.get('라벨재단세로', 38.0))}mm"
                )

        unit_price = matched_val
        pure_print_total = total_qty * unit_price
        ship_fee = v.get("배송비", 0)
        if v.get("무료배송액", 0) > 0 and pure_print_total >= v.get("무료배송액", 0):
            ship_fee = 0
        final_total = pure_print_total + ship_fee

        badges = [
            f"개당 {unit_price:,}원",
            "무료배송" if ship_fee == 0 else f"배송비 {ship_fee:,}원",
            f"세로 편집폭 {bleed_h:g}mm",
        ]
        if lbl_info:
            badges.append(lbl_info)
        badges.extend(_lead_badges(v))

        results.append({
            "업체명": f"{v.get('업체명', '')} ({v.get('상품명', '')})",
            "최종총가격": final_total,
            "badges": badges,
            "note": "",
        })
    return sorted(results, key=lambda x: x["최종총가격"])


def calc_acrylic_results(goods_type, total_qty):
    vendors = st.session_state.vendors.get("아크릴", [])
    results = []
    for v in vendors:
        if goods_type not in v.get("제공굿즈종류", []):
            continue

        pricing_rule = v.get("단가결정방식", "옵션 조합별 단가 직접 설정")
        base_p = v.get("기준단가", 0) if pricing_rule == "기준단가 + 옵션 추가금 합산" else 0
        matrix = v.get("조합단가표", [])

        matched_val = 0
        found = False
        for row in matrix:
            if row.get("굿즈종류") == goods_type:
                if row.get("최소수량", 1) <= total_qty <= row.get("최대수량", 999999):
                    matched_val = row.get("적용값", 0)
                    found = True
                    break
        if not found:
            continue

        if pricing_rule == "기준단가 + 옵션 추가금 합산":
            eff_unit = base_p + matched_val
            calc_note = f"기준료({base_p:,}원) + 조합추가금({matched_val:,}원)"
        else:
            eff_unit = matched_val
            calc_note = f"조합 고유 단가 {eff_unit:,}원 적용"

        print_total = total_qty * eff_unit
        ship_fee = v.get("배송비", 0)
        if v.get("무료배송액", 0) > 0 and print_total >= v.get("무료배송액", 0):
            ship_fee = 0
        final_total = print_total + ship_fee

        badges = [f"개당 {eff_unit:,}원", "무료배송" if ship_fee == 0 else f"배송비 {ship_fee:,}원"]
        badges.extend(_lead_badges(v))

        results.append({
            "업체명": f"{v.get('업체명', '')} ({v.get('상품명', '')})",
            "최종총가격": final_total,
            "badges": badges,
            "note": calc_note,
        })
    return sorted(results, key=lambda x: x["최종총가격"])


def render_vendor_card(rank, r):
    rc = "" if rank == 1 else ("r2" if rank == 2 else "r3")
    badges_html = "".join(f'<span class="badge">{esc(b)}</span>' for b in r["badges"])
    note_html = f'<div class="note">{esc(r["note"])}</div>' if r.get("note") else ""
    st.html(f'''<div class="vendor">
      <div class="rank {rc}">{rank}</div>
      <div class="body">
        <div class="vtop">
          <div class="vname">{esc(r["업체명"])}</div>
          <div class="score">{r["최종총가격"]:,}원</div>
        </div>
        <div class="badges">{badges_html}</div>
        {note_html}
      </div>
    </div>''')


def render_step3():
    product = st.session_state.match_product
    st.html('<h2 class="step-title">추천 업체</h2>')

    if product == "마스킹 테이프":
        user_type = st.session_state.get("mc_mt_type")
        user_w = st.session_state.get("mc_mt_w")
        user_h = st.session_state.get("mc_mt_h")
        qty = st.session_state.get("mc_mt_qty", 10)
        user_color = st.session_state.get("mc_mt_color")
        user_pack = st.session_state.get("mc_mt_pack")
        summary = f"{user_type} · {user_w}x{user_h} · {qty:,}개 · {user_color} · {user_pack}"
        results = calc_tape_results(user_type, user_w, user_h, qty, user_color, user_pack)

    elif product == "엽서":
        size_choice_mode = st.session_state.get("mc_pc_sizemode", "고정 사이즈 선택")
        if size_choice_mode == "고정 사이즈 선택":
            sel_fixed_str = st.session_state.get("mc_pc_fixed")
            try:
                dims = sel_fixed_str.split(" ")[0].split("x")
                target_w, target_h = int(dims[0]), int(dims[1])
            except Exception:
                target_w, target_h = 100, 148
            size_label = sel_fixed_str or "고정 사이즈"
        else:
            sel_fixed_str = None
            target_w = st.session_state.get("mc_pc_w", 100)
            target_h = st.session_state.get("mc_pc_h", 148)
            size_label = f"{target_w}x{target_h}mm (자유)"
        qty = st.session_state.get("mc_pc_qty", 100)
        sel_m = st.session_state.get("mc_pc_method")
        sel_c = st.session_state.get("mc_pc_color")
        sel_p = st.session_state.get("mc_pc_paper")
        req_white = st.session_state.get("mc_pc_white") == "화이트 인쇄 필요"
        req_rgb = st.session_state.get("mc_pc_profile") == "RGB (웹/모니터 표준)"
        sel_posts = st.session_state.get("mc_pc_posts") or []
        summary = f"{size_label} · {qty:,}장 · {sel_m} · {sel_c} · {sel_p}"
        results = calc_postcard_results(size_choice_mode, sel_fixed_str, target_w, target_h, qty,
                                         sel_m, sel_c, sel_p, req_white, req_rgb, sel_posts)

    elif product == "아크릴":
        goods_type = st.session_state.get("mc_ac_goods")
        qty = st.session_state.get("mc_ac_qty", 50)
        summary = f"{goods_type} · {qty:,}개"
        results = calc_acrylic_results(goods_type, qty)

    else:
        size_w = st.session_state.get("mc_gen_w", 50)
        size_h = st.session_state.get("mc_gen_h", 50)
        qty = st.session_state.get("mc_gen_qty", 500)
        sel_p = st.session_state.get("mc_gen_paper")
        sel_g = st.session_state.get("mc_gen_glue")
        sel_b = st.session_state.get("mc_gen_back")
        sel_c = st.session_state.get("mc_gen_coat")
        req_white = st.session_state.get("mc_gen_white") == "화이트 인쇄 필요"
        req_rgb = st.session_state.get("mc_gen_profile") == "RGB (웹/모니터 표준)"
        cuts = st.session_state.get("mc_gen_cuts") or []
        req_half = "반칼선 추가" in cuts
        req_full = "완칼선 추가" in cuts
        summary = f"{size_w}x{size_h}mm · {qty:,}개 · {sel_p}/{sel_g}/{sel_b}/{sel_c}"
        results = calc_generic_results(product, size_w, size_h, qty, sel_p, sel_g, sel_b, sel_c,
                                        req_white, req_rgb, req_half, req_full)

    st.html(f'<div class="result-summary"><b>{esc(product)}</b> · {esc(summary)}</div>')

    if not results:
        st.html('<div class="state err">조건에 맞는 등록된 업체가 없습니다.</div>')
    else:
        for i, r in enumerate(results):
            render_vendor_card(i + 1, r)
        st.success(f"최저가 1위는 [{results[0]['업체명']}]이며, 최종 결제액은 {results[0]['최종총가격']:,}원이다.")

    st.html("<div style='height:8px'></div>")
    nc1, nc2 = st.columns(2)
    with nc1:
        if st.button("조건 수정", key="nav_back_3", use_container_width=True):
            st.session_state.match_step = 2
            st.rerun()
    with nc2:
        if st.button("처음부터", key="nav_restart", type="primary", use_container_width=True):
            st.session_state.match_step = 1
            st.session_state.match_product = None
            st.rerun()

    st.html('<div class="admin-link">')
    if st.button("← 처음 화면으로", key="nav_landing_3"):
        st.session_state.page = "landing"
        st.session_state.match_step = 1
        st.session_state.match_product = None
        st.rerun()
    st.html('</div>')


def render_match():
    render_top_bar(steps_current=st.session_state.match_step)

    if st.session_state.match_step == 1:
        render_step1()
    elif st.session_state.match_step == 2:
        render_step2()
    else:
        render_step3()


if st.session_state.page == "match":
    render_match()
else:
    render_landing()
