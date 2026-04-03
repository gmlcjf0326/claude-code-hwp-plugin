"""문서/표/폰트 프리셋 정의 — 공무원/비즈니스 실무용"""

# ── 한글 폰트 라이브러리 (40+) ──

KOREAN_FONTS = {
    # 공문서 표준 (기본 설치)
    "바탕": {"latin": "Batang", "category": "serif", "gov": True},
    "바탕체": {"latin": "BatangChe", "category": "serif_mono"},
    "신명조": {"latin": "New Batang", "category": "serif", "gov": True},
    "굴림": {"latin": "Gulim", "category": "sans"},
    "굴림체": {"latin": "GulimChe", "category": "sans_mono"},
    "돋움": {"latin": "Dotum", "category": "sans"},
    "돋움체": {"latin": "DotumChe", "category": "sans_mono"},
    "궁서": {"latin": "Gungsuh", "category": "cursive"},
    "궁서체": {"latin": "GungsuhChe", "category": "cursive_mono"},
    # 맑은 고딕 (Windows Vista+)
    "맑은 고딕": {"latin": "Malgun Gothic", "category": "sans", "gov": True},
    "맑은고딕": {"alias": "맑은 고딕"},
    # 나눔 계열 (무료 배포)
    "나눔고딕": {"latin": "NanumGothic", "category": "sans"},
    "나눔명조": {"latin": "NanumMyeongjo", "category": "serif"},
    "나눔바른고딕": {"latin": "NanumBarunGothic", "category": "sans"},
    "나눔바른펜": {"latin": "NanumBarunpen", "category": "handwriting"},
    "나눔스퀘어": {"latin": "NanumSquare", "category": "sans"},
    "나눔스퀘어라운드": {"latin": "NanumSquareRound", "category": "sans"},
    "나눔스퀘어 Bold": {"latin": "NanumSquare Bold", "category": "sans_bold"},
    # 함초롬 (한컴 기본)
    "함초롬바탕": {"latin": "HCR Batang", "category": "serif"},
    "함초롬돋움": {"latin": "HCR Dotum", "category": "sans"},
    # 한컴 기본 글꼴
    "한컴바탕": {"latin": "Hancom Batang", "category": "serif"},
    "한컴돋움": {"latin": "Hancom Dotum", "category": "sans"},
    "한컴 말랑말랑": {"category": "casual"},
    # HY 계열 (한컴/윈도우 기본)
    "HY울릉도": {"category": "display"},
    "HY견고딕": {"category": "sans_bold"},
    "HY중고딕": {"category": "sans_medium"},
    "HY신명조": {"category": "serif", "gov": True},
    "HY견명조": {"category": "serif_bold"},
    "HY헤드라인M": {"category": "display"},
    # 휴먼 계열
    "휴먼명조": {"category": "serif"},
    "휴먼엑스포": {"category": "display"},
    "휴먼고딕": {"category": "sans"},
    # 영문 주요 폰트
    "Times New Roman": {"category": "serif_en", "gov": True},
    "Arial": {"category": "sans_en"},
    "Calibri": {"category": "sans_en"},
    "Georgia": {"category": "serif_en"},
    "Verdana": {"category": "sans_en"},
    "Cambria": {"category": "serif_en"},
    "Consolas": {"category": "mono_en"},
}

def resolve_font_name(name):
    """폰트 별칭 해결 + 검증. '맑은고딕' → '맑은 고딕'"""
    if name in KOREAN_FONTS:
        info = KOREAN_FONTS[name]
        if "alias" in info:
            return info["alias"]
        return name
    return name  # 미등록 폰트는 그대로 전달 (사용자 설치 폰트일 수 있음)

def get_font_list(category=None, gov_only=False):
    """카테고리별 폰트 목록 반환"""
    result = []
    for name, info in KOREAN_FONTS.items():
        if "alias" in info:
            continue
        if category and info.get("category") != category:
            continue
        if gov_only and not info.get("gov"):
            continue
        result.append({"name": name, "category": info.get("category", ""), "gov": info.get("gov", False)})
    return result


# ── 문서 프리셋 ──

DOCUMENT_PRESETS = {
    "공문서": {
        "page": {"top": 20, "bottom": 15, "left": 20, "right": 20, "header": 10, "footer": 10},
        "body": {"font": "바탕", "size": 11, "line_spacing": 180, "align": "justify"},
        "title": {"font": "맑은 고딕", "size": 16, "bold": True, "align": "center"},
        "heading1": {"font": "맑은 고딕", "size": 14, "bold": True},
        "heading2": {"font": "맑은 고딕", "size": 12, "bold": True},
        "numbering": ["1.", "가.", "1)", "가)", "(1)", "(가)", "①", "㉮"],
    },
    "사업계획서": {
        "page": {"top": 25, "bottom": 20, "left": 25, "right": 25, "header": 10, "footer": 10},
        "body": {"font": "바탕", "size": 11, "line_spacing": 180, "align": "justify"},
        "title": {"font": "맑은 고딕", "size": 22, "bold": True, "align": "center"},
        "heading1": {"font": "맑은 고딕", "size": 16, "bold": True},
        "heading2": {"font": "맑은 고딕", "size": 14, "bold": True},
        "heading3": {"font": "맑은 고딕", "size": 12, "bold": True},
        "sections": ["사업 개요", "추진 배경 및 필요성", "추진 전략", "세부 추진 계획", "추진 일정", "소요 예산", "기대 효과"],
    },
    "제안서": {
        "page": {"top": 20, "bottom": 20, "left": 25, "right": 25},
        "body": {"font": "나눔고딕", "size": 11, "line_spacing": 170, "align": "justify"},
        "title": {"font": "나눔스퀘어", "size": 20, "bold": True, "align": "center"},
        "heading1": {"font": "나눔고딕", "size": 16, "bold": True},
    },
    "보고서": {
        "page": {"top": 25, "bottom": 20, "left": 30, "right": 25},
        "body": {"font": "바탕", "size": 11, "line_spacing": 180, "align": "justify"},
        "title": {"font": "맑은 고딕", "size": 18, "bold": True, "align": "center"},
    },
    "계약서": {
        "page": {"top": 30, "bottom": 25, "left": 25, "right": 25},
        "body": {"font": "바탕", "size": 11, "line_spacing": 180, "align": "justify"},
        "title": {"font": "바탕", "size": 16, "bold": True, "align": "center"},
    },
    "동의서": {
        "page": {"top": 25, "bottom": 20, "left": 25, "right": 25},
        "body": {"font": "바탕", "size": 11, "line_spacing": 170, "align": "justify"},
        "title": {"font": "맑은 고딕", "size": 16, "bold": True, "align": "center"},
    },
}


# ── 표 스타일 프리셋 ──

TABLE_STYLES = {
    "정부표준": {
        "header_bg": "#666666", "header_color": [255, 255, 255],
        "header_bold": True, "header_align": "center",
        "border_width": 0.3, "border_color": "#333333",
        "row_alt_bg": None,
    },
    "비즈니스": {
        "header_bg": "#2C3E50", "header_color": [255, 255, 255],
        "header_bold": True, "header_align": "center",
        "border_width": 0.5, "border_color": "#BDC3C7",
        "row_alt_bg": "#F8F9FA",
    },
    "심플": {
        "header_bg": "#F0F0F0", "header_color": [0, 0, 0],
        "header_bold": True, "header_align": "center",
        "border_width": 0.3, "border_color": "#999999",
        "row_alt_bg": None,
    },
    "강조": {
        "header_bg": "#C0392B", "header_color": [255, 255, 255],
        "header_bold": True, "header_align": "center",
        "border_width": 0.5, "border_color": "#E74C3C",
        "row_alt_bg": "#FDEDEC",
    },
}


# ── 라벨 별칭 확장 (50+) ──

LABEL_ALIASES = {
    # 기관/회사 정보
    "기업명": ["회사명", "상호명", "법인명", "업체명", "사업자명", "기관명"],
    "대표자": ["대표자성명", "대표이사", "대표", "성명"],
    "사업자등록번호": ["사업자번호", "등록번호", "사업자 등록번호"],
    "법인등록번호": ["법인번호", "법인 등록번호"],
    "주소": ["사업장주소", "소재지", "본점소재지", "사업장소재지", "주소지"],
    "전화번호": ["대표전화", "전화", "TEL", "연락처", "대표전화번호"],
    "팩스": ["팩스번호", "FAX"],
    "이메일": ["EMAIL", "E-mail", "e-mail", "전자우편"],
    # 재무 정보
    "자본금": ["납입자본금", "자본금액", "자본"],
    "매출액": ["연매출", "연매출액", "매출", "수익"],
    "사업비": ["총사업비", "사업예산", "총예산"],
    "계약금액": ["계약금", "도급금액", "금액", "총액"],
    # 인사 정보
    "직위": ["직급", "직책", "보직"],
    "부서": ["소속", "소속부서", "팀"],
    "입사일": ["채용일", "근무시작일"],
    # 사업계획서
    "사업명": ["과업명", "프로젝트명", "과제명", "연구과제명"],
    "사업기간": ["과업기간", "수행기간", "연구기간", "계약기간"],
    "사업목적": ["과업목적", "연구목적", "목적"],
    "기대효과": ["예상효과", "성과", "기대성과"],
    # 문서 관리
    "문서번호": ["관리번호", "접수번호"],
    "시행일자": ["시행일", "발행일", "작성일"],
    "수신자": ["수신", "받는 곳"],
    "발신자": ["발신", "보내는 곳"],
    # 계약 관련
    "갑": ["발주자", "주문자", "의뢰인"],
    "을": ["수급자", "공급자", "수탁자"],
    "납기": ["납품일", "완료일", "인도일"],
    "하자보증기간": ["보증기간", "하자보수기간"],
}
