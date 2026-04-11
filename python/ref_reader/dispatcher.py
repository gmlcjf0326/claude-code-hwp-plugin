"""ref_reader.dispatcher — 확장자 라우터 + 볼륨 경고.

역할:
- 파일 크기 기반 volume warning 생성 (_check_volume_warning)
- 확장자로 적절한 reader 함수 선택 (read_reference)
- 모든 결과에 volume_warning 주입
"""
import os

from .readers import (
    _read_text, _read_csv, _read_excel, _read_json,
    _read_pdf, _read_html, _read_xml, _read_hwp_structured,
)
from .conversion import _read_via_pdf_conversion


def _check_volume_warning(file_path, size_bytes):
    """v0.7.4.8 Part 4: 파일 크기 기반 볼륨 가이드 warning 생성.

    Returns: dict | None — 권장 범위 초과 시 warning 객체, 아니면 None.
    """
    # 임계값 (policy 와 동기화 유지)
    OPTIMAL_KB = 60
    RECOMMEND_KB = 150
    size_kb = size_bytes / 1024
    if size_kb <= OPTIMAL_KB:
        return {
            "volume_level": "optimal",
            "icon": "🟢",
            "size_kb": round(size_kb, 1),
            "message": f"최적 크기 — LLM focus 최상 ({round(size_kb, 1)}KB ≤ {OPTIMAL_KB}KB)",
        }
    if size_kb <= RECOMMEND_KB:
        return {
            "volume_level": "adequate",
            "icon": "🟡",
            "size_kb": round(size_kb, 1),
            "message": f"적정 크기 — 실무 기본 범위 ({round(size_kb, 1)}KB ≤ {RECOMMEND_KB}KB)",
        }
    if size_kb <= 500:
        return {
            "volume_level": "warning",
            "icon": "🟠",
            "size_kb": round(size_kb, 1),
            "message": (
                f"주의: {round(size_kb, 1)}KB > {RECOMMEND_KB}KB — "
                f"focus degradation 가능. 파일 분할 또는 핵심 발췌 권장."
            ),
        }
    return {
        "volume_level": "critical",
        "icon": "🔴",
        "size_kb": round(size_kb, 1),
        "message": (
            f"경고: {round(size_kb, 1)}KB > 500KB — "
            f"품질 저하 가능성 높음. split 또는 요약 필수."
        ),
    }


def read_reference(file_path, max_chars=30000, hwp=None):
    """참고자료 파일에서 텍스트 추출.

    Args:
        file_path: 파일 경로
        max_chars: 최대 추출 글자수 (기본 30,000). 초과 시 truncated 플래그 세팅
        hwp: .hwp/.hwpx 파일 읽을 때 필요한 pyhwpx Hwp() 인스턴스 (v0.7.4.8 신규)

    Returns:
        dict — format 별 다름. 공통 필드:
          format, file_name, char_count, truncated, original_char_count,
          volume_warning (v0.7.4.8 신규 — 150KB 초과 시 주의 메시지)
    """
    file_path = os.path.abspath(file_path)
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")

    # v0.7.4.8 Part 4: 볼륨 가이드 warning 사전 생성
    try:
        size_bytes = os.path.getsize(file_path)
        volume_warning = _check_volume_warning(file_path, size_bytes)
    except Exception:
        volume_warning = None

    ext = os.path.splitext(file_path)[1].lower()

    # 모든 return 에 volume_warning 주입 (wrapper)
    def _merge_warning(result):
        if isinstance(result, dict) and volume_warning:
            result["volume_warning"] = volume_warning
        return result

    if ext in ('.txt', '.md', '.log'):
        return _merge_warning(_read_text(file_path, max_chars))
    elif ext == '.csv':
        return _merge_warning(_read_csv(file_path, max_chars))
    elif ext in ('.xlsx', '.xls'):
        return _merge_warning(_read_excel(file_path, max_chars))
    elif ext == '.json':
        return _merge_warning(_read_json(file_path, max_chars))
    elif ext == '.pdf':
        return _merge_warning(_read_pdf(file_path, max_chars))
    elif ext in ('.html', '.htm'):
        return _merge_warning(_read_html(file_path, max_chars))
    elif ext == '.xml':
        return _merge_warning(_read_xml(file_path, max_chars))
    # v0.7.4.8 Fix C1/C2: .hwp/.hwpx 는 hwp_analyzer.analyze_document 로 structured extraction
    elif ext in ('.hwp', '.hwpx'):
        if hwp is None:
            # Fallback: dispatch handler 가 hwp 인스턴스를 넘기지 않은 경우 PDF 변환 fallback
            return _merge_warning(_read_via_pdf_conversion(file_path, max_chars))
        return _merge_warning(_read_hwp_structured(hwp, file_path, max_chars))
    elif ext in ('.docx', '.doc', '.pptx', '.ppt', '.rtf', '.odt', '.odp'):
        return _merge_warning(_read_via_pdf_conversion(file_path, max_chars))
    else:
        raise ValueError(
            f"지원하지 않는 파일 형식: {ext}. "
            f"지원: .txt, .md, .csv, .xlsx, .json, .pdf, .html, .xml, .hwp, .hwpx, .docx, .pptx"
        )
