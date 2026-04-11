"""pytest conftest — mcp-server/python/ 를 sys.path 에 추가.

테스트 모듈에서 다음과 같이 import 가능:
    from hwp_analyzer.label import _normalize, _match_label
    from hwp_core._helpers import normalize_for_match
    from hwp_core.text_editing._internal import _find_heading_positions
    from pdf_clone.native import _make_paragraph, _detect_list_markers
    from pdf_clone._models import TextBlock
"""
import sys
from pathlib import Path

# mcp-server/python/ 절대 경로
PYTHON_ROOT = Path(__file__).resolve().parent.parent.parent
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))
