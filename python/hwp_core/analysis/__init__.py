"""hwp_core.analysis — MCP 분석 핸들러 (sub-package).

v0.7.9 Phase 6: analysis.py (922L) → hwp_core/analysis/ 분할

sub-modules (부작용 import 로 @register 데코레이터 실행):
- metadata.py     — get_page_setup, get_cursor_context (2 handlers)
- profile.py      — extract_style_profile, extract_full_profile, extract_template_structure, snapshot_template_style (4 handlers)
- verification.py — verify_5stage, verify_layout, validate_consistency (3 handlers)
- detection.py    — detect_document_type, analyze_writing_patterns, estimate_workload, form_detect (4 handlers)

총 13 @register 핸들러.
"""
from . import metadata        # noqa: F401 - @register 실행 목적
from . import profile         # noqa: F401
from . import verification    # noqa: F401
from . import detection       # noqa: F401
from . import form_handler    # noqa: F401 — v0.7.10+ Phase 5A/5B/5F (analyze_form, detect_placeholders, mark_review_required)
