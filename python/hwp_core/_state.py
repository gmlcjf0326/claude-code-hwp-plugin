"""HWP Core State — module-level state management.

hwp_service.py 의 _current_doc_path 같은 전역 상태를 관리.
여러 handler 가 읽고 쓰는 공유 상태를 한 곳에서 관리.

사용 예:
    from hwp_core._state import get_current_doc_path, set_current_doc_path

    def open_document_handler(hwp, params):
        ...
        set_current_doc_path(file_path)
        return {"status": "ok"}

    def save_document_handler(hwp, params):
        path = get_current_doc_path()
        ...
"""

# Module-level state
_current_doc_path = None


def get_current_doc_path():
    """현재 열린 문서의 경로 반환. 문서가 열려 있지 않으면 None."""
    return _current_doc_path


def set_current_doc_path(path):
    """현재 문서 경로 설정. None 은 문서 닫힘 상태."""
    global _current_doc_path
    _current_doc_path = path


def clear_current_doc_path():
    """문서 경로 초기화 (close_document 후 호출)."""
    global _current_doc_path
    _current_doc_path = None
