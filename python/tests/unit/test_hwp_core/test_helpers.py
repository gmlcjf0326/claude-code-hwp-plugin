"""Unit tests for hwp_core._helpers module.

Tests pure functions: normalize_unicode, normalize_for_match, normalize_for_display,
validate_params, validate_file_path (file existence only).

v0.7.7: 정규화 함수 도입 + v0.7.9 Phase 0-9 분할 후 회귀 검증.
"""
import os
import tempfile

import pytest

from hwp_core._helpers import (
    normalize_unicode,
    normalize_for_match,
    normalize_for_display,
    validate_params,
    validate_file_path,
)


# ───────────────────────────────────────────────────────────────────
# normalize_unicode: NFKC + fullwidth → halfwidth 괄호
# ───────────────────────────────────────────────────────────────────

class TestNormalizeUnicode:
    def test_fullwidth_parens(self):
        # （） → ()
        assert normalize_unicode("（1）") == "(1)"

    def test_fullwidth_brackets(self):
        # 【】 → []
        assert normalize_unicode("【제목】") == "[제목]"

    def test_korean_brackets(self):
        # 「」 → []
        assert normalize_unicode("「본문」") == "[본문]"

    def test_nfkc_fullwidth_digits(self):
        # 1 (fullwidth) → 1 (halfwidth)
        assert normalize_unicode("１２３") == "123"

    def test_nfkc_fullwidth_alpha(self):
        # ＡＢＣ → ABC
        assert normalize_unicode("ＡＢＣ") == "ABC"

    def test_preserves_korean(self):
        assert normalize_unicode("한글 텍스트") == "한글 텍스트"

    def test_empty(self):
        assert normalize_unicode("") == ""


# ───────────────────────────────────────────────────────────────────
# normalize_for_match: 정규화 + 공백 단일화 + strip
# ───────────────────────────────────────────────────────────────────

class TestNormalizeForMatch:
    def test_collapses_multiple_spaces(self):
        assert normalize_for_match("a  b   c") == "a b c"

    def test_strips_edges(self):
        assert normalize_for_match("  hello  ") == "hello"

    def test_combines_unicode_and_whitespace(self):
        # fullwidth + multiple spaces
        assert normalize_for_match("（１）  test") == "(1) test"

    def test_korean_heading(self):
        # 사업계획서 양식 표준 패턴
        assert normalize_for_match("(1)  산업의   특성") == "(1) 산업의 특성"

    def test_tab_to_space(self):
        assert normalize_for_match("col1\tcol2") == "col1 col2"

    def test_newline_to_space(self):
        assert normalize_for_match("line1\nline2") == "line1 line2"


# ───────────────────────────────────────────────────────────────────
# normalize_for_display: 공백만 정리 (UI 용)
# ───────────────────────────────────────────────────────────────────

class TestNormalizeForDisplay:
    def test_collapses_spaces(self):
        assert normalize_for_display("a  b") == "a b"

    def test_strips(self):
        assert normalize_for_display("  hi  ") == "hi"

    def test_does_not_change_unicode(self):
        # display 는 fullwidth 변환 안 함
        assert normalize_for_display("（１）") == "（１）"


# ───────────────────────────────────────────────────────────────────
# validate_params: 필수 파라미터 검증
# ───────────────────────────────────────────────────────────────────

class TestValidateParams:
    def test_all_present(self):
        # 누락 없으면 통과 (None 반환)
        result = validate_params(
            {"file_path": "x", "text": "y"},
            ["file_path", "text"],
            "test_method",
        )
        assert result is None

    def test_missing_one(self):
        with pytest.raises(ValueError) as exc:
            validate_params(
                {"file_path": "x"},
                ["file_path", "text"],
                "test_method",
            )
        assert "test_method" in str(exc.value)
        assert "text" in str(exc.value)

    def test_missing_multiple(self):
        with pytest.raises(ValueError) as exc:
            validate_params(
                {},
                ["a", "b", "c"],
                "method_x",
            )
        msg = str(exc.value)
        assert "a" in msg and "b" in msg and "c" in msg

    def test_no_required(self):
        # 빈 required_keys 면 항상 통과
        validate_params({}, [], "noop")

    def test_empty_value_still_present(self):
        # value 가 None/빈문자열이어도 key 존재만 체크
        validate_params({"x": None}, ["x"], "test")
        validate_params({"x": ""}, ["x"], "test")


# ───────────────────────────────────────────────────────────────────
# validate_file_path: 경로 검증 (must_exist=True 만 — 보안 + 존재)
# ───────────────────────────────────────────────────────────────────

class TestValidateFilePath:
    def test_existing_file(self, tmp_path):
        f = tmp_path / "test.txt"
        f.write_text("hello")
        result = validate_file_path(str(f), must_exist=True)
        assert os.path.isabs(result)
        assert os.path.exists(result)

    def test_missing_file_raises(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            validate_file_path(str(tmp_path / "nonexistent.txt"), must_exist=True)

    def test_must_exist_false_returns_path(self, tmp_path):
        # 저장 대상 — 파일 없어도 OK (디렉토리만 있으면 됨)
        target = tmp_path / "newfile.txt"
        result = validate_file_path(str(target), must_exist=False)
        assert os.path.isabs(result)
        # 실제 파일은 없어야 함
        assert not os.path.exists(result)

    def test_must_exist_false_missing_dir_raises(self):
        with pytest.raises(FileNotFoundError):
            validate_file_path(
                "/nonexistent_dir_xyz/file.txt",
                must_exist=False,
            )

    def test_returns_absolute_path(self, tmp_path):
        f = tmp_path / "test.txt"
        f.write_text("hi")
        # 상대 경로 → 절대 경로 변환
        cwd = os.getcwd()
        os.chdir(str(tmp_path))
        try:
            result = validate_file_path("test.txt", must_exist=True)
            assert os.path.isabs(result)
        finally:
            os.chdir(cwd)
