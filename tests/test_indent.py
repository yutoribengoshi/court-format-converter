"""インデント設定 set_indent / set_heading_indent のテスト。

全角文字単位（chars=×100, twips=×CHAR_TWIPS）で w:ind を組み立てる現行仕様を
XML レベルで固定する。書式崩れの回帰検出が目的。
"""

from docx import Document
from docx.oxml.ns import qn

from court_format_converter import (
    CHAR_TWIPS,
    set_indent,
    set_heading_indent,
)


def _ind(p):
    pPr = p._element.find(qn("w:pPr"))
    if pPr is None:
        return None
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        return None
    return {k.split("}")[-1]: v for k, v in ind.attrib.items()}


def _outline(p):
    pPr = p._element.find(qn("w:pPr"))
    if pPr is None:
        return None
    o = pPr.find(qn("w:outlineLvl"))
    return o.get(qn("w:val")) if o is not None else None


def _para(text="x"):
    return Document().add_paragraph(text)


def test_left_indent_chars_and_twips():
    p = _para()
    set_indent(p, left_chars=3)
    assert _ind(p) == {"leftChars": "300", "left": str(3 * CHAR_TWIPS)}


def test_first_line_indent():
    p = _para()
    set_indent(p, first_line_chars=1)
    assert _ind(p) == {"firstLineChars": "100", "firstLine": str(CHAR_TWIPS)}


def test_left_plus_first_line():
    p = _para()
    set_indent(p, left_chars=3, first_line_chars=1)
    assert _ind(p) == {
        "leftChars": "300", "left": str(3 * CHAR_TWIPS),
        "firstLineChars": "100", "firstLine": str(CHAR_TWIPS),
    }


def test_hanging_takes_precedence_over_first_line():
    # hanging と first_line を同時指定すると hanging が優先（firstLine は出ない）
    p = _para()
    set_indent(p, left_chars=2, first_line_chars=1, hanging_chars=2)
    got = _ind(p)
    assert got == {
        "leftChars": "200", "left": str(2 * CHAR_TWIPS),
        "hangingChars": "200", "hanging": str(2 * CHAR_TWIPS),
    }
    assert "firstLine" not in got


def test_all_zero_removes_ind():
    p = _para()
    set_indent(p, left_chars=2)       # 一度付けて
    set_indent(p)                     # 全ゼロで呼ぶと削除される
    assert _ind(p) is None


def test_heading_sets_outline_level():
    # 見出しは見た目に関わらず outlineLvl = level-1 が付く（Word目次生成用）
    for level in range(1, 8):
        p = _para("短い見出し")
        set_heading_indent(p, level)
        assert _outline(p) == str(level - 1)


def test_heading_short_title_indent():
    # 短いタイトル: level1 は ind なし、level>=2 は首行1字下げのみ
    p1 = _para("結論")
    set_heading_indent(p1, 1)
    assert _ind(p1) is None

    p2 = _para("当事者")
    set_heading_indent(p2, 2)
    assert _ind(p2) == {"firstLineChars": "100", "firstLine": str(CHAR_TWIPS)}


def test_heading_long_body_uses_hanging():
    # 本文兼用の長い見出し（本文20字超）はぶら下げインデントになる
    long_text = "あ" * 25
    p = _para(long_text)
    set_heading_indent(p, 2)
    got = _ind(p)
    assert got is not None
    assert "hanging" in got and "hangingChars" in got
