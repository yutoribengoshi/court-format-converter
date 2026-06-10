"""半角→全角変換 to_zenkaku のテスト。

裁判所提出書面では数字・英字・カタカナ・記号を全角に統一する。
半角小書きカナの欠落は提出書面の文字化けに直結するため重点的に固定する。
"""

import pytest

from court_format_converter import to_zenkaku


@pytest.mark.parametrize("src,expected", [
    # 数字
    ("平成30年5月", "平成３０年５月"),
    ("123", "１２３"),
    # 英字
    ("abcXYZ", "ａｂｃＸＹＺ"),
    # 記号
    ("(1)", "（１）"),
    ("50%", "５０％"),
    ("A/B-C", "Ａ／Ｂ−Ｃ"),
    # 半角カタカナ（清音・濁音・半濁音）
    ("ﾃﾞｰﾀ", "データ"),
    ("ﾊﾟｿｺﾝ", "パソコン"),
    ("ｳｲﾝﾄﾞｳ", "ウインドウ"),
    # 半角小書きカナ・ヲ（回帰防止: 欠落すると「チｮコ」になる）
    ("ﾁｮｺ", "チョコ"),
    ("ｷｬﾝｾﾙ", "キャンセル"),
    ("ﾌｧｲﾙ", "ファイル"),
    ("ﾗｲﾌﾟﾁｮﾝ", "ライプチョン"),
    ("ｦ", "ヲ"),
    ("ｧｨｩｪｫｯｬｭｮ", "ァィゥェォッャュョ"),
])
def test_to_zenkaku(src, expected):
    assert to_zenkaku(src) == expected


@pytest.mark.parametrize("src", [
    # 既に全角のものは不変
    "既に全角４６ＡＢ（１）",
    "漢字はそのまま",
])
def test_already_zenkaku_unchanged(src):
    assert to_zenkaku(src) == src


def test_halfwidth_space_and_tab_preserved():
    # スペース・タブは組版で扱うため変換対象外（仕様の pin）
    assert to_zenkaku("半角 と全角") == "半角 と全角"
    assert to_zenkaku("タブ\there") == "タブ\tｈｅｒｅ"


def test_no_leftover_halfwidth_kana():
    # 全角化後に半角カナ（U+FF61–U+FF9F）が一切残らないこと
    src = "ｱｲｳｴｵｧｨｩｪｫｦﾝﾞﾟ"
    out = to_zenkaku(src)
    assert not any(0xFF61 <= ord(ch) <= 0xFF9F for ch in out), repr(out)
