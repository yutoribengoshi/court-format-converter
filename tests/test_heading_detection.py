"""見出しレベル判定 detect_heading_level のテスト。

裁判所書式の階層 第１ / １ / (1) / ア / (ｱ) / ａ / (a) を正しく判定し、
番号だけ・本文中の数字・定型句（以上/記）を誤検出しないことを固定する。
"""

import pytest

from court_format_converter import detect_heading_level


@pytest.mark.parametrize("text,level", [
    ("第１　請求の趣旨", 1),
    ("第１０　当事者", 1),
    ("１　当事者", 2),
    ("(1)　事実経過", 3),
    ("（１）契約の成立", 3),
    ("⑴　概要", 3),          # 丸囲み数字（(1) の表記揺れ）
    ("ア　原告の主張", 4),
    ("(ｱ)　補足", 5),
    ("（ア）詳細", 5),
    ("ａ　第一", 6),
    ("(a)　例外", 7),
])
def test_detects_heading_levels(text, level):
    assert detect_heading_level(text) == level


@pytest.mark.parametrize("text", [
    "１",                  # 番号のみ（本文なし）
    "第１",                # 番号のみ
    "以上",                # 定型句
    "記",
    "別紙",
    "１０万円を支払え",      # 数字直後にテキスト（スペースなし）
    "第１回口頭弁論期日",    # 「第１回」はスペースがないので見出しでない
    "アスファルト舗装",      # カタカナ語の冒頭
    "令和６年６月１日",      # 日付
    "",                    # 空行
    "その他の主張について",  # 通常の本文
])
def test_rejects_non_headings(text):
    assert detect_heading_level(text) is None


def test_known_limitation_number_space_text():
    # 既知の制限: 「１０　名の従業員」のように数字＋全角スペース＋本文は
    # 見出し（L2）として誤検出される。挙動が変わったら気付けるよう pin。
    assert detect_heading_level("１０　名の従業員") == 2
