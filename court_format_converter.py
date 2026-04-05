#!/usr/bin/env python3
"""
court_format_converter.py — 裁判所提出書面の書式整形ツール
=============================================================
文化審議会建議「公用文作成の考え方」（令和4年）および裁判所実務の
書式慣行に準拠し、docxファイルのページ設定・フォント・インデント・
全角変換を一括で整形します。

使い方:
    python3 court_format_converter.py input.docx [output.docx]
    出力ファイル未指定時は「<元ファイル名>_裁判所書式.docx」に保存。

変換内容:
    - ページ設定: A4、余白(上35/下25/左30/右20mm)、26行x37文字グリッド
    - フォント: MS 明朝 / Times New Roman 12pt に統一
    - 見出し: テキストパターンから自動判定しインデント適用
    - テーブル: フォント統一＋レイアウト調整
    - フッター: ページ番号（中央）
    - 半角→全角変換: 数字・英字・カタカナ・括弧・記号を全角に統一

ライセンス: MIT
"""

import sys
import re
import os
from docx import Document
from docx.shared import Pt, Mm, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml


# ============================================================
# 定数
# ============================================================


# 見出しレベルごとの設定: (title_start, number_hang)
# title_start: タイトル開始位置（文字数） = 本文の左インデントと一致
# number_hang: 見出し番号のぶら下げ幅（文字数）
# 見出し: left=title_start, hanging=number_hang → 番号が左に飛び出す
# 本文:   left=title_start → タイトルと頭が揃う
HEADING_LEVELS = {
    1: (3, 3),   # 第１　→ left=3, hang=3 → 番号は0から開始
    2: (4, 2),   # １　→ left=4, hang=2 → 番号は2から開始
    3: (6, 3),   # (1)　→ left=6, hang=3 → 番号は3から開始
    4: (6, 2),   # ア　→ left=6, hang=2 → 番号は4から開始
    5: (8, 3),   # (ｱ)　→ left=8, hang=3 → 番号は5から開始
    6: (8, 2),   # ａ　→ left=8, hang=2 → 番号は6から開始
    7: (10, 3),  # (a)　→ left=10, hang=3 → 番号は7から開始
}

# 本文インデント: (左インデント文字数, 首行字下げ文字数)
# 首行(1行目) = 左 + 首行 = タイトル開始位置 → 「走」と「被」が揃う
# 2行目以降 = 左 = タイトル開始位置 - 1字
BODY_INDENT = {
    0: (0, 1),   # 見出しなし直後
    1: (2, 1),   # 第１直下 → 1行目: 2+1=3=「走」位置、2行目: 2
    2: (3, 1),   # １直下 → 1行目: 3+1=4=「走」位置、2行目: 3
    3: (5, 1),   # (1)直下 → 1行目: 5+1=6、2行目: 5
    4: (5, 1),   # ア直下 → 1行目: 5+1=6、2行目: 5
    5: (7, 1),   # (ｱ)直下 → 1行目: 7+1=8、2行目: 7
    6: (7, 1),   # ａ直下 → 1行目: 7+1=8、2行目: 7
    7: (9, 1),   # (a)直下 → 1行目: 9+1=10、2行目: 9
}


# ============================================================
# 半角→全角変換
# ============================================================

_HANKAKU_TO_ZENKAKU = str.maketrans(
    '0123456789'
    'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    'abcdefghijklmnopqrstuvwxyz'
    '()[]{}!?.,;:/-+=%&#@*~',
    '０１２３４５６７８９'
    'ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ'
    'ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ'
    '（）［］｛｝！？．，；：／−＋＝％＆＃＠＊〜'
)

_HANKAKU_KANA_MAP = {
    'ｱ': 'ア', 'ｲ': 'イ', 'ｳ': 'ウ', 'ｴ': 'エ', 'ｵ': 'オ',
    'ｶ': 'カ', 'ｷ': 'キ', 'ｸ': 'ク', 'ｹ': 'ケ', 'ｺ': 'コ',
    'ｻ': 'サ', 'ｼ': 'シ', 'ｽ': 'ス', 'ｾ': 'セ', 'ｿ': 'ソ',
    'ﾀ': 'タ', 'ﾁ': 'チ', 'ﾂ': 'ツ', 'ﾃ': 'テ', 'ﾄ': 'ト',
    'ﾅ': 'ナ', 'ﾆ': 'ニ', 'ﾇ': 'ヌ', 'ﾈ': 'ネ', 'ﾉ': 'ノ',
    'ﾊ': 'ハ', 'ﾋ': 'ヒ', 'ﾌ': 'フ', 'ﾍ': 'ヘ', 'ﾎ': 'ホ',
    'ﾏ': 'マ', 'ﾐ': 'ミ', 'ﾑ': 'ム', 'ﾒ': 'メ', 'ﾓ': 'モ',
    'ﾔ': 'ヤ', 'ﾕ': 'ユ', 'ﾖ': 'ヨ',
    'ﾗ': 'ラ', 'ﾘ': 'リ', 'ﾙ': 'ル', 'ﾚ': 'レ', 'ﾛ': 'ロ',
    'ﾜ': 'ワ', 'ﾝ': 'ン',
    'ﾞ': '゛', 'ﾟ': '゜', 'ｰ': 'ー', '｡': '。', '｢': '「', '｣': '」', '､': '、',
}

_DAKUTEN_MAP = {
    'カ゛': 'ガ', 'キ゛': 'ギ', 'ク゛': 'グ', 'ケ゛': 'ゲ', 'コ゛': 'ゴ',
    'サ゛': 'ザ', 'シ゛': 'ジ', 'ス゛': 'ズ', 'セ゛': 'ゼ', 'ソ゛': 'ゾ',
    'タ゛': 'ダ', 'チ゛': 'ヂ', 'ツ゛': 'ヅ', 'テ゛': 'デ', 'ト゛': 'ド',
    'ハ゛': 'バ', 'ヒ゛': 'ビ', 'フ゛': 'ブ', 'ヘ゛': 'ベ', 'ホ゛': 'ボ',
    'ウ゛': 'ヴ',
    'ハ゜': 'パ', 'ヒ゜': 'ピ', 'フ゜': 'プ', 'ヘ゜': 'ペ', 'ホ゜': 'ポ',
}


def to_zenkaku(text):
    """半角文字を全角に変換する。"""
    result = []
    for ch in text:
        result.append(_HANKAKU_KANA_MAP.get(ch, ch))
    text = ''.join(result)

    for src, dst in _DAKUTEN_MAP.items():
        text = text.replace(src, dst)

    text = text.translate(_HANKAKU_TO_ZENKAKU)
    return text


def convert_run_to_zenkaku(run):
    """run内のテキストを全角に変換。"""
    if run.text:
        run.text = to_zenkaku(run.text)


# ============================================================
# 見出し判定
# ============================================================

HEADING_PATTERNS = [
    (1, re.compile(r'^[\s　]*第[１２３４５６７８９０\d]+[\s　]')),
    (3, re.compile(r'^[\s　]*[\(（][１２３４５６７８９０\d]+[\)）][\s　]')),
    (3, re.compile(r'^[\s　]*[⑴⑵⑶⑷⑸⑹⑺⑻⑼⑽⑾⑿⒀⒁⒂⒃⒄⒅⒆⒇][\s　]')),  # 括弧付き数字（(1)の表記揺れ）
    # ①②③は意図的に使うケースがあるため見出しとして検出しない
    (5, re.compile(r'^[\s　]*[\(（][ｱ-ﾝア-ン]+[\)）][\s　]')),
    (7, re.compile(r'^[\s　]*[\(（][a-zａ-ｚ]+[\)）][\s　]')),
    (2, re.compile(r'^[\s　]*[１２３４５６７８９０\d]+[\s　]')),
    (4, re.compile(r'^[\s　]*[ア-ン][　\s]')),
    (6, re.compile(r'^[\s　]*[ａ-ｚ][　\s]')),
]

SKIP_PATTERNS = re.compile(r'^[\s　]*(以上|記|別紙|添付|目録)[\s　]*$')

HEADER_PATTERNS = [
    re.compile(r'(原告|被告|申立人|被申立人|相手方|抗告人|債権者|債務者)'),
    re.compile(r'(準備書面|訴状|答弁書|意見書|報告書|申立書|陳述書|上申書)'),
    re.compile(r'(令和|平成|昭和)[０-９\d]+年'),
    re.compile(r'(弁護士|弁護人|代理人)'),
    re.compile(r'(裁判所|御[　\s]*中|殿)'),
    re.compile(r'(号証|甲|乙|丙)第?[０-９\d]'),
    re.compile(r'^[\s　]*(第[０-９\d]+[　\s])'),
]

# タイトル行の判定（16pt太字にする対象）
TITLE_PATTERN = re.compile(
    r'(準備書面|訴状|答弁書|意見書|報告書|申立書|陳述書|上申書|'
    r'申請書|請求書|通知書|催告書|告訴状|告発状|嘆願書|'
    r'抗告理由書|控訴理由書|上告理由書)'
)


def detect_heading_level(text):
    """段落テキストから見出しレベルを判定。見出しでなければ None を返す。
    番号の後に見出しテキストがない場合（数字だけの段落等）は見出しと見なさない。
    """
    stripped = text.strip().replace('\u3000', '　')

    if not stripped:
        return None

    if SKIP_PATTERNS.match(stripped):
        return None

    for level, pattern in HEADING_PATTERNS:
        if pattern.match(stripped):
            # 番号を剥いだ後にテキストが残るか確認
            body = HEADING_STRIP_RE.sub('', stripped, count=1).strip()
            if body:
                return level
            return None  # 番号だけの段落は見出しではない

    return None


def is_header_section(text):
    """冒頭セクション（事件番号〜弁護士名）かどうか判定。"""
    for pat in HEADER_PATTERNS:
        if pat.search(text):
            return True
    return False


# ============================================================
# 見出し番号の剥ぎ取りと再付番
# ============================================================

# 見出し番号を剥ぎ取る正規表現（全パターン対応）
HEADING_STRIP_RE = re.compile(
    r'^[\s　]*'
    r'(?:'
    r'第[１２３４５６７８９０\d]+'          # 第１、第２
    r'|[\(（][１２３４５６７８９０\d]+[\)）]'  # (1)、（１）
    r'|[⑴⑵⑶⑷⑸⑹⑺⑻⑼⑽⑾⑿⒀⒁⒂⒃⒄⒅⒆⒇]'  # ⑴、⑵（(1)の表記揺れ）
    r'|[\(（][ｱ-ﾝア-ン]+[\)）]'             # (ｱ)、（ア）
    r'|[\(（][a-zａ-ｚ]+[\)）]'             # (a)、（ａ）
    r'|[１２３４５６７８９０\d]+'            # １、２
    r'|[ア-ン]'                             # ア、イ
    r'|[ａ-ｚ]'                             # ａ、ｂ
    r')'
    r'[\s　]*'  # 番号の後のスペース
)

# 全角数字テーブル
_ZEN_DIGITS = '０１２３４５６７８９'
# 全角カタカナ順
_ZEN_KATAKANA = 'アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワ'
# 半角カタカナ順
_HAN_KATAKANA = 'ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜ'
# 全角小文字英字順
_ZEN_ALPHA = 'ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ'
# 半角小文字英字順
_HAN_ALPHA = 'abcdefghijklmnopqrstuvwxyz'


def _to_zenkaku_num(n):
    """整数を全角数字文字列に変換。"""
    return ''.join(_ZEN_DIGITS[int(d)] for d in str(n))


def generate_heading_number(level, counter):
    """レベルとカウンター値から裁判所書式の見出し番号を生成。"""
    if level == 1:
        return f'第{_to_zenkaku_num(counter)}　'
    elif level == 2:
        return f'{_to_zenkaku_num(counter)}　'
    elif level == 3:
        return f'({counter})　'
    elif level == 4:
        idx = counter - 1
        if idx < len(_ZEN_KATAKANA):
            return f'{_ZEN_KATAKANA[idx]}　'
        return f'{_ZEN_KATAKANA[0]}　'
    elif level == 5:
        idx = counter - 1
        if idx < len(_HAN_KATAKANA):
            return f'({_HAN_KATAKANA[idx]})　'
        return f'({_HAN_KATAKANA[0]})　'
    elif level == 6:
        idx = counter - 1
        if idx < len(_ZEN_ALPHA):
            return f'{_ZEN_ALPHA[idx]}　'
        return f'{_ZEN_ALPHA[0]}　'
    elif level == 7:
        idx = counter - 1
        if idx < len(_HAN_ALPHA):
            return f'({_HAN_ALPHA[idx]})　'
        return f'({_HAN_ALPHA[0]})　'
    return ''


def strip_heading_number(text):
    """見出しテキストから既存の番号部分を除去し、本文部分だけ返す。"""
    return HEADING_STRIP_RE.sub('', text, count=1)


class HeadingCounter:
    """階層別の見出しカウンター。上位レベルが進むと下位をリセット。"""

    def __init__(self):
        self._counts = {i: 0 for i in range(1, 8)}

    def increment(self, level):
        """指定レベルのカウンターを進め、下位レベルをリセット。"""
        self._counts[level] += 1
        for lv in range(level + 1, 8):
            self._counts[lv] = 0
        return self._counts[level]


# ============================================================
# フォント設定
# ============================================================

def set_run_font(run, size=12):
    """runのフォントをMS明朝/Times New Roman に設定。"""
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = parse_xml(f'<w:rFonts {nsdecls("w")}/>')
        rpr.insert(0, rfonts)
    rfonts.set(qn('w:eastAsia'), 'ＭＳ 明朝')
    rfonts.set(qn('w:ascii'), 'Times New Roman')
    rfonts.set(qn('w:hAnsi'), 'Times New Roman')


def set_paragraph_font(para, size=12):
    """段落内の全runのフォントを設定。"""
    for run in para.runs:
        set_run_font(run, size)


# ============================================================
# インデント設定
# ============================================================

def _clear_indent(para):
    """段落の既存インデント設定をクリア。"""
    pPr = para._element.get_or_add_pPr()
    existing_ind = pPr.find(qn('w:ind'))
    if existing_ind is not None:
        pPr.remove(existing_ind)


def _set_indent_twips(para, left_twips=0, hanging_twips=0, first_line_twips=0):
    """段落のインデントをtwips絶対値で設定。グリッドに依存しない。"""
    _clear_indent(para)
    if left_twips == 0 and hanging_twips == 0 and first_line_twips == 0:
        return

    ind = parse_xml(f'<w:ind {nsdecls("w")}/>')
    if left_twips > 0:
        ind.set(qn('w:left'), str(left_twips))
    if hanging_twips > 0:
        ind.set(qn('w:hanging'), str(hanging_twips))
    if first_line_twips > 0:
        ind.set(qn('w:firstLine'), str(first_line_twips))
    para._element.get_or_add_pPr().append(ind)


# 12pt MS明朝 + グリッド補正: 全角1文字 ≈ 242twips
CHAR_WIDTH = 242

def set_heading_indent(para, level):
    """見出し段落のインデント設定。
    左インデント=タイトル開始位置、ぶら下げ=番号幅。
    番号が左に飛び出し、タイトルは左インデント位置から始まる。
    """
    title_start, number_hang = HEADING_LEVELS[level]
    _set_indent_twips(para,
                      left_twips=title_start * CHAR_WIDTH,
                      hanging_twips=number_hang * CHAR_WIDTH)


def set_body_indent(para, current_heading_level):
    """本文段落のインデント設定。
    2行目以降 = 見出しタイトル位置に揃う。
    1行目 = さらに1字右（首行字下げ）。
    """
    left, fl = BODY_INDENT.get(current_heading_level, (1, 0))
    _set_indent_twips(para,
                      left_twips=left * CHAR_WIDTH,
                      first_line_twips=fl * CHAR_WIDTH)


# ============================================================
# ページ設定
# ============================================================

def setup_page(doc):
    """ページ設定を裁判所書式に変更。"""
    for section in doc.sections:
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.top_margin = Mm(35)
        section.bottom_margin = Mm(25)
        section.left_margin = Mm(30)
        section.right_margin = Mm(20)
        section.header_distance = Mm(0)
        section.footer_distance = Mm(15)

        sectPr = section._sectPr
        docGrid = sectPr.find(qn('w:docGrid'))
        if docGrid is None:
            docGrid = parse_xml(
                f'<w:docGrid {nsdecls("w")} '
                f'w:type="linesAndChars" w:linePitch="516" w:charSpace="1057"/>'
            )
            sectPr.append(docGrid)
        else:
            docGrid.set(qn('w:type'), 'linesAndChars')
            docGrid.set(qn('w:linePitch'), '516')
            docGrid.set(qn('w:charSpace'), '1057')


# ============================================================
# ページ番号
# ============================================================

def add_page_number(doc):
    """フッター中央にページ番号を挿入。"""
    for section in doc.sections:
        footer = section.footer
        footer.is_linked_to_previous = False
        for p in footer.paragraphs:
            p.clear()

        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER

        run1 = fp.add_run()
        set_run_font(run1, size=10)
        fld_begin = parse_xml(
            f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>'
        )
        run1._element.append(fld_begin)

        run2 = fp.add_run()
        set_run_font(run2, size=10)
        instr = parse_xml(
            f'<w:instrText {nsdecls("w")} xml:space="preserve"> PAGE </w:instrText>'
        )
        run2._element.append(instr)

        run3 = fp.add_run()
        set_run_font(run3, size=10)
        fld_end = parse_xml(
            f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>'
        )
        run3._element.append(fld_end)


# ============================================================
# テーブル処理
# ============================================================

def format_tables(doc):
    """テーブルのフォントを10ptに縮小、セル内インデントをリセット、列幅を自動調整。"""
    for table in doc.tables:
        tbl = table._tbl
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = parse_xml(f'<w:tblPr {nsdecls("w")}/>')
            tbl.insert(0, tblPr)

        # 列幅自動調整
        existing = tblPr.find(qn('w:tblLayout'))
        if existing is not None:
            tblPr.remove(existing)
        layout = parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="autofit"/>')
        tblPr.append(layout)

        for row in table.rows:
            for cell in row.cells:
                # セル内余白を最小化
                tcPr = cell._tc.get_or_add_tcPr()
                existing_mar = tcPr.find(qn('w:tcMar'))
                if existing_mar is not None:
                    tcPr.remove(existing_mar)
                tcMar = parse_xml(
                    f'<w:tcMar {nsdecls("w")}>'
                    f'<w:top w:w="0" w:type="dxa"/>'
                    f'<w:left w:w="28" w:type="dxa"/>'
                    f'<w:bottom w:w="0" w:type="dxa"/>'
                    f'<w:right w:w="28" w:type="dxa"/>'
                    f'</w:tcMar>'
                )
                tcPr.append(tcMar)

                for para in cell.paragraphs:
                    # セル内段落のインデントをリセット
                    _clear_indent(para)
                    # フォント10pt
                    set_paragraph_font(para, size=10)


# ============================================================
# デフォルトスタイル設定
# ============================================================

def setup_default_style(doc):
    """Normalスタイルのフォントを設定。"""
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)

    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = parse_xml(f'<w:rFonts {nsdecls("w")}/>')
        rpr.insert(0, rfonts)
    rfonts.set(qn('w:eastAsia'), 'ＭＳ 明朝')
    rfonts.set(qn('w:ascii'), 'Times New Roman')
    rfonts.set(qn('w:hAnsi'), 'Times New Roman')


# ============================================================
# メイン変換処理
# ============================================================

def _detect_level_offset(doc):
    """文書内の見出しレベルをスキャンし、最上位レベルへのオフセットを算出。

    例: 文書が「１」(L2)始まりなら offset=1 → L2をL1にシフト。
    「第１」(L1)始まりなら offset=0。
    """
    in_header = True
    found_levels = []

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        if in_header:
            level = detect_heading_level(text)
            if level is not None:
                in_header = False
                found_levels.append(level)
            elif is_header_section(text):
                continue
            else:
                continue
        else:
            level = detect_heading_level(text)
            if level is not None:
                found_levels.append(level)

    if not found_levels:
        return 0

    min_level = min(found_levels)
    return min_level - 1  # L1始まりなら0、L2始まりなら1、L3始まりなら2


def _remap_level(raw_level, offset):
    """検出されたレベルをオフセット分シフトして正規化。"""
    adjusted = raw_level - offset
    return max(1, min(adjusted, 7))


def convert(input_path, output_path=None):
    """docxファイルを裁判所書式に変換。"""

    if output_path is None:
        base, ext = os.path.splitext(input_path)
        output_path = f"{base}_裁判所書式{ext}"

    doc = Document(input_path)

    # Pass 1: レベルオフセット算出
    level_offset = _detect_level_offset(doc)

    setup_page(doc)
    setup_default_style(doc)

    # Pass 2: 変換適用（再付番あり）
    current_heading_level = 0
    in_header_section = True
    counter = HeadingCounter()

    for para in doc.paragraphs:
        for run in para.runs:
            convert_run_to_zenkaku(run)

        text = para.text.strip()
        set_paragraph_font(para, size=12)

        if not text:
            continue

        if in_header_section:
            level = detect_heading_level(text)
            if level is not None:
                in_header_section = False
            elif is_header_section(text):
                # タイトル行を検出して16pt太字に
                if TITLE_PATTERN.search(text):
                    set_paragraph_font(para, size=16)
                    for run in para.runs:
                        run.font.bold = True
                continue
            else:
                continue

        if not in_header_section:
            level = detect_heading_level(text)

            if level is not None:
                adjusted = _remap_level(level, level_offset)
                current_heading_level = adjusted

                # 再付番: 元の番号を剥がして裁判所書式の番号に置換
                body_text = strip_heading_number(para.text)
                count = counter.increment(adjusted)
                new_number = generate_heading_number(adjusted, count)
                new_text = new_number + body_text

                # 段落テキストを置換（最初のrunに全テキスト、残りを空に）
                for i, run in enumerate(para.runs):
                    if i == 0:
                        run.text = new_text
                    else:
                        run.text = ''

                set_heading_indent(para, adjusted)
                para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                if SKIP_PATTERNS.match(text):
                    para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    _clear_indent(para)
                else:
                    # 元の手動インデント（先頭の全角スペース）を除去
                    raw = para.text
                    stripped = raw.lstrip('\u3000 \t')
                    # 箇条書き番号の後の全角スペースを除去（「１．　」→「１．」）
                    stripped = re.sub(r'^([０-９\d]+．)[\u3000\s]+', r'\1', stripped)
                    if stripped != raw and para.runs:
                        for i, run in enumerate(para.runs):
                            if i == 0:
                                run.text = stripped
                            else:
                                run.text = ''

                    # 箇条書き（１．, ２．等）はぶら下げインデントで番号後に揃える
                    current_text = para.runs[0].text if para.runs else stripped
                    list_match = re.match(r'^[０-９\d]+．', current_text)
                    if list_match:
                        num_width = len(list_match.group())  # 番号部分の文字数
                        body_left, _ = BODY_INDENT.get(current_heading_level, (0, 0))
                        # 左インデント = 本文位置 + 番号幅、ぶら下げ = 番号幅
                        total_left = (body_left + num_width) * CHAR_WIDTH
                        hang = num_width * CHAR_WIDTH
                        _set_indent_twips(para, left_twips=total_left, hanging_twips=hang)
                    else:
                        set_body_indent(para, current_heading_level)
                    if para.alignment == WD_ALIGN_PARAGRAPH.CENTER:
                        para.alignment = WD_ALIGN_PARAGRAPH.LEFT

    format_tables(doc)
    add_page_number(doc)
    doc.save(output_path)
    return output_path


# ============================================================
# CLI
# ============================================================

if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(
        description='裁判所提出書面の書式整形ツール'
    )
    parser.add_argument('input', help='入力docxファイル')
    parser.add_argument('output', nargs='?', default=None,
                        help='出力docxファイル（省略時は_裁判所書式.docxを付加）')

    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"エラー: ファイルが見つかりません: {args.input}")
        sys.exit(1)

    result = convert(args.input, args.output)
    print(f"変換完了: {result}")
