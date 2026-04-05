# court-format-converter アドイン版 引き継ぎ

## プロジェクト概要
裁判所提出書面（準備書面・訴状等）のdocxを裁判所書式に整形するツール。
Python CLI版とWord アドイン版（Office.js）の2版がある。

- リポジトリ: https://github.com/yutoribengoshi/court-format-converter
- Python版: `court_format_converter.py` — **全機能正常動作**
- アドイン版: `word-addin/src/converter.js` — **3つの未解決問題あり**

## アーキテクチャ

```
court-format-converter/
├── court_format_converter.py        ← Python CLI版（正常動作・リファレンス実装）
├── word-addin/
│   ├── manifest.xml                 ← Wordアドイン定義（GitHub Pagesホスト）
│   └── src/
│       ├── taskpane.html            ← UI
│       ├── taskpane.js              ← UI制御
│       └── converter.js             ← 変換ロジック（★ここに問題あり）
```

## 変換処理の内容

1. ページ設定（A4、余白35/25/30/20mm、26行x37文字グリッド）
2. フォント統一（MS明朝 12pt / Times New Roman）
3. 見出し自動検出 + 7階層インデント（ぶら下げインデント方式）
4. 見出し番号の自動再付番（「１→(1)→ア」を「第１→１→(1)」に変換）
5. 半角→全角変換
6. タイトル行の16pt太字化
7. 箇条書き（１．等）のぶら下げインデント
8. テーブル処理（10pt、セル余白最小、autofit）
9. フッターページ番号
10. 元文書の手動全角スペースインデント除去

## インデント設計（Python版・正常動作）

```python
CHAR_WIDTH = 242  # twips (12pt MS明朝 + グリッド補正)

# 見出し: [タイトル開始位置(文字数), 番号ぶら下げ幅(文字数)]
HEADING_LEVELS = {
    1: (3, 3),   # 第１　→ left=3*242=726, hanging=3*242=726
    2: (4, 2),   # １
    3: (6, 3),   # (1)
    ...
}

# 本文: [左(文字数), 首行(文字数)]
# 1行目 = left + firstLine = タイトル開始位置（見出しタイトルと揃う）
# 2行目以降 = left = タイトル開始位置 - 1字
BODY_INDENT = {
    1: (2, 1),   # 第１直下 → 1行目: 2+1=3=「走」位置、2行目: 2
    2: (3, 1),   # １直下 → 1行目: 3+1=4、2行目: 3
    ...
}
```

Python版はtwips絶対値で`w:left`と`w:hanging`/`w:firstLine`をOOXML直接設定。
`w:leftChars`は使わない（グリッドとの誤差を回避）。

## ★未解決問題（アドイン版 converter.js）

### 問題1: leftChars残骸でインデントが1字右にズレる

**症状**: 過去にアドインで変換済みの文書を再変換すると、本文が1全角文字分右にズレる。

**原因**: Office.jsの`para.leftIndent = x`は`w:left`を設定するが、既存の`w:leftChars`属性を削除しない。Wordは`w:leftChars`を優先するため、古い値が残って競合する。

**Python版での対処**: `_clear_indent()`で`w:ind`要素を丸ごと削除してから再設定。

```python
def _clear_indent(para):
    pPr = para._element.get_or_add_pPr()
    existing_ind = pPr.find(qn('w:ind'))
    if existing_ind is not None:
        pPr.remove(existing_ind)
```

**JS版で必要な対処**: Office.jsにはOOXML要素を直接削除するAPIがない。`insertOoxml`で段落全体を置換する方法が考えられるが、テキスト・書式の保持が複雑。

### 問題2: body.paragraphsがテーブル内段落を含む

**症状**: 全角変換・フォント12pt設定がテーブル内のテキストにも適用され、テーブルのレイアウトが崩れる。

**原因**: python-docxの`doc.paragraphs`はテーブル内段落を含まないが、Office.jsの`context.document.body.paragraphs`はテーブル内段落も含む。

**試したこと**:
- `paragraph.parentTableOrNullObject`でテーブル内判定 → ロード時にハング
- テーブル処理で後から10ptに戻す → `tables.load('*')`でフリーズ

**JS版で必要な対処**: テーブル内段落をスキップする安定した方法が必要。

### 問題3: テーブル操作でフリーズ

**症状**: `tables.load('*')`や`table.getRange()`を使うとWordが応答しなくなる。

**現状**: テーブル処理を完全に削除済み。テーブルの書式調整はPython CLI版で行う運用。

## 現在のJS版converter.jsの動作状況

| 機能 | 状態 |
|---|---|
| ページ設定 | ✅ 動作 |
| フォント12pt | ✅ 動作（ただしテーブルにも適用される） |
| 見出し検出・インデント | ✅ 新規ファイルでは動作 |
| 見出し再付番 | ✅ 動作 |
| タイトル16pt太字 | ✅ 動作 |
| 全角変換 | ⚠️ デフォルトOFF（テーブルにも適用されるため） |
| テーブル処理 | ❌ 削除済み（フリーズするため） |
| 再変換時のインデント | ❌ leftChars残骸で1字ズレる |

## テスト用ファイル

- `~/Downloads/疎1_GPS解析報告書.docx` — 原本（テーブルなし）
- `~/Downloads/疎2_ナビアプリ走行履歴.docx` — 原本（テーブルあり）

## 期待するゴール

1. テーブル内段落を処理対象から除外し、全角変換・フォント変更がテーブルに影響しない
2. 既に変換済みの文書を再変換してもインデントがズレない（leftChars残骸の除去）
3. テーブルのフォントを10ptに設定できる（フリーズしない方法で）
