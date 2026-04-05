/**
 * court-format-converter: Word アドイン版 変換ロジック
 *
 * 文化審議会建議「公用文作成の考え方」（令和4年）および
 * 裁判所実務の書式慣行に準拠した書式整形を行う。
 */

// ============================================================
// 定数
// ============================================================

const PAGE = {
  topMargin: 99.21,    // 35mm
  bottomMargin: 70.87, // 25mm
  leftMargin: 85.04,   // 30mm
  rightMargin: 56.69,  // 20mm
};

const FONT = {
  western: 'Times New Roman',
  size: 12,
};

// 12pt MS明朝 + グリッド補正: 全角1文字 ≈ 12.1pt
const CHAR_PT = 12.1;

// 見出しレベル: [タイトル開始位置(文字数), 番号ぶら下げ幅(文字数)]
const HEADING_LEVELS = {
  1: [3, 3],   // 第１
  2: [4, 2],   // １
  3: [6, 3],   // (1)
  4: [6, 2],   // ア
  5: [8, 3],   // (ｱ)
  6: [8, 2],   // ａ
  7: [10, 3],  // (a)
};

// 本文インデント: [左(文字数), 首行(文字数)]
const BODY_INDENT = {
  0: [0, 1],
  1: [2, 1],   // 1行目: 2+1=3=タイトル位置、2行目: 2
  2: [3, 1],
  3: [5, 1],
  4: [5, 1],
  5: [7, 1],
  6: [7, 1],
  7: [9, 1],
};

// タイトル行判定
const TITLE_PATTERN = new RegExp(
  '(準備書面|訴状|答弁書|意見書|報告書|申立書|陳述書|上申書|' +
  '申請書|請求書|通知書|催告書|告訴状|告発状|嘆願書|' +
  '抗告理由書|控訴理由書|上告理由書)'
);

// 箇条書き番号パターン
const LIST_PATTERN = /^[０-９\d]+．/;

// ============================================================
// 半角→全角変換
// ============================================================

const HANKAKU_KANA_MAP = {
  'ｱ':'ア','ｲ':'イ','ｳ':'ウ','ｴ':'エ','ｵ':'オ',
  'ｶ':'カ','ｷ':'キ','ｸ':'ク','ｹ':'ケ','ｺ':'コ',
  'ｻ':'サ','ｼ':'シ','ｽ':'ス','ｾ':'セ','ｿ':'ソ',
  'ﾀ':'タ','ﾁ':'チ','ﾂ':'ツ','ﾃ':'テ','ﾄ':'ト',
  'ﾅ':'ナ','ﾆ':'ニ','ﾇ':'ヌ','ﾈ':'ネ','ﾉ':'ノ',
  'ﾊ':'ハ','ﾋ':'ヒ','ﾌ':'フ','ﾍ':'ヘ','ﾎ':'ホ',
  'ﾏ':'マ','ﾐ':'ミ','ﾑ':'ム','ﾒ':'メ','ﾓ':'モ',
  'ﾔ':'ヤ','ﾕ':'ユ','ﾖ':'ヨ',
  'ﾗ':'ラ','ﾘ':'リ','ﾙ':'ル','ﾚ':'レ','ﾛ':'ロ',
  'ﾜ':'ワ','ﾝ':'ン',
  'ﾞ':'゛','ﾟ':'゜','ｰ':'ー','｡':'。','｢':'「','｣':'」','､':'、',
};

const DAKUTEN_PAIRS = [
  ['カ゛','ガ'],['キ゛','ギ'],['ク゛','グ'],['ケ゛','ゲ'],['コ゛','ゴ'],
  ['サ゛','ザ'],['シ゛','ジ'],['ス゛','ズ'],['セ゛','ゼ'],['ソ゛','ゾ'],
  ['タ゛','ダ'],['チ゛','ヂ'],['ツ゛','ヅ'],['テ゛','デ'],['ト゛','ド'],
  ['ハ゛','バ'],['ヒ゛','ビ'],['フ゛','ブ'],['ヘ゛','ベ'],['ホ゛','ボ'],
  ['ウ゛','ヴ'],
  ['ハ゜','パ'],['ヒ゜','ピ'],['フ゜','プ'],['ヘ゜','ペ'],['ホ゜','ポ'],
];

function toZenkaku(text) {
  let result = '';
  for (const ch of text) {
    result += HANKAKU_KANA_MAP[ch] || ch;
  }
  for (const [src, dst] of DAKUTEN_PAIRS) {
    result = result.split(src).join(dst);
  }
  let out = '';
  for (const ch of result) {
    const code = ch.charCodeAt(0);
    if (code >= 0x21 && code <= 0x7E) {
      out += String.fromCharCode(code + 0xFEE0);
    } else {
      out += ch;
    }
  }
  return out;
}

// ============================================================
// 見出し判定
// ============================================================

const HEADING_PATTERNS = [
  { level: 1, re: /^[\s\u3000]*第[１２３４５６７８９０\d]+[\s\u3000]/ },
  { level: 3, re: /^[\s\u3000]*[(\uff08][１２３４５６７８９０\d]+[)\uff09][\s\u3000]/ },
  { level: 5, re: /^[\s\u3000]*[(\uff08][ｱ-ﾝア-ン]+[)\uff09][\s\u3000]/ },
  { level: 7, re: /^[\s\u3000]*[(\uff08][a-zａ-ｚ]+[)\uff09][\s\u3000]/ },
  { level: 2, re: /^[\s\u3000]*[１２３４５６７８９０\d]+[\s\u3000]/ },
  { level: 4, re: /^[\s\u3000]*[ア-ン][\u3000\s]/ },
  { level: 6, re: /^[\s\u3000]*[ａ-ｚ][\u3000\s]/ },
];

const SKIP_PATTERN = /^[\s\u3000]*(以上|記|別紙|添付|目録)[\s\u3000]*$/;

const HEADER_PATTERNS = [
  /(原告|被告|申立人|被申立人|相手方|抗告人|債権者|債務者)/,
  /(準備書面|訴状|答弁書|意見書|報告書|申立書|陳述書|上申書)/,
  /(令和|平成|昭和)[０-９\d]+年/,
  /(弁護士|弁護人|代理人)/,
  /(裁判所|御[\u3000\s]*中|殿)/,
  /(号証|甲|乙|丙)第?[０-９\d]/,
];

const HEADING_STRIP_RE = new RegExp(
  '^[\\s\u3000]*(?:' +
  '第[１２３４５６７８９０\\d]+' +
  '|[\\(（][１２３４５６７８９０\\d]+[\\)）]' +
  '|[\\(（][ｱ-ﾝア-ン]+[\\)）]' +
  '|[\\(（][a-zａ-ｚ]+[\\)）]' +
  '|[１２３４５６７８９０\\d]+' +
  '|[ア-ン]' +
  '|[ａ-ｚ]' +
  ')[\\s\u3000]*'
);

function detectHeadingLevel(text) {
  const stripped = text.trim();
  if (!stripped) return null;
  if (SKIP_PATTERN.test(stripped)) return null;
  for (const { level, re } of HEADING_PATTERNS) {
    if (re.test(stripped)) {
      const body = stripped.replace(HEADING_STRIP_RE, '').trim();
      if (body) return level;
      return null;
    }
  }
  return null;
}

function isHeaderSection(text) {
  for (const pat of HEADER_PATTERNS) {
    if (pat.test(text)) return true;
  }
  return false;
}

// ============================================================
// 見出し番号の再付番
// ============================================================

const ZEN_DIGITS = '０１２３４５６７８９';
const ZEN_KATAKANA = 'アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワ';
const HAN_KATAKANA = 'ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜ';
const ZEN_ALPHA = 'ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ';
const HAN_ALPHA = 'abcdefghijklmnopqrstuvwxyz';

function toZenkakuNum(n) {
  return String(n).split('').map(d => ZEN_DIGITS[parseInt(d)]).join('');
}

function generateHeadingNumber(level, counter) {
  if (level === 1) return `第${toZenkakuNum(counter)}\u3000`;
  if (level === 2) return `${toZenkakuNum(counter)}\u3000`;
  if (level === 3) return `(${counter})\u3000`;
  if (level === 4) return `${ZEN_KATAKANA[counter - 1] || ZEN_KATAKANA[0]}\u3000`;
  if (level === 5) return `(${HAN_KATAKANA[counter - 1] || HAN_KATAKANA[0]})\u3000`;
  if (level === 6) return `${ZEN_ALPHA[counter - 1] || ZEN_ALPHA[0]}\u3000`;
  if (level === 7) return `(${HAN_ALPHA[counter - 1] || HAN_ALPHA[0]})\u3000`;
  return '';
}

function stripHeadingNumber(text) {
  return text.replace(HEADING_STRIP_RE, '');
}

class HeadingCounter {
  constructor() {
    this._counts = {1:0, 2:0, 3:0, 4:0, 5:0, 6:0, 7:0};
  }
  increment(level) {
    this._counts[level]++;
    for (let lv = level + 1; lv <= 7; lv++) {
      this._counts[lv] = 0;
    }
    return this._counts[level];
  }
}

function remapLevel(rawLevel, offset) {
  return Math.max(1, Math.min(rawLevel - offset, 7));
}

// ============================================================
// テキスト前処理
// ============================================================

function stripLeadingSpaces(text) {
  // 先頭の全角スペース・半角スペース・タブを除去
  let stripped = text.replace(/^[\u3000 \t]+/, '');
  // 箇条書き番号の後の全角スペースを除去（「１．　」→「１．」）
  stripped = stripped.replace(/^([０-９\d]+．)[\u3000\s]+/, '$1');
  return stripped;
}

// ============================================================
// 書式設定（新規文書セットアップ）
// ============================================================

async function setupFormat() {
  await Word.run(async (context) => {
    const sections = context.document.sections;
    sections.load('body');
    await context.sync();

    for (const section of sections.items) {
      const ps = section.pageSetup;
      ps.paperSize = Word.PaperSize.a4;
      ps.topMargin = PAGE.topMargin;
      ps.bottomMargin = PAGE.bottomMargin;
      ps.leftMargin = PAGE.leftMargin;
      ps.rightMargin = PAGE.rightMargin;
    }

    await context.sync();
  });
}

// ============================================================
// メイン変換処理
// ============================================================

async function convertDocument(options) {
  const {
    page: doPage = true,
    font: doFont = true,
    indent: doIndent = true,
    zenkaku: doZenkaku = true,
    footer: doFooter = true,
  } = options;

  await Word.run(async (context) => {
    // ページ設定
    if (doPage) {
      const sections = context.document.sections;
      sections.load('body');
      await context.sync();

      for (const section of sections.items) {
        const ps = section.pageSetup;
        ps.paperSize = Word.PaperSize.a4;
        ps.topMargin = PAGE.topMargin;
        ps.bottomMargin = PAGE.bottomMargin;
        ps.leftMargin = PAGE.leftMargin;
        ps.rightMargin = PAGE.rightMargin;
      }
      await context.sync();
    }

    // ========== Phase 1: テキスト修正（全角変換・スペース除去・再付番）==========
    let paragraphs = context.document.body.paragraphs;
    paragraphs.load('text, font, alignment');
    await context.sync();

    // レベルオフセット算出
    let levelOffset = 0;
    if (doIndent) {
      let inHdr = true;
      const foundLevels = [];
      for (const para of paragraphs.items) {
        const t = para.text.trim();
        if (!t) continue;
        if (inHdr) {
          const lv = detectHeadingLevel(t);
          if (lv !== null) { inHdr = false; foundLevels.push(lv); }
          else if (isHeaderSection(t)) continue;
          else continue;
        } else {
          const lv = detectHeadingLevel(t);
          if (lv !== null) foundLevels.push(lv);
        }
      }
      if (foundLevels.length > 0) {
        levelOffset = Math.min(...foundLevels) - 1;
      }
    }

    // テキスト修正パス
    {
      let inHdr = true;
      let firstHdg = false;
      const ctr = new HeadingCounter();

      for (const para of paragraphs.items) {
        const text = para.text.trim();

        // フォント統一
        if (doFont) {
          para.font.name = FONT.western;
          para.font.size = FONT.size;
        }

        // 半角→全角変換
        if (doZenkaku && text) {
          const converted = toZenkaku(para.text);
          if (converted !== para.text) {
            para.insertText(converted, Word.InsertLocation.replace);
          }
        }

        if (!text) continue;

        if (doIndent) {
          if (inHdr && !firstHdg) {
            const level = detectHeadingLevel(text);
            if (level !== null) {
              inHdr = false;
              firstHdg = true;
              const adj = remapLevel(level, levelOffset);
              const body = stripHeadingNumber(para.text);
              const num = generateHeadingNumber(adj, ctr.increment(adj));
              para.insertText(num + body, Word.InsertLocation.replace);
            } else if (isHeaderSection(text)) {
              if (TITLE_PATTERN.test(text)) {
                para.font.size = 16;
                para.font.bold = true;
              }
            }
            continue;
          }

          const level = detectHeadingLevel(text);
          if (level !== null) {
            const adj = remapLevel(level, levelOffset);
            const body = stripHeadingNumber(para.text);
            const num = generateHeadingNumber(adj, ctr.increment(adj));
            para.insertText(num + body, Word.InsertLocation.replace);
          } else if (!SKIP_PATTERN.test(text)) {
            // 本文: 先頭全角スペース除去
            const raw = para.text;
            const stripped = stripLeadingSpaces(raw);
            if (stripped !== raw) {
              para.insertText(stripped, Word.InsertLocation.replace);
            }
          }
        }
      }
    }

    // テキスト修正をコミット
    await context.sync();

    // ========== Phase 2: インデント設定（修正後のテキストを再読み込み）==========
    if (doIndent) {
      paragraphs = context.document.body.paragraphs;
      paragraphs.load('text, alignment, leftIndent, firstLineIndent');
      await context.sync();

      // 全段落の既存インデントをリセット（leftChars等の残骸を消す）
      for (const para of paragraphs.items) {
        para.leftIndent = 0;
        para.firstLineIndent = 0;
      }
      await context.sync();

      // 再読み込み
      paragraphs = context.document.body.paragraphs;
      paragraphs.load('text, alignment, leftIndent, firstLineIndent');
      await context.sync();

      let currentHeadingLevel = 0;
      let inHeaderSection = true;
      let firstHeadingFound = false;

      for (const para of paragraphs.items) {
        const text = para.text.trim();
        if (!text) continue;

        // 冒頭セクション判定
        if (inHeaderSection && !firstHeadingFound) {
          const level = detectHeadingLevel(text);
          if (level !== null) {
            inHeaderSection = false;
            firstHeadingFound = true;
            const adjusted = remapLevel(level, levelOffset);
            currentHeadingLevel = adjusted;

            const [titleStart, numHang] = HEADING_LEVELS[adjusted];
            para.leftIndent = titleStart * CHAR_PT;
            para.firstLineIndent = -(numHang * CHAR_PT);
            para.alignment = Word.Alignment.left;
            continue;
          } else if (isHeaderSection(text)) {
            continue;
          } else {
            continue;
          }
        }

        // 見出し or 本文
        const level = detectHeadingLevel(text);
        if (level !== null) {
          const adjusted = remapLevel(level, levelOffset);
          currentHeadingLevel = adjusted;

          const [titleStart, numHang] = HEADING_LEVELS[adjusted];
          para.leftIndent = titleStart * CHAR_PT;
          para.firstLineIndent = -(numHang * CHAR_PT);
          para.alignment = Word.Alignment.left;
        } else if (SKIP_PATTERN.test(text)) {
          para.alignment = Word.Alignment.right;
          para.leftIndent = 0;
          para.firstLineIndent = 0;
        } else {
          // 箇条書き
          const listMatch = text.match(LIST_PATTERN);
          if (listMatch) {
            const numWidth = listMatch[0].length;
            const [bodyLeft] = BODY_INDENT[currentHeadingLevel] || [0, 0];
            para.leftIndent = (bodyLeft + numWidth) * CHAR_PT;
            para.firstLineIndent = -(numWidth * CHAR_PT);
          } else {
            const [left, fl] = BODY_INDENT[currentHeadingLevel] || [0, 1];
            para.leftIndent = left * CHAR_PT;
            para.firstLineIndent = fl * CHAR_PT;
          }

          if (para.alignment === Word.Alignment.centered) {
            para.alignment = Word.Alignment.left;
          }
        }
      }

      await context.sync();
    }

    // ========== Phase 3: テーブル処理（10pt）==========
    {
      const tables = context.document.body.tables;
      tables.load('*');
      await context.sync();

      for (const table of tables.items) {
        const range = table.getRange();
        range.font.size = 10;
        range.font.name = FONT.western;
      }
      await context.sync();
    }

    // フッターにページ番号
    if (doFooter) {
      const sections = context.document.sections;
      sections.load('body');
      await context.sync();

      for (const section of sections.items) {
        const footer = section.getFooter(Word.HeaderFooterType.primary);
        footer.clear();

        const pageFieldOoxml =
          '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
          '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">' +
          '<pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
          '</Relationships></pkg:xmlData></pkg:part>' +
          '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">' +
          '<pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
          '<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>' +
          '<w:r><w:rPr><w:sz w:val="20"/></w:rPr>' +
          '<w:fldChar w:fldCharType="begin"/></w:r>' +
          '<w:r><w:rPr><w:sz w:val="20"/></w:rPr>' +
          '<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
          '<w:r><w:rPr><w:sz w:val="20"/></w:rPr>' +
          '<w:fldChar w:fldCharType="separate"/></w:r>' +
          '<w:r><w:rPr><w:sz w:val="20"/></w:rPr>' +
          '<w:t>1</w:t></w:r>' +
          '<w:r><w:rPr><w:sz w:val="20"/></w:rPr>' +
          '<w:fldChar w:fldCharType="end"/></w:r>' +
          '</w:p></w:body></w:document></pkg:xmlData></pkg:part>' +
          '</pkg:package>';

        footer.insertOoxml(pageFieldOoxml, Word.InsertLocation.replace);
      }

      await context.sync();
    }
  });
}
