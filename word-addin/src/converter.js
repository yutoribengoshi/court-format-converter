/**
 * court-format-converter: Word アドイン版 変換ロジック
 *
 * 文化審議会建議「公用文作成の考え方」（令和4年）および
 * 裁判所実務の書��慣行に���拠した書式整形を行う。
 */

// ============================================================
// 定数
// ============================================================

// ページ設定（ポイント単位: 1mm = 2.8346pt）
const PAGE = {
  width: 595.28,   // 210mm (A4)
  height: 841.89,  // 297mm (A4)
  topMargin: 99.21,    // 35mm
  bottomMargin: 70.87, // 25mm
  leftMargin: 85.04,   // 30mm
  rightMargin: 56.69,  // 20mm
};

// フォント
const FONT = {
  japanese: 'ＭＳ 明朝',
  western: 'Times New Roman',
  size: 12,
  tableSize: 10,
};

// 見出しレベルごとの左インデント（文字数）
// 1文字 = 12pt（フォントサイズ基準）
const HEADING_INDENT = {
  1: 0,  // 第１ → 左端
  2: 2,  // １   → 2字
  3: 3,  // (1)  → 3字
  4: 4,  // ア   → 4字
  5: 5,  // (ｱ)  → 5字
  6: 6,  // ａ   → 6字
  7: 7,  // (a)  → 7字
};

// 本文インデント: [左インデント文字数, 首行字下げ文字数]
// 原則: 左 + 首行 = 見出しの左インデント + 番号幅（タイトル開始位置と揃える）
const BODY_INDENT = {
  0: [0, 1],
  1: [2, 1],   // 第１直下 → 0+3=3 → (2,1)=3
  2: [3, 1],   // １直下 → 2+2=4 → (3,1)=4
  3: [5, 1],   // (1)直下 → 3+3=6 → (5,1)=6
  4: [5, 1],   // ア直下 → 4+2=6 → (5,1)=6
  5: [7, 1],   // (ｱ)直下 → 5+3=8 → (7,1)=8
  6: [7, 1],   // ａ直下 → 6+2=8 → (7,1)=8
  7: [9, 1],   // (a)直下 → 7+3=10 → (9,1)=10
};

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
  ['ハ゜','パ'],['���゜','ピ'],['フ゜','プ'],['ヘ゜','ペ'],['ホ゜','ポ'],
];

function toZenkaku(text) {
  // 1. 半角カタカナ→全角
  let result = '';
  for (const ch of text) {
    result += HANKAKU_KANA_MAP[ch] || ch;
  }
  // 2. 濁点・半濁点合成
  for (const [src, dst] of DAKUTEN_PAIRS) {
    result = result.split(src).join(dst);
  }
  // 3. ASCII半角→全角（数字・英字・記号）
  let out = '';
  for (const ch of result) {
    const code = ch.charCodeAt(0);
    // ! (0x21) ~ ~ (0x7E) → 全角 (0xFF01 ~ 0xFF5E)
    // ただしスペース (0x20) は変換しない
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
  { level: 2, re: /^[\s\u3000]*[１２３４５��７８９０\d]+[\s\u3000]/ },
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

function detectHeadingLevel(text) {
  const stripped = text.trim();
  if (!stripped) return null;
  if (SKIP_PATTERN.test(stripped)) return null;
  for (const { level, re } of HEADING_PATTERNS) {
    if (re.test(stripped)) {
      // 番号を剥いだ後にテキストが残るか確認
      const body = stripped.replace(HEADING_STRIP_RE, '').trim();
      if (body) return level;
      return null; // 番号だけの段落は見出しではない
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
// 見出し番号の剥ぎ取りと再付番
// ============================================================

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

// ============================================================
// メイン変換処理
// ============================================================

function remapLevel(rawLevel, offset) {
  const adjusted = rawLevel - offset;
  return Math.max(1, Math.min(adjusted, 7));
}

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

    // 段落処理
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load('text, font, alignment, leftIndent, firstLineIndent');
    await context.sync();

    // Pass 1: レベルオフセット算出
    let levelOffset = 0;
    if (doIndent) {
      let inHdr = true;
      const foundLevels = [];
      for (const para of paragraphs.items) {
        const t = para.text.trim();
        if (!t) continue;
        if (inHdr) {
          const lv = detectHeadingLevel(t);
          if (lv !== null) {
            inHdr = false;
            foundLevels.push(lv);
          } else if (isHeaderSection(t)) {
            continue;
          } else {
            continue;
          }
        } else {
          const lv = detectHeadingLevel(t);
          if (lv !== null) foundLevels.push(lv);
        }
      }
      if (foundLevels.length > 0) {
        levelOffset = Math.min(...foundLevels) - 1;
      }
    }

    // Pass 2: 変換適用（再付番あり）
    let currentHeadingLevel = 0;
    let inHeaderSection = true;
    let firstHeadingFound = false;
    const counter = new HeadingCounter();

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

      // インデント処理
      if (doIndent) {
        // 冒頭セクション判定
        if (inHeaderSection && !firstHeadingFound) {
          const level = detectHeadingLevel(text);
          if (level !== null) {
            inHeaderSection = false;
            firstHeadingFound = true;
            const adjusted = remapLevel(level, levelOffset);
            currentHeadingLevel = adjusted;

            // 再付番
            const bodyText = stripHeadingNumber(para.text);
            const count = counter.increment(adjusted);
            const newText = generateHeadingNumber(adjusted, count) + bodyText;
            para.insertText(newText, Word.InsertLocation.replace);

            para.leftIndent = HEADING_INDENT[adjusted] * FONT.size;
            para.firstLineIndent = 0;
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

          // 再付番
          const bodyText = stripHeadingNumber(para.text);
          const count = counter.increment(adjusted);
          const newText = generateHeadingNumber(adjusted, count) + bodyText;
          para.insertText(newText, Word.InsertLocation.replace);

          para.leftIndent = HEADING_INDENT[adjusted] * FONT.size;
          para.firstLineIndent = 0;
          para.alignment = Word.Alignment.left;
        } else if (SKIP_PATTERN.test(text)) {
          para.alignment = Word.Alignment.right;
          para.leftIndent = 0;
          para.firstLineIndent = 0;
        } else {
          const [left, fl] = BODY_INDENT[currentHeadingLevel] || [0, 1];
          para.leftIndent = left * FONT.size;
          para.firstLineIndent = fl * FONT.size;
        }
      }
    }

    await context.sync();

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
