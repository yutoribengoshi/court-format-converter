/**
 * court-format-converter: Word アドイン版
 * Python版 court_format_converter.py の忠実な移植
 *
 * Office.jsのleftIndentはw:leftCharsを消せない問題があるため、
 * インデント設定はOOXML直接操作で行う。
 */

// ============================================================
// 定数（Python版と同一）
// ============================================================

const PAGE = {
  topMargin: 99.21,    // 35mm in pt
  bottomMargin: 70.87, // 25mm
  leftMargin: 85.04,   // 30mm
  rightMargin: 56.69,  // 20mm
};

// 12pt MS明朝 + グリッド補正
const CHAR_WIDTH = 242; // twips (Python版と同一)

const HEADING_LEVELS = {
  1: [3, 3], 2: [4, 2], 3: [6, 3], 4: [6, 2],
  5: [8, 3], 6: [8, 2], 7: [10, 3],
};

const BODY_INDENT = {
  0: [0, 1], 1: [2, 1], 2: [3, 1], 3: [5, 1],
  4: [5, 1], 5: [7, 1], 6: [7, 1], 7: [9, 1],
};

const TITLE_PATTERN = /(準備書面|訴状|答弁書|意見書|報告書|申立書|陳述書|上申書|申請書|請求書|通知書|催告書|告訴状|告発状|嘆願書|抗告理由書|控訴理由書|上告理由書)/;
const LIST_PATTERN = /^[０-９\d]+．/;

// ============================================================
// 半角→全角変換
// ============================================================

const HANKAKU_KANA_MAP = {'ｱ':'ア','ｲ':'イ','ｳ':'ウ','ｴ':'エ','ｵ':'オ','ｶ':'カ','ｷ':'キ','ｸ':'ク','ｹ':'ケ','ｺ':'コ','ｻ':'サ','ｼ':'シ','ｽ':'ス','ｾ':'セ','ｿ':'ソ','ﾀ':'タ','ﾁ':'チ','ﾂ':'ツ','ﾃ':'テ','ﾄ':'ト','ﾅ':'ナ','ﾆ':'ニ','ﾇ':'ヌ','ﾈ':'ネ','ﾉ':'ノ','ﾊ':'ハ','ﾋ':'ヒ','ﾌ':'フ','ﾍ':'ヘ','ﾎ':'ホ','ﾏ':'マ','ﾐ':'ミ','ﾑ':'ム','ﾒ':'メ','ﾓ':'モ','ﾔ':'ヤ','ﾕ':'ユ','ﾖ':'ヨ','ﾗ':'ラ','ﾘ':'リ','ﾙ':'ル','ﾚ':'レ','ﾛ':'ロ','ﾜ':'ワ','ﾝ':'ン','ﾞ':'゛','ﾟ':'゜','ｰ':'ー','｡':'。','｢':'「','｣':'」','､':'、'};
const DAKUTEN_PAIRS = [['カ゛','ガ'],['キ゛','ギ'],['ク゛','グ'],['ケ゛','ゲ'],['コ゛','ゴ'],['サ゛','ザ'],['シ゛','ジ'],['ス゛','ズ'],['セ゛','ゼ'],['ソ゛','ゾ'],['タ゛','ダ'],['チ゛','ヂ'],['ツ゛','ヅ'],['テ゛','デ'],['ト゛','ド'],['ハ゛','バ'],['ヒ゛','ビ'],['フ゛','ブ'],['ヘ゛','ベ'],['ホ゛','ボ'],['ウ゛','ヴ'],['ハ゜','パ'],['ヒ゜','ピ'],['フ゜','プ'],['ヘ゜','ペ'],['ホ゜','ポ']];

function toZenkaku(text) {
  let r = '';
  for (const ch of text) r += HANKAKU_KANA_MAP[ch] || ch;
  for (const [s, d] of DAKUTEN_PAIRS) r = r.split(s).join(d);
  let o = '';
  for (const ch of r) {
    const c = ch.charCodeAt(0);
    o += (c >= 0x21 && c <= 0x7E) ? String.fromCharCode(c + 0xFEE0) : ch;
  }
  return o;
}

// ============================================================
// 見出し判定（Python版と同一ロジック）
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
const HEADER_PATTERNS = [/(原告|被告|申立人|被申立人|相手方|抗告人|債権者|債務者)/,/(準備書面|訴状|答弁書|意見書|報告書|申立書|陳述書|上申書)/,/(令和|平成|昭和)[０-９\d]+年/,/(弁護士|弁護人|代理人)/,/(裁判所|御[\u3000\s]*中|殿)/,/(号証|甲|乙|丙)第?[０-９\d]/];
const HEADING_STRIP_RE = /^[\s\u3000]*(?:第[１２３４５６７８９０\d]+|[\(（][１２３４５６７８９０\d]+[\)）]|[\(（][ｱ-ﾝア-ン]+[\)）]|[\(（][a-zａ-ｚ]+[\)）]|[１２３４５６７８９０\d]+|[ア-ン]|[ａ-ｚ])[\s\u3000]*/;

function detectHeadingLevel(text) {
  const s = text.trim();
  if (!s || SKIP_PATTERN.test(s)) return null;
  for (const { level, re } of HEADING_PATTERNS) {
    if (re.test(s)) {
      if (s.replace(HEADING_STRIP_RE, '').trim()) return level;
      return null;
    }
  }
  return null;
}
function isHeaderSection(text) {
  return HEADER_PATTERNS.some(p => p.test(text));
}

// ============================================================
// 再付番
// ============================================================

const ZEN_DIGITS = '０１２３４５６７８９';
const ZEN_KATA = 'アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワ';
const HAN_KATA = 'ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜ';
const ZEN_ALPHA = 'ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ';
const HAN_ALPHA = 'abcdefghijklmnopqrstuvwxyz';

function zenNum(n) { return String(n).split('').map(d => ZEN_DIGITS[+d]).join(''); }
function genNum(lv, c) {
  if (lv===1) return `第${zenNum(c)}\u3000`;
  if (lv===2) return `${zenNum(c)}\u3000`;
  if (lv===3) return `(${c})\u3000`;
  if (lv===4) return `${ZEN_KATA[c-1]||ZEN_KATA[0]}\u3000`;
  if (lv===5) return `(${HAN_KATA[c-1]||HAN_KATA[0]})\u3000`;
  if (lv===6) return `${ZEN_ALPHA[c-1]||ZEN_ALPHA[0]}\u3000`;
  if (lv===7) return `(${HAN_ALPHA[c-1]||HAN_ALPHA[0]})\u3000`;
  return '';
}
function stripNum(t) { return t.replace(HEADING_STRIP_RE, ''); }

class Counter {
  constructor() { this.c = {1:0,2:0,3:0,4:0,5:0,6:0,7:0}; }
  inc(lv) { this.c[lv]++; for(let i=lv+1;i<=7;i++) this.c[i]=0; return this.c[lv]; }
}
function remap(raw, off) { return Math.max(1, Math.min(raw - off, 7)); }
function stripSpaces(t) {
  let s = t.replace(/^[\u3000 \t]+/, '');
  return s.replace(/^([０-９\d]+．)[\u3000\s]+/, '$1');
}

// ============================================================
// OOXML直接操作によるインデント設定
// ============================================================

function makeIndentOoxml(text, leftTwips, options) {
  // options: { hanging, firstLine }
  const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  let indAttr = `w:left="${leftTwips}"`;
  if (options.hanging) indAttr += ` w:hanging="${options.hanging}"`;
  if (options.firstLine) indAttr += ` w:firstLine="${options.firstLine}"`;

  // テキストをエスケープ
  const escaped = text.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');

  return `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document ${ns}><w:body><w:p><w:pPr><w:ind ${indAttr}/></w:pPr><w:r><w:t xml:space="preserve">${escaped}</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
}

// ============================================================
// 書式設定（ページ設定のみ）
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
// メイン変換
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

    // 段落読み込み
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load('text, font, alignment');
    await context.sync();

    // レベルオフセット算出
    let levelOffset = 0;
    if (doIndent) {
      let inH = true;
      const lvs = [];
      for (const p of paragraphs.items) {
        const t = p.text.trim();
        if (!t) continue;
        if (inH) {
          const lv = detectHeadingLevel(t);
          if (lv !== null) { inH = false; lvs.push(lv); }
          else if (isHeaderSection(t)) continue;
          else continue;
        } else {
          const lv = detectHeadingLevel(t);
          if (lv !== null) lvs.push(lv);
        }
      }
      if (lvs.length > 0) levelOffset = Math.min(...lvs) - 1;
    }

    // テキスト修正 + フォント + インデント
    let curHL = 0;
    let inHeader = true;
    let firstHdg = false;
    const ctr = new Counter();

    for (const para of paragraphs.items) {
      const text = para.text.trim();

      // フォント
      if (doFont) {
        para.font.name = 'Times New Roman';
        para.font.size = 12;
      }

      // 全角変換
      if (doZenkaku && text) {
        const conv = toZenkaku(para.text);
        if (conv !== para.text) para.insertText(conv, Word.InsertLocation.replace);
      }

      if (!text) continue;

      if (doIndent) {
        // 冒頭セクション
        if (inHeader && !firstHdg) {
          const lv = detectHeadingLevel(text);
          if (lv !== null) {
            inHeader = false;
            firstHdg = true;
            const adj = remap(lv, levelOffset);
            curHL = adj;
            const body = stripNum(para.text);
            const num = genNum(adj, ctr.inc(adj));
            para.insertText(num + body, Word.InsertLocation.replace);
            // インデント: OOXML直接（leftChars残骸を上書き）
            const [ts, nh] = HEADING_LEVELS[adj];
            para.leftIndent = ts * CHAR_WIDTH / 20;
            para.firstLineIndent = -(nh * CHAR_WIDTH / 20);
            para.alignment = Word.Alignment.left;
          } else if (isHeaderSection(text)) {
            if (TITLE_PATTERN.test(text)) {
              para.font.size = 16;
              para.font.bold = true;
            }
          }
          continue;
        }

        // 見出し
        const lv = detectHeadingLevel(text);
        if (lv !== null) {
          const adj = remap(lv, levelOffset);
          curHL = adj;
          const body = stripNum(para.text);
          const num = genNum(adj, ctr.inc(adj));
          para.insertText(num + body, Word.InsertLocation.replace);
          const [ts, nh] = HEADING_LEVELS[adj];
          para.leftIndent = ts * CHAR_WIDTH / 20;
          para.firstLineIndent = -(nh * CHAR_WIDTH / 20);
          para.alignment = Word.Alignment.left;
        } else if (SKIP_PATTERN.test(text)) {
          para.alignment = Word.Alignment.right;
          para.leftIndent = 0;
          para.firstLineIndent = 0;
        } else {
          // 本文: スペース除去
          const raw = para.text;
          const stripped = stripSpaces(raw);
          if (stripped !== raw) para.insertText(stripped, Word.InsertLocation.replace);

          // インデント
          const cur = stripped || raw;
          const lm = cur.match(LIST_PATTERN);
          if (lm) {
            const nw = lm[0].length;
            const [bl] = BODY_INDENT[curHL] || [0, 0];
            para.leftIndent = (bl + nw) * CHAR_WIDTH / 20;
            para.firstLineIndent = -(nw * CHAR_WIDTH / 20);
          } else {
            const [l, f] = BODY_INDENT[curHL] || [0, 1];
            para.leftIndent = l * CHAR_WIDTH / 20;
            para.firstLineIndent = f * CHAR_WIDTH / 20;
          }

          if (para.alignment === Word.Alignment.centered) {
            para.alignment = Word.Alignment.left;
          }
        }
      }
    }

    await context.sync();

    // テーブル: 10ptフォント
    try {
      const tables = context.document.body.tables;
      tables.load('*');
      await context.sync();
      for (const table of tables.items) {
        const range = table.getRange();
        range.font.size = 10;
        range.font.name = 'Times New Roman';
      }
      await context.sync();
    } catch (e) {
      // テーブルなしの場合はスキップ
    }

    // フッターにページ番号
    if (doFooter) {
      const sections = context.document.sections;
      sections.load('body');
      await context.sync();
      for (const section of sections.items) {
        const footer = section.getFooter(Word.HeaderFooterType.primary);
        footer.clear();
        footer.insertOoxml(
          '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
          '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part>' +
          '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
          '<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>' +
          '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>' +
          '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
          '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>' +
          '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>1</w:t></w:r>' +
          '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>' +
          '</w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>',
          Word.InsertLocation.replace
        );
      }
      await context.sync();
    }
  });
}
