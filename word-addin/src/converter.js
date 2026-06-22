/**
 * court-format-converter: Word アドイン版
 * Python版 court_format_converter.py の参照実装に寄せつつ、
 * 本文段落とテーブルを OOXML 上で分離して処理する。
 */

// ============================================================
// 定数
// ============================================================

const PAGE = {
  topMargin: 99.21,    // 35mm in pt
  bottomMargin: 70.87, // 25mm
  leftMargin: 85.04,   // 30mm
  rightMargin: 56.69,  // 20mm
};

// 1全角文字 = 245 twips (12pt MS明朝基準)
const CHAR_TWIPS = 245;
const BODY_FONT_SIZE = 12;
const TABLE_FONT_SIZE = 10;
const TITLE_FONT_SIZE = 16;

// ------------------------------------------------------------
// インデント方式の切替フラグ（CLI版 court_format_converter.py と統一）
//   true  : 段落スタイル方式（裁判L1〜L5 / 裁判本文L1〜L5 を割当）。デフォルト。
//           インデントを styles.xml 側に持たせ、Word でスタイル定義を1箇所
//           変えれば階層全体のインデントを一括調整できる。
//   false : 従来の直接インデント方式（各段落に w:ind を直接設定）。
// ------------------------------------------------------------
const USE_STYLE_MODE = true;

// 平野晋（筑波大法科大学院）テンプレートの標準インデント値（単位: cm）。
// CLI版 _HIRANO_STYLE_CM と同一。1cm = 567 twips。
// 見出しは hanging（ぶら下げ）、本文は firstLine（字下げ）。
//   level -> { headingLeft, headingHanging, bodyLeft, bodyFirstLine }
const HIRANO_STYLE_CM = {
  1: { headingLeft: 0.801, headingHanging: 0.801, bodyLeft: 0.499, bodyFirstLine: 0.499 },
  2: { headingLeft: 0.741, headingHanging: 0.37, bodyLeft: 0.741, bodyFirstLine: 0.37 },
  3: { headingLeft: 1.109, headingHanging: 0.741, bodyLeft: 1.109, bodyFirstLine: 0.37 },
  4: { headingLeft: 1.482, headingHanging: 0.37, bodyLeft: 1.482, bodyFirstLine: 0.37 },
  5: { headingLeft: 2.223, headingHanging: 1.111, bodyLeft: 2.223, bodyFirstLine: 0.37 },
};

// cm -> twips 変換係数（CHAR_TWIPS=245 とは別系統。スタイルの w:ind は twips 直値で書く）
const CM_TO_TWIPS = 567;

// スタイル方式が扱う見出しレベルの上限（L1〜L5）。
// L6・L7 は最深の L5 スタイルにフォールバックする。
const STYLE_MAX_LEVEL = 5;

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PKG_NS = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const XML_NS = 'http://www.w3.org/XML/1998/namespace';
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

// 見出しレベルごとの設定: [left_chars, 番号説明]
// 全角文字単位の整数。半角は使わない。
const HEADING_LEVELS = {
  1: [0, '第１'],
  2: [2, '１'],
  3: [3, '(1)'],
  4: [4, 'ア'],
  5: [5, '(ｱ)'],
  6: [6, 'ａ'],
  7: [7, '(a)'],
};

// 本文インデント: [left_chars, first_line_chars]
const BODY_INDENT = {
  0: [0, 1],
  1: [2, 1],
  2: [2, 1],
  3: [3, 1],
  4: [4, 1],
  5: [5, 1],
  6: [6, 1],
  7: [7, 1],
};

// 見出しレベルごとのぶら下げ幅（全角文字単位）
const _HEADING_HANGING = {
  1: 0,
  2: 1,
  3: 3,
  4: 1,
  5: 3,
  6: 1,
  7: 3,
};

const TITLE_PATTERN = /(準備書面|訴状|答弁書|意見書|報告書|申立書|陳述書|上申書|申請書|請求書|通知書|催告書|告訴状|告発状|嘆願書|抗告理由書|控訴理由書|上告理由書)/;
const LIST_PATTERN = /^[０-９\d]+．/;
const BULLET_PATTERN = /^[・－\-※①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]/;

const FOOTER_PAGE_NUMBER_OOXML =
  '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
  '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">' +
  '<pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
  '</Relationships></pkg:xmlData></pkg:part>' +
  '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">' +
  '<pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
  '<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>' +
  '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
  '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>1</w:t></w:r>' +
  '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>' +
  '</w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>';

// ============================================================
// 半角→全角変換
// ============================================================

const HANKAKU_KANA_MAP = {'ｱ':'ア','ｲ':'イ','ｳ':'ウ','ｴ':'エ','ｵ':'オ','ｶ':'カ','ｷ':'キ','ｸ':'ク','ｹ':'ケ','ｺ':'コ','ｻ':'サ','ｼ':'シ','ｽ':'ス','ｾ':'セ','ｿ':'ソ','ﾀ':'タ','ﾁ':'チ','ﾂ':'ツ','ﾃ':'テ','ﾄ':'ト','ﾅ':'ナ','ﾆ':'ニ','ﾇ':'ヌ','ﾈ':'ネ','ﾉ':'ノ','ﾊ':'ハ','ﾋ':'ヒ','ﾌ':'フ','ﾍ':'ヘ','ﾎ':'ホ','ﾏ':'マ','ﾐ':'ミ','ﾑ':'ム','ﾒ':'メ','ﾓ':'モ','ﾔ':'ヤ','ﾕ':'ユ','ﾖ':'ヨ','ﾗ':'ラ','ﾘ':'リ','ﾙ':'ル','ﾚ':'レ','ﾛ':'ロ','ﾜ':'ワ','ﾝ':'ン','ﾞ':'゛','ﾟ':'゜','ｰ':'ー','｡':'。','｢':'「','｣':'」','､':'、'};
const DAKUTEN_PAIRS = [['カ゛','ガ'],['キ゛','ギ'],['ク゛','グ'],['ケ゛','ゲ'],['コ゛','ゴ'],['サ゛','ザ'],['シ゛','ジ'],['ス゛','ズ'],['セ゛','ゼ'],['ソ゛','ゾ'],['タ゛','ダ'],['チ゛','ヂ'],['ツ゛','ヅ'],['テ゛','デ'],['ト゛','ド'],['ハ゛','バ'],['ヒ゛','ビ'],['フ゛','ブ'],['ヘ゛','ベ'],['ホ゛','ボ'],['ウ゛','ヴ'],['ハ゜','パ'],['ヒ゜','ピ'],['フ゜','プ'],['ヘ゜','ペ'],['ホ゜','ポ']];

function toZenkaku(text) {
  let kanaConverted = '';
  for (const ch of text) {
    kanaConverted += HANKAKU_KANA_MAP[ch] || ch;
  }

  let composed = kanaConverted;
  for (const [src, dst] of DAKUTEN_PAIRS) {
    composed = composed.split(src).join(dst);
  }

  let output = '';
  for (const ch of composed) {
    const code = ch.charCodeAt(0);
    output += (code >= 0x21 && code <= 0x7E)
      ? String.fromCharCode(code + 0xFEE0)
      : ch;
  }
  return output;
}

// ============================================================
// 見出し判定
// ============================================================

const HEADING_PATTERNS = [
  { level: 1, re: /^[\s\u3000]*第[１２３４５６７８９０\d]+[\s\u3000]/ },
  { level: 3, re: /^[\s\u3000]*[(\uff08][１２３４５６７８９０\d]+[)\uff09][\s\u3000]?/ },
  { level: 5, re: /^[\s\u3000]*[(\uff08][ｱ-ﾝア-ン]+[)\uff09][\s\u3000]?/ },
  { level: 7, re: /^[\s\u3000]*[(\uff08][a-zａ-ｚ]+[)\uff09][\s\u3000]?/ },
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
  /^[\s\u3000]*(第[０-９\d]+[\u3000\s])/,
];

const HEADING_STRIP_RE = /^[\s\u3000]*(?:第[１２３４５６７８９０\d]+|[\(（][１２３４５６７８９０\d]+[\)）]|[\(（][ｱ-ﾝア-ン]+[\)）]|[\(（][a-zａ-ｚ]+[\)）]|[１２３４５６７８９０\d]+|[ア-ン]|[ａ-ｚ])[\s\u3000]*/;

function detectHeadingLevel(text) {
  const stripped = text.trim();
  if (!stripped || SKIP_PATTERN.test(stripped)) {
    return null;
  }

  for (const { level, re } of HEADING_PATTERNS) {
    if (!re.test(stripped)) {
      continue;
    }

    if (stripped.replace(HEADING_STRIP_RE, '').trim()) {
      return level;
    }
    return null;
  }

  return null;
}

function isHeaderSection(text) {
  return HEADER_PATTERNS.some((pattern) => pattern.test(text));
}

// ============================================================
// 再付番
// ============================================================

const ZEN_DIGITS = '０１２３４５６７８９';
const ZEN_KATA = 'アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワ';
const HAN_KATA = 'ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜ';
const ZEN_ALPHA = 'ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ';
const HAN_ALPHA = 'abcdefghijklmnopqrstuvwxyz';

function zenNum(n) {
  return String(n).split('').map((digit) => ZEN_DIGITS[Number(digit)]).join('');
}

function genNum(level, count) {
  if (level === 1) return `第${zenNum(count)}\u3000`;
  if (level === 2) return `${zenNum(count)}\u3000`;
  if (level === 3) return `(${count})\u3000`;
  if (level === 4) return `${ZEN_KATA[count - 1] || ZEN_KATA[0]}\u3000`;
  if (level === 5) return `(${HAN_KATA[count - 1] || HAN_KATA[0]})\u3000`;
  if (level === 6) return `${ZEN_ALPHA[count - 1] || ZEN_ALPHA[0]}\u3000`;
  if (level === 7) return `(${HAN_ALPHA[count - 1] || HAN_ALPHA[0]})\u3000`;
  return '';
}

function stripNum(text) {
  return text.replace(HEADING_STRIP_RE, '');
}

class Counter {
  constructor() {
    this.counts = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 7: 0 };
  }

  inc(level) {
    this.counts[level] += 1;
    for (let nextLevel = level + 1; nextLevel <= 7; nextLevel += 1) {
      this.counts[nextLevel] = 0;
    }
    return this.counts[level];
  }
}

function remap(rawLevel, offset) {
  return Math.max(1, Math.min(rawLevel - offset, 7));
}

function stripSpaces(text) {
  const trimmedLeft = text.replace(/^[\u3000 \t]+/, '');
  return trimmedLeft.replace(/^([０-９\d]+．)[\u3000\s]+/, '$1');
}

// ============================================================
// OOXML ヘルパー
// ============================================================

function isWordElement(node, localName) {
  return Boolean(node)
    && node.nodeType === 1
    && node.namespaceURI === W_NS
    && node.localName === localName;
}

function getWordAttr(element, localName) {
  return element.getAttributeNS(W_NS, localName)
    || element.getAttribute(`w:${localName}`)
    || element.getAttribute(localName)
    || '';
}

function setWordAttr(element, localName, value) {
  element.setAttributeNS(W_NS, `w:${localName}`, String(value));
}

function createWordElement(doc, localName) {
  return doc.createElementNS(W_NS, `w:${localName}`);
}

function getDirectWordChild(parent, localName) {
  return Array.from(parent.childNodes).find((node) => isWordElement(node, localName)) || null;
}

function getDirectWordChildren(parent, localName) {
  return Array.from(parent.childNodes).filter((node) => isWordElement(node, localName));
}

function removeDirectWordChildren(parent, localName) {
  for (const child of getDirectWordChildren(parent, localName)) {
    parent.removeChild(child);
  }
}

function ensureWordChild(parent, localName, { prepend = false } = {}) {
  const existing = getDirectWordChild(parent, localName);
  if (existing) {
    return existing;
  }

  const created = createWordElement(parent.ownerDocument, localName);
  if (prepend) {
    parent.insertBefore(created, parent.firstChild);
  } else {
    parent.appendChild(created);
  }
  return created;
}

function hasAncestor(node, localName) {
  let current = node.parentNode;
  while (current) {
    if (isWordElement(current, localName)) {
      return true;
    }
    current = current.parentNode;
  }
  return false;
}

function parseOoxmlPackage(ooxml) {
  if (typeof DOMParser === 'undefined' || typeof XMLSerializer === 'undefined') {
    throw new Error('この環境では OOXML 変換 API を利用できません。');
  }

  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(ooxml, 'application/xml');
  const parseError = xmlDoc.getElementsByTagName('parsererror')[0];
  if (parseError) {
    throw new Error(`OOXML の解析に失敗しました: ${parseError.textContent}`);
  }
  return xmlDoc;
}

function getPackageAttr(element, localName) {
  return element.getAttributeNS(PKG_NS, localName)
    || element.getAttribute(`pkg:${localName}`)
    || element.getAttribute(localName)
    || '';
}

function getBodyNode(packageDoc) {
  const part = Array.from(packageDoc.getElementsByTagNameNS(PKG_NS, 'part'))
    .find((candidate) => getPackageAttr(candidate, 'name') === '/word/document.xml');
  if (!part) {
    throw new Error('document.xml パートが見つかりません。');
  }

  const xmlData = Array.from(part.childNodes)
    .find((node) => node.nodeType === 1 && node.namespaceURI === PKG_NS && node.localName === 'xmlData');
  if (!xmlData) {
    throw new Error('document.xml の XML データが見つかりません。');
  }

  const documentNode = Array.from(xmlData.childNodes)
    .find((node) => isWordElement(node, 'document'));
  if (!documentNode) {
    throw new Error('w:document ノードが見つかりません。');
  }

  const bodyNode = getDirectWordChild(documentNode, 'body');
  if (!bodyNode) {
    throw new Error('w:body ノードが見つかりません。');
  }

  return bodyNode;
}

function getBodyParagraphs(bodyNode) {
  return Array.from(bodyNode.getElementsByTagNameNS(W_NS, 'p'))
    .filter((paragraph) => !hasAncestor(paragraph, 'tbl'));
}

function getParagraphText(paragraph) {
  return Array.from(paragraph.getElementsByTagNameNS(W_NS, 't'))
    .map((node) => node.textContent || '')
    .join('');
}

function cloneFirstRunProperties(paragraph) {
  const firstRun = paragraph.getElementsByTagNameNS(W_NS, 'r')[0];
  if (!firstRun) {
    return null;
  }

  const runProps = getDirectWordChild(firstRun, 'rPr');
  return runProps ? runProps.cloneNode(true) : null;
}

function setParagraphText(paragraph, text) {
  const runProps = cloneFirstRunProperties(paragraph);
  const paragraphProps = getDirectWordChild(paragraph, 'pPr');

  for (const child of Array.from(paragraph.childNodes)) {
    if (child !== paragraphProps) {
      paragraph.removeChild(child);
    }
  }

  const run = createWordElement(paragraph.ownerDocument, 'r');
  if (runProps) {
    run.appendChild(runProps);
  }

  const textNode = createWordElement(paragraph.ownerDocument, 't');
  textNode.setAttributeNS(XML_NS, 'xml:space', 'preserve');
  textNode.textContent = text;
  run.appendChild(textNode);
  paragraph.appendChild(run);
}

function ensureParagraphProperties(paragraph) {
  return ensureWordChild(paragraph, 'pPr', { prepend: true });
}

function ensureRunProperties(run) {
  return ensureWordChild(run, 'rPr', { prepend: true });
}

function ensureRunFonts(runProperties) {
  return ensureWordChild(runProperties, 'rFonts', { prepend: true });
}

function ensureRunProperty(runProperties, localName) {
  return ensureWordChild(runProperties, localName);
}

function setRunFont(run, size) {
  const runProperties = ensureRunProperties(run);
  const runFonts = ensureRunFonts(runProperties);
  setWordAttr(runFonts, 'eastAsia', 'ＭＳ 明朝');
  setWordAttr(runFonts, 'ascii', 'Times New Roman');
  setWordAttr(runFonts, 'hAnsi', 'Times New Roman');

  const sizeNode = ensureRunProperty(runProperties, 'sz');
  setWordAttr(sizeNode, 'val', size * 2);

  const complexSizeNode = ensureRunProperty(runProperties, 'szCs');
  setWordAttr(complexSizeNode, 'val', size * 2);
}

function setRunBold(run, value) {
  const runProperties = ensureRunProperties(run);

  const normalBold = ensureRunProperty(runProperties, 'b');
  setWordAttr(normalBold, 'val', value ? 1 : 0);

  const complexBold = ensureRunProperty(runProperties, 'bCs');
  setWordAttr(complexBold, 'val', value ? 1 : 0);
}

function getParagraphRuns(paragraph) {
  return Array.from(paragraph.getElementsByTagNameNS(W_NS, 'r'));
}

function setParagraphFont(paragraph, size) {
  for (const run of getParagraphRuns(paragraph)) {
    setRunFont(run, size);
  }
}

function setParagraphBold(paragraph, value) {
  for (const run of getParagraphRuns(paragraph)) {
    setRunBold(run, value);
  }
}

function getParagraphAlignment(paragraph) {
  const paragraphProps = getDirectWordChild(paragraph, 'pPr');
  if (!paragraphProps) {
    return '';
  }

  const alignmentNode = getDirectWordChild(paragraphProps, 'jc');
  return alignmentNode ? getWordAttr(alignmentNode, 'val') : '';
}

function setParagraphAlignment(paragraph, value) {
  const paragraphProps = ensureParagraphProperties(paragraph);
  removeDirectWordChildren(paragraphProps, 'jc');
  if (!value) {
    return;
  }

  const alignmentNode = createWordElement(paragraph.ownerDocument, 'jc');
  setWordAttr(alignmentNode, 'val', value);
  paragraphProps.appendChild(alignmentNode);
}

function clearParagraphIndent(paragraph) {
  const paragraphProps = ensureParagraphProperties(paragraph);
  removeDirectWordChildren(paragraphProps, 'ind');
}

// ------------------------------------------------------------
// 段落スタイル方式（裁判L1〜L5 / 裁判本文L1〜L5）
// ------------------------------------------------------------

// styleId は ASCII（saibanL1 / saibanBodyL1）。w:name に日本語（裁判L1 / 裁判本文L1）。
function headingStyleId(level) {
  const lv = Math.min(Math.max(level, 1), STYLE_MAX_LEVEL);
  return `saibanL${lv}`;
}

function headingStyleName(level) {
  const lv = Math.min(Math.max(level, 1), STYLE_MAX_LEVEL);
  return `裁判L${lv}`;
}

function bodyStyleId(level) {
  const lv = level < 1 ? 1 : Math.min(level, STYLE_MAX_LEVEL);
  return `saibanBodyL${lv}`;
}

function bodyStyleName(level) {
  const lv = level < 1 ? 1 : Math.min(level, STYLE_MAX_LEVEL);
  return `裁判本文L${lv}`;
}

function cmToTwips(cm) {
  return Math.round(cm * CM_TO_TWIPS);
}

// styles.xml の <w:styles> 要素を取得する。part が無ければ生成して取得する。
// 戻り値: 取得/生成した <w:styles> 要素。
function getOrCreateStylesRoot(packageDoc) {
  const pkgRoot = packageDoc.documentElement; // <pkg:package>

  // 既存の /word/styles.xml part を探す
  let stylesPart = Array.from(packageDoc.getElementsByTagNameNS(PKG_NS, 'part'))
    .find((candidate) => getPackageAttr(candidate, 'name') === '/word/styles.xml');

  if (stylesPart) {
    const xmlData = Array.from(stylesPart.childNodes)
      .find((node) => node.nodeType === 1 && node.namespaceURI === PKG_NS && node.localName === 'xmlData');
    if (xmlData) {
      const stylesEl = Array.from(xmlData.childNodes)
        .find((node) => isWordElement(node, 'styles'));
      if (stylesEl) {
        return stylesEl;
      }
      // xmlData はあるが w:styles が無い異常系 → 作る
      const created = createWordElement(packageDoc, 'styles');
      xmlData.appendChild(created);
      return created;
    }
  }

  // part 自体が無い → part + xmlData + w:styles を生成して package に追加
  stylesPart = packageDoc.createElementNS(PKG_NS, 'pkg:part');
  stylesPart.setAttributeNS(PKG_NS, 'pkg:name', '/word/styles.xml');
  stylesPart.setAttributeNS(
    PKG_NS,
    'pkg:contentType',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
  );

  const xmlData = packageDoc.createElementNS(PKG_NS, 'pkg:xmlData');
  const stylesEl = createWordElement(packageDoc, 'styles');
  xmlData.appendChild(stylesEl);
  stylesPart.appendChild(xmlData);
  pkgRoot.appendChild(stylesPart);

  ensureStylesRelationship(packageDoc);

  return stylesEl;
}

// /word/_rels/document.xml.rels に styles.xml への関係を追加（無ければ）。
// getOoxml() の返すパッケージには通常 styles part が無いため、生成時に rel も張る。
function ensureStylesRelationship(packageDoc) {
  const relsPart = Array.from(packageDoc.getElementsByTagNameNS(PKG_NS, 'part'))
    .find((candidate) => getPackageAttr(candidate, 'name') === '/word/_rels/document.xml.rels');
  if (!relsPart) {
    return; // rels part 自体が無ければ何もしない（Word 側が補完する）
  }

  const xmlData = Array.from(relsPart.childNodes)
    .find((node) => node.nodeType === 1 && node.namespaceURI === PKG_NS && node.localName === 'xmlData');
  if (!xmlData) {
    return;
  }

  const relationships = Array.from(xmlData.childNodes)
    .find((node) => node.nodeType === 1 && node.localName === 'Relationships');
  if (!relationships) {
    return;
  }

  const STYLE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
  const hasStylesRel = Array.from(relationships.childNodes)
    .some((node) => node.nodeType === 1
      && node.localName === 'Relationship'
      && (node.getAttribute('Type') === STYLE_REL_TYPE
        || node.getAttribute('Target') === 'styles.xml'));
  if (hasStylesRel) {
    return;
  }

  // 既存の rId と衝突しない ID を採番
  const usedIds = new Set(
    Array.from(relationships.childNodes)
      .filter((node) => node.nodeType === 1 && node.localName === 'Relationship')
      .map((node) => node.getAttribute('Id')),
  );
  let n = 1;
  while (usedIds.has(`rId${n}`)) {
    n += 1;
  }

  const rel = packageDoc.createElementNS(PKG_REL_NS, 'Relationship');
  rel.setAttribute('Id', `rId${n}`);
  rel.setAttribute('Type', STYLE_REL_TYPE);
  rel.setAttribute('Target', 'styles.xml');
  relationships.appendChild(rel);
}

// 1スタイル要素 <w:style w:type="paragraph"> を構築して返す。
function buildCourtStyle(packageDoc, { styleId, name, leftTwips, hangingTwips = 0, firstLineTwips = 0, outlineLevel = null }) {
  const style = createWordElement(packageDoc, 'style');
  setWordAttr(style, 'type', 'paragraph');
  setWordAttr(style, 'styleId', styleId);

  const nameEl = createWordElement(packageDoc, 'name');
  setWordAttr(nameEl, 'val', name);
  style.appendChild(nameEl);

  // Normal 継承（フォント・サイズは標準を崩さない）
  const basedOn = createWordElement(packageDoc, 'basedOn');
  setWordAttr(basedOn, 'val', 'Normal');
  style.appendChild(basedOn);

  const pPr = createWordElement(packageDoc, 'pPr');

  const ind = createWordElement(packageDoc, 'ind');
  setWordAttr(ind, 'left', Math.round(leftTwips));
  if (hangingTwips > 0) {
    setWordAttr(ind, 'hanging', Math.round(hangingTwips));
  } else if (firstLineTwips > 0) {
    setWordAttr(ind, 'firstLine', Math.round(firstLineTwips));
  }
  pPr.appendChild(ind);

  if (outlineLevel !== null) {
    const olvl = createWordElement(packageDoc, 'outlineLvl');
    setWordAttr(olvl, 'val', outlineLevel);
    pPr.appendChild(olvl);
  }

  style.appendChild(pPr);
  return style;
}

// パッケージに裁判書式用の段落スタイル（裁判L1〜L5 / 裁判本文L1〜L5）を冪等に定義する。
// 既に同 styleId が存在すれば再定義しない。インデントはスタイルの pPr に持たせるため、
// Word 上でスタイル定義を1箇所変えればその階層の全段落が一括で動く。
// フォント等は Normal を継承し、明朝・本文サイズを崩さない。平野晋テンプレート標準値。
function ensureCourtStyles(packageDoc) {
  const stylesRoot = getOrCreateStylesRoot(packageDoc);

  const existingIds = new Set(
    getDirectWordChildren(stylesRoot, 'style').map((s) => getWordAttr(s, 'styleId')),
  );

  for (let level = 1; level <= STYLE_MAX_LEVEL; level += 1) {
    const vals = HIRANO_STYLE_CM[level];

    // 見出しスタイル: 裁判L{level}（hanging + outlineLvl）
    const hId = headingStyleId(level);
    if (!existingIds.has(hId)) {
      stylesRoot.appendChild(buildCourtStyle(packageDoc, {
        styleId: hId,
        name: headingStyleName(level),
        leftTwips: cmToTwips(vals.headingLeft),
        hangingTwips: cmToTwips(vals.headingHanging),
        outlineLevel: level - 1,
      }));
      existingIds.add(hId);
    }

    // 本文スタイル: 裁判本文L{level}（firstLine）
    const bId = bodyStyleId(level);
    if (!existingIds.has(bId)) {
      stylesRoot.appendChild(buildCourtStyle(packageDoc, {
        styleId: bId,
        name: bodyStyleName(level),
        leftTwips: cmToTwips(vals.bodyLeft),
        firstLineTwips: cmToTwips(vals.bodyFirstLine),
      }));
      existingIds.add(bId);
    }
  }

  return stylesRoot;
}

// 段落の pPr に w:pStyle（styleId）を設定する。既存 pStyle は置換する。
function setParagraphStyle(paragraph, styleId) {
  const paragraphProps = ensureParagraphProperties(paragraph);
  removeDirectWordChildren(paragraphProps, 'pStyle');
  const pStyle = createWordElement(paragraph.ownerDocument, 'pStyle');
  setWordAttr(pStyle, 'val', styleId);
  // pStyle は pPr 先頭に置くのが OOXML 慣例
  paragraphProps.insertBefore(pStyle, paragraphProps.firstChild);
}

function setIndent(paragraph, { leftChars = 0, hangingChars = 0, firstLineChars = 0 } = {}) {
  clearParagraphIndent(paragraph);

  if (!leftChars && !hangingChars && !firstLineChars) {
    return;
  }

  const indentNode = createWordElement(paragraph.ownerDocument, 'ind');
  if (leftChars) {
    setWordAttr(indentNode, 'leftChars', leftChars * 100);
    setWordAttr(indentNode, 'left', leftChars * CHAR_TWIPS);
  }
  if (hangingChars) {
    setWordAttr(indentNode, 'hangingChars', hangingChars * 100);
    setWordAttr(indentNode, 'hanging', hangingChars * CHAR_TWIPS);
  } else if (firstLineChars) {
    setWordAttr(indentNode, 'firstLineChars', firstLineChars * 100);
    setWordAttr(indentNode, 'firstLine', firstLineChars * CHAR_TWIPS);
  }

  ensureParagraphProperties(paragraph).appendChild(indentNode);
}

function setOutlineLevel(paragraph, level) {
  const paragraphProps = ensureParagraphProperties(paragraph);
  removeDirectWordChildren(paragraphProps, 'outlineLvl');
  // outlineLvl: 0=Level1, 1=Level2, ... 8=本文
  const olvl = createWordElement(paragraph.ownerDocument, 'outlineLvl');
  setWordAttr(olvl, 'val', level - 1);
  paragraphProps.appendChild(olvl);
}

function setHeadingIndent(paragraph, level) {
  // USE_STYLE_MODE=true（既定）: 段落スタイル 裁判L{level} を割当 + 直接 w:ind を除去。
  //   outlineLvl はスタイル側にも持たせているが、スタイルを剥がしても Word 目次が
  //   崩れないよう段落にも残す（CLI版と同じ堅牢性方針）。
  if (USE_STYLE_MODE) {
    setOutlineLevel(paragraph, level);
    setParagraphStyle(paragraph, headingStyleId(level));
    clearParagraphIndent(paragraph);
    return;
  }
  setHeadingIndentLegacy(paragraph, level);
}

// 【従来方式】見出し段落のインデント＋アウトラインレベルを直接 w:ind で設定。
function setHeadingIndentLegacy(paragraph, level) {
  const leftChars = HEADING_LEVELS[level][0];
  setOutlineLevel(paragraph, level);

  // 番号部分を除いた本文の長さで判定
  const text = getParagraphText(paragraph).trim();
  const body = text.replace(
    /^[\s\u3000]*(第[１-９０-９\d]+|[１-９０-９\d]+|[(\uff08][１-９０-９\d]+[)\uff09]|[ア-ン]|[(\uff08][ｱ-ﾝ]+[)\uff09]|[ａ-ｚ]|[(\uff08][a-z]+[)\uff09])[\s\u3000]*/,
    '');

  if (body.length > 20) {
    // 本文兼用 → ぶら下げインデント
    const hanging = _HEADING_HANGING[level] || 1;
    if (level === 1) {
      setIndent(paragraph, { leftChars });
    } else {
      setIndent(paragraph, { leftChars, hangingChars: hanging });
    }
  } else {
    // 短い小タイトル → 左インデント0、首行1字下げ
    if (level === 1) {
      setIndent(paragraph, { leftChars: 0 });
    } else {
      setIndent(paragraph, { leftChars: 0, firstLineChars: 1 });
    }
  }
}

function setBodyIndent(paragraph, currentHeadingLevel) {
  // USE_STYLE_MODE=true（既定）: 段落スタイル 裁判本文L{level} を割当 + 直接 w:ind を除去。
  // level=0 や範囲外は 裁判本文L1 にフォールバック（bodyStyleId 内で処理）。
  if (USE_STYLE_MODE) {
    setParagraphStyle(paragraph, bodyStyleId(currentHeadingLevel));
    clearParagraphIndent(paragraph);
    return;
  }
  setBodyIndentLegacy(paragraph, currentHeadingLevel);
}

// 【従来方式】本文段落のインデントを直接 w:ind で設定。
function setBodyIndentLegacy(paragraph, currentHeadingLevel) {
  const [left, fl] = BODY_INDENT[currentHeadingLevel] || [0, 1];
  setIndent(paragraph, { leftChars: left, firstLineChars: fl });
}

function setListIndent(paragraph, currentHeadingLevel, markerLength) {
  // 箇条書き・番号リストはマーカー幅ぶら下げが段落ごとに変わるため、スタイルだけでは
  // 表現できない。スタイル方式でも本文スタイルを土台に割り当てつつ、ぶら下げ位置は
  // 直接 w:ind で補正する（マーカー幅 = markerLength 字）。
  if (USE_STYLE_MODE) {
    setParagraphStyle(paragraph, bodyStyleId(currentHeadingLevel));
  }
  const [bodyLeft] = BODY_INDENT[currentHeadingLevel] || [0, 0];
  setIndent(paragraph, {
    leftChars: bodyLeft + markerLength,
    hangingChars: markerLength,
  });
}

function ensureCellMargins(cellProperties) {
  removeDirectWordChildren(cellProperties, 'tcMar');

  const margins = createWordElement(cellProperties.ownerDocument, 'tcMar');
  const defs = [
    ['top', 0],
    ['left', 28],
    ['bottom', 0],
    ['right', 28],
  ];

  for (const [side, width] of defs) {
    const margin = createWordElement(cellProperties.ownerDocument, side);
    setWordAttr(margin, 'w', width);
    setWordAttr(margin, 'type', 'dxa');
    margins.appendChild(margin);
  }

  cellProperties.appendChild(margins);
}

function formatTables(bodyNode) {
  const tables = Array.from(bodyNode.getElementsByTagNameNS(W_NS, 'tbl'));
  for (const table of tables) {
    const tableProperties = ensureWordChild(table, 'tblPr', { prepend: true });
    removeDirectWordChildren(tableProperties, 'tblLayout');

    const layout = createWordElement(bodyNode.ownerDocument, 'tblLayout');
    setWordAttr(layout, 'type', 'autofit');
    tableProperties.appendChild(layout);

    const cells = Array.from(table.getElementsByTagNameNS(W_NS, 'tc'));
    for (const cell of cells) {
      const cellProperties = ensureWordChild(cell, 'tcPr', { prepend: true });
      ensureCellMargins(cellProperties);

      const paragraphs = Array.from(cell.getElementsByTagNameNS(W_NS, 'p'));
      for (const paragraph of paragraphs) {
        clearParagraphIndent(paragraph);
        setParagraphFont(paragraph, TABLE_FONT_SIZE);
      }
    }
  }
}

function detectLevelOffset(texts) {
  let inHeaderSection = true;
  const levels = [];

  for (const text of texts) {
    const trimmed = text.trim();
    if (!trimmed) {
      continue;
    }

    if (inHeaderSection) {
      const level = detectHeadingLevel(trimmed);
      if (level !== null) {
        inHeaderSection = false;
        levels.push(level);
      } else if (isHeaderSection(trimmed)) {
        continue;
      } else {
        continue;
      }
    } else {
      const level = detectHeadingLevel(trimmed);
      if (level !== null) {
        levels.push(level);
      }
    }
  }

  return levels.length > 0 ? Math.min(...levels) - 1 : 0;
}

function serializeOoxml(doc) {
  return new XMLSerializer().serializeToString(doc);
}

// 岡口マクロが設定するスタイル名パターン
const OKAGUCHI_STYLE_RE = /^(ランク[１-９1-9]|本文[１-９1-9]|標準\(太郎文書スタイル\))/;

function getParagraphStyleId(paragraph) {
  const pPr = getDirectWordChild(paragraph, 'pPr');
  if (!pPr) return '';
  const pStyle = getDirectWordChild(pPr, 'pStyle');
  if (!pStyle) return '';
  return getWordAttr(pStyle, 'val');
}

function hasOkaguchiStyles(paragraphs) {
  return paragraphs.some((p) => OKAGUCHI_STYLE_RE.test(getParagraphStyleId(p)));
}

function transformBodyOoxml(ooxml, options) {
  const { font: doFont, indent: doIndent, zenkaku: doZenkaku } = options;
  if (!doFont && !doIndent && !doZenkaku) {
    return ooxml;
  }

  const xmlDoc = parseOoxmlPackage(ooxml);
  const bodyNode = getBodyNode(xmlDoc);
  const bodyParagraphs = getBodyParagraphs(bodyNode);

  // 岡口マクロスタイル検出: インデント処理をスキップ
  const skipIndent = doIndent && hasOkaguchiStyles(bodyParagraphs);
  if (skipIndent) {
    // 全角変換とフォントのみ実行
    for (const paragraph of bodyParagraphs) {
      if (doZenkaku) {
        const text = getParagraphText(paragraph);
        if (text) {
          const converted = toZenkaku(text);
          if (converted !== text) {
            setParagraphText(paragraph, converted);
          }
        }
      }
      if (doFont) {
        setParagraphFont(paragraph, BODY_FONT_SIZE);
      }
    }
    if (doFont) {
      formatTables(bodyNode);
    }
    return serializeOoxml(xmlDoc);
  }

  const candidateTexts = bodyParagraphs.map((paragraph) => {
    const text = getParagraphText(paragraph);
    return doZenkaku && text ? toZenkaku(text) : text;
  });

  const levelOffset = doIndent ? detectLevelOffset(candidateTexts) : 0;

  // 段落スタイル方式: 裁判L1〜L5 / 裁判本文L1〜L5 を styles.xml に冪等定義する。
  // 各段落の setHeadingIndent/setBodyIndent が w:pStyle を割り当てる。
  // パッケージ全体を1回処理（styles part を生成 or 既存に追記）。
  if (doIndent && USE_STYLE_MODE) {
    ensureCourtStyles(xmlDoc);
  }

  let currentHeadingLevel = 0;
  let inHeaderSection = true;
  const counter = new Counter();

  for (const paragraph of bodyParagraphs) {
    let text = getParagraphText(paragraph);

    // 先頭全角スペース除去（冒頭セクション含む全段落）
    const strippedLeading = text.replace(/^[\u3000 \t]+/, '');
    if (strippedLeading !== text) {
      setParagraphText(paragraph, strippedLeading);
      text = strippedLeading;
    }

    if (doZenkaku && text) {
      const converted = toZenkaku(text);
      if (converted !== text) {
        setParagraphText(paragraph, converted);
        text = converted;
      }
    }

    const trimmed = text.trim();
    if (!trimmed) {
      if (doFont) {
        setParagraphFont(paragraph, BODY_FONT_SIZE);
      }
      continue;
    }

    if (doIndent) {
      if (inHeaderSection) {
        const firstLevel = detectHeadingLevel(trimmed);
        if (firstLevel !== null) {
          inHeaderSection = false;
        } else if (isHeaderSection(trimmed)) {
          if (TITLE_PATTERN.test(trimmed)) {
            setParagraphFont(paragraph, TITLE_FONT_SIZE);
            setParagraphBold(paragraph, true);
          } else if (doFont) {
            setParagraphFont(paragraph, BODY_FONT_SIZE);
          }
          continue;
        } else {
          if (doFont) {
            setParagraphFont(paragraph, BODY_FONT_SIZE);
          }
          continue;
        }
      }

      const level = detectHeadingLevel(trimmed);
      if (level !== null) {
        const adjusted = remap(level, levelOffset);
        currentHeadingLevel = adjusted;
        const bodyText = stripNum(text);
        const newText = genNum(adjusted, counter.inc(adjusted)) + bodyText;
        if (newText !== text) {
          setParagraphText(paragraph, newText);
          text = newText;
        }

        setHeadingIndent(paragraph, adjusted);
        setParagraphAlignment(paragraph, 'left');
      } else if (SKIP_PATTERN.test(trimmed)) {
        setParagraphAlignment(paragraph, 'right');
        clearParagraphIndent(paragraph);
      } else {
        const stripped = stripSpaces(text);
        if (stripped !== text) {
          setParagraphText(paragraph, stripped);
          text = stripped;
        }

        const listMatch = text.match(LIST_PATTERN);
        const bulletMatch = text.match(BULLET_PATTERN);
        if (listMatch) {
          setListIndent(paragraph, currentHeadingLevel, listMatch[0].length);
        } else if (bulletMatch) {
          setListIndent(paragraph, currentHeadingLevel, 1); // 記号1文字分
        } else {
          setBodyIndent(paragraph, currentHeadingLevel);
        }

        if (getParagraphAlignment(paragraph) === 'center') {
          setParagraphAlignment(paragraph, 'left');
        }
      }
    }

    if (doFont) {
      setParagraphFont(paragraph, BODY_FONT_SIZE);
    }
  }

  if (doFont) {
    formatTables(bodyNode);
  }

  return serializeOoxml(xmlDoc);
}

// ============================================================
// Office.js ラッパー
// ============================================================

async function applyPageSetup(context) {
  const sections = context.document.sections;
  sections.load('items');
  await context.sync();

  for (const section of sections.items) {
    const pageSetup = section.pageSetup;
    pageSetup.paperSize = Word.PaperSize.a4;
    pageSetup.topMargin = PAGE.topMargin;
    pageSetup.bottomMargin = PAGE.bottomMargin;
    pageSetup.leftMargin = PAGE.leftMargin;
    pageSetup.rightMargin = PAGE.rightMargin;
  }

  await context.sync();
}

async function setupFormat() {
  await Word.run(async (context) => {
    await applyPageSetup(context);
  });
}

async function replaceBodyWithTransformedOoxml(context, options) {
  const body = context.document.body;
  const bodyOoxml = body.getOoxml();
  await context.sync();

  const transformed = transformBodyOoxml(bodyOoxml.value, options);
  if (transformed !== bodyOoxml.value) {
    body.insertOoxml(transformed, Word.InsertLocation.replace);
    await context.sync();
  }
}

async function applyFooterPageNumbers(context) {
  const sections = context.document.sections;
  sections.load('items');
  await context.sync();

  for (const section of sections.items) {
    const footer = section.getFooter(Word.HeaderFooterType.primary);
    footer.clear();
    footer.insertOoxml(FOOTER_PAGE_NUMBER_OOXML, Word.InsertLocation.replace);
  }

  await context.sync();
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
    if (doPage) {
      await applyPageSetup(context);
    }

    if (doFont || doIndent || doZenkaku) {
      await replaceBodyWithTransformedOoxml(context, {
        font: doFont,
        indent: doIndent,
        zenkaku: doZenkaku,
      });
    }

    if (doFooter) {
      await applyFooterPageNumbers(context);
    }
  });
}

// ============================================================
// CommonJS エクスポート（ユニットテスト用）
// ブラウザ/Office.js 実行時は module が未定義のためこのブロックは無視される。
// 純粋な OOXML 変換ロジック（Word 非依存）だけをテストから参照できるようにする。
// ============================================================
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    // フラグ・定数
    USE_STYLE_MODE,
    HIRANO_STYLE_CM,
    CM_TO_TWIPS,
    STYLE_MAX_LEVEL,
    CHAR_TWIPS,
    // 変換エントリポイント
    transformBodyOoxml,
    // スタイル方式ヘルパー
    ensureCourtStyles,
    getOrCreateStylesRoot,
    setParagraphStyle,
    headingStyleId,
    headingStyleName,
    bodyStyleId,
    bodyStyleName,
    cmToTwips,
    // 既存の純ロジック（テスト補助）
    toZenkaku,
    detectHeadingLevel,
    setIndent,
    setHeadingIndent,
    setBodyIndent,
  };
}
