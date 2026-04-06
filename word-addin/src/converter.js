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

const CHAR_WIDTH = 242; // twips (12pt MS明朝 + グリッド補正)
const BODY_FONT_SIZE = 12;
const TABLE_FONT_SIZE = 10;
const TITLE_FONT_SIZE = 16;

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PKG_NS = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const XML_NS = 'http://www.w3.org/XML/1998/namespace';

// 1全角文字幅 (twips)
const _F = 242; // full-width char

// 1全角文字幅 (twips)
// const _F = 242; // already defined above

// 見出しレベルごとの左インデント (chars単位)
// 岡口マクロVBAソース準拠: https://www.slaw.jp/2015/08/ms-wordvba.html
const HEADING_LEFT_CHARS = {
  1: 2,  // 第１ — left=24pt(2字)
  2: 2,  // １   — left=24pt(2字)
  3: 3,  // (1)  — left=36pt(3字)
  4: 4,  // ア   — left=48pt(4字)
  5: 5,  // (ｱ)  — left=60pt(5字)
  6: 6,  // ａ   — left=72pt(6字)
  7: 7,  // (a)  — left=84pt(7字)
};

// ぶら下げ幅 (chars単位)
// 全レベル1字、第１だけ2字（VBAソース: FirstLineIndent=-12pt, 第１のみ-24pt）
const HEADING_HANGING_CHARS = {
  1: 2,  // 第１ → 2字（left2-hang2=0字目から）
  2: 1,  // １　 → 1字（left2-hang1=1字目から）
  3: 1,  // (1)  → 1字（left3-hang1=2字目から）
  4: 1,  // ア　 → 1字（left4-hang1=3字目から）
  5: 1,  // (ｱ)  → 1字（left5-hang1=4字目から）
  6: 1,  // ａ　 → 1字（left6-hang1=5字目から）
  7: 1,  // (a)  → 1字（left7-hang1=6字目から）
};

// 本文インデント: [left_chars, firstLine_chars]
// 番号＋全角スペースの直後から本文開始。番号の下には文字が来ない。
// 1行目も2行目も同じ位置（firstLine=0）。
const BODY_INDENT = {
  0: [0, 1],    // 見出しなし → 首行1字のみ
  1: [3, 0],    // 第１直下 → 3字目から
  2: [3, 0],    // １直下 → 3字目から
  3: [6, 0],    // (1)直下 → 6字目から
  4: [5, 0],    // ア直下 → 5字目から
  5: [8, 0],    // (ｱ)直下 → 8字目から
  6: [7, 0],    // ａ直下 → 7字目から
  7: [10, 0],   // (a)直下 → 10字目から
};

const TITLE_PATTERN = /(準備書面|訴状|答弁書|意見書|報告書|申立書|陳述書|上申書|申請書|請求書|通知書|催告書|告訴状|告発状|嘆願書|抗告理由書|控訴理由書|上告理由書|証拠保全請求書|接見禁止.{0,4}申請書|押収物還付請求書|勾留.{0,6}請求書|保釈請求書|弁論要旨)/;
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

function normalizeHeadingSpacing(text) {
  // 見出し番号の後のスペースを全角スペース1個に正規化
  const ZS = '\u3000';
  const patterns = [
    /^([\s\u3000]*[(\uff08][１２３４５６７８９０\d]+[)\uff09])[\s\u3000]*(.*)/s,
    /^([\s\u3000]*[(\uff08][ｱ-ﾝア-ン]+[)\uff09])[\s\u3000]*(.*)/s,
    /^([\s\u3000]*[(\uff08][a-zａ-ｚ]+[)\uff09])[\s\u3000]*(.*)/s,
    /^([\s\u3000]*第[１２３４５６７８９０\d]+)[\s\u3000]*(.*)/s,
    /^([\s\u3000]*[１２３４５６７８９０\d]+)[\s\u3000]*(.*)/s,
    /^([\s\u3000]*[ア-ン])[\s\u3000]*(.*)/s,
    /^([\s\u3000]*[ａ-ｚ])[\s\u3000]*(.*)/s,
  ];
  for (const re of patterns) {
    const m = text.match(re);
    if (m) {
      return m[1] + ZS + m[2];
    }
  }
  return text;
}

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

function setParagraphIndent(paragraph, { leftTwips = 0, hangingTwips = 0, firstLineTwips = 0 } = {}) {
  clearParagraphIndent(paragraph);

  if (!leftTwips && !hangingTwips && !firstLineTwips) {
    return;
  }

  const indentNode = createWordElement(paragraph.ownerDocument, 'ind');
  if (leftTwips) {
    setWordAttr(indentNode, 'left', leftTwips);
  }
  if (hangingTwips) {
    setWordAttr(indentNode, 'hanging', hangingTwips);
  }
  if (firstLineTwips) {
    setWordAttr(indentNode, 'firstLine', firstLineTwips);
  }

  ensureParagraphProperties(paragraph).appendChild(indentNode);
}

function setHeadingIndent(paragraph, level, text) {
  // 岡口マクロVBAソース準拠: 全レベルでぶら下げインデント
  const leftChars = HEADING_LEFT_CHARS[level] || 0;
  const hangChars = HEADING_HANGING_CHARS[level] || 1;
  setParagraphIndent(paragraph, {
    leftTwips: leftChars * _F,
    hangingTwips: hangChars * _F,
  });
}

function setBodyIndent(paragraph, currentHeadingLevel) {
  const [leftChars, firstLineChars] = BODY_INDENT[currentHeadingLevel] || [0, 1];
  setParagraphIndent(paragraph, {
    leftTwips: leftChars * _F,
    firstLineTwips: firstLineChars * _F,
  });
}

function setListIndent(paragraph, currentHeadingLevel, markerLength) {
  const [bodyLeft] = BODY_INDENT[currentHeadingLevel] || [0, 0];
  const markerTwips = markerLength * _F;
  setParagraphIndent(paragraph, {
    leftTwips: bodyLeft + markerTwips,
    hangingTwips: markerTwips,
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

function transformBodyOoxml(ooxml, options) {
  const { font: doFont, indent: doIndent, zenkaku: doZenkaku } = options;
  if (!doFont && !doIndent && !doZenkaku) {
    return ooxml;
  }

  const xmlDoc = parseOoxmlPackage(ooxml);
  const bodyNode = getBodyNode(xmlDoc);
  const bodyParagraphs = getBodyParagraphs(bodyNode);

  const candidateTexts = bodyParagraphs.map((paragraph) => {
    const text = getParagraphText(paragraph);
    return doZenkaku && text ? toZenkaku(text) : text;
  });

  const levelOffset = doIndent ? detectLevelOffset(candidateTexts) : 0;
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
        let newText = genNum(adjusted, counter.inc(adjusted)) + bodyText;
        // 番号後のスペースを全角1個に正規化
        newText = normalizeHeadingSpacing(newText);
        if (newText !== text) {
          setParagraphText(paragraph, newText);
          text = newText;
        }

        setHeadingIndent(paragraph, adjusted, text);
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

  // insertOoxmlの後に明示的に中央揃えを設定
  for (const section of sections.items) {
    const footer = section.getFooter(Word.HeaderFooterType.primary);
    const paras = footer.paragraphs;
    paras.load('items');
    await context.sync();
    for (const p of paras.items) {
      p.alignment = Word.Alignment.centered;
    }
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
