/**
 * 段落スタイル方式（裁判L1〜L5 / 裁判本文L1〜L5）のユニットテスト。
 *
 * Office.js（getOoxml/insertOoxml/Word.run）は Word 実機が必要なためテスト対象外。
 * ここでは Word 非依存の純粋な OOXML 変換ロジック（DOM 操作）だけを検証する：
 *   (a) styles.xml に 10 スタイル定義が冪等に追加される
 *   (b) 見出し段落の pPr に w:pStyle が付く
 *   (c) スタイルを割り当てた段落に直接 w:ind が残らない
 *
 * DOMParser / XMLSerializer は @xmldom/xmldom をグローバルに注入して再現する
 * （converter.js はブラウザ/Office.js のグローバル API を前提にしているため）。
 */

const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');

global.DOMParser = DOMParser;
global.XMLSerializer = XMLSerializer;

const converter = require('../src/converter.js');

const {
  transformBodyOoxml,
  ensureCourtStyles,
  cmToTwips,
  HIRANO_STYLE_CM,
  STYLE_MAX_LEVEL,
  USE_STYLE_MODE,
} = converter;

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PKG_NS = 'http://schemas.microsoft.com/office/2006/xmlPackage';

// ------------------------------------------------------------
// テスト用ヘルパー
// ------------------------------------------------------------

// getOoxml() が返すような最小パッケージ（document.xml + 各 _rels、styles.xml は無い）。
// 段落:
//   p0 = 見出しL1（既存の直接 w:ind を含む → 除去されるべき）
//   p1 = 見出しL2
//   p2 = 本文（直前見出しL2 → 裁判本文L2 が付くべき）
//   p3 = 見出しL3（カッコ番号）
function buildPackage() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="${PKG_NS}">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="${W_NS}">
        <w:body>
          <w:p><w:pPr><w:ind w:left="999" w:leftChars="400"/></w:pPr><w:r><w:t>第１　総論</w:t></w:r></w:p>
          <w:p><w:r><w:t>１　事実関係について述べる。</w:t></w:r></w:p>
          <w:p><w:r><w:t>本文の段落である。これは見出しの直後の本文。</w:t></w:r></w:p>
          <w:p><w:r><w:t>(1)　細目について論じる。</w:t></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
}

function parse(xml) {
  return new DOMParser().parseFromString(xml, 'application/xml');
}

function styleElements(doc) {
  return Array.from(doc.getElementsByTagNameNS(W_NS, 'style'));
}

function attr(el, localName) {
  if (!el) return '';
  return el.getAttributeNS(W_NS, localName) || el.getAttribute(`w:${localName}`) || '';
}

function styleId(styleEl) {
  return attr(styleEl, 'styleId');
}

// body 直下（テーブル外）の段落だけを返す。styles.xml 内の style/pPr は除外。
function bodyParagraphs(doc) {
  return Array.from(doc.getElementsByTagNameNS(W_NS, 'p')).filter((p) => {
    let n = p.parentNode;
    while (n) {
      if (n.localName === 'body') return true;
      n = n.parentNode;
    }
    return false;
  });
}

function directChild(parent, localName) {
  if (!parent) return null;
  return Array.from(parent.childNodes).find((n) => n.localName === localName) || null;
}

function paragraphPStyle(paragraph) {
  const pPr = directChild(paragraph, 'pPr');
  const pStyle = directChild(pPr, 'pStyle');
  return pStyle ? attr(pStyle, 'val') : '';
}

function paragraphDirectInd(paragraph) {
  const pPr = directChild(paragraph, 'pPr');
  return directChild(pPr, 'ind');
}

function paragraphText(paragraph) {
  return Array.from(paragraph.getElementsByTagNameNS(W_NS, 't'))
    .map((t) => t.textContent || '')
    .join('');
}

// ------------------------------------------------------------
// (a) styles.xml に 10 スタイルが定義される
// ------------------------------------------------------------

describe('段落スタイル方式: styles.xml への 10 スタイル定義', () => {
  test('USE_STYLE_MODE デフォルトは true', () => {
    expect(USE_STYLE_MODE).toBe(true);
  });

  test('変換後パッケージに 裁判L1〜L5 / 裁判本文L1〜L5 の 10 スタイルが揃う', () => {
    const out = transformBodyOoxml(buildPackage(), { font: false, indent: true, zenkaku: false });
    const doc = parse(out);
    const ids = styleElements(doc).map(styleId).sort();

    const expected = [];
    for (let lv = 1; lv <= STYLE_MAX_LEVEL; lv += 1) {
      expected.push(`saibanL${lv}`, `saibanBodyL${lv}`);
    }
    expect(ids).toEqual(expected.sort());
    expect(ids).toHaveLength(10);
  });

  test('styles.xml part が無いパッケージでも part を生成する', () => {
    // 入力に /word/styles.xml part は存在しない
    const inputDoc = parse(buildPackage());
    const inputHasStyles = Array.from(inputDoc.getElementsByTagNameNS(PKG_NS, 'part'))
      .some((p) => (p.getAttributeNS(PKG_NS, 'name') || p.getAttribute('pkg:name')) === '/word/styles.xml');
    expect(inputHasStyles).toBe(false);

    const out = transformBodyOoxml(buildPackage(), { font: false, indent: true, zenkaku: false });
    const doc = parse(out);
    const outHasStyles = Array.from(doc.getElementsByTagNameNS(PKG_NS, 'part'))
      .some((p) => (p.getAttributeNS(PKG_NS, 'name') || p.getAttribute('pkg:name')) === '/word/styles.xml');
    expect(outHasStyles).toBe(true);
  });

  test('ensureCourtStyles は冪等（2回呼んでも 10 のまま、重複しない）', () => {
    const doc = parse(buildPackage());
    ensureCourtStyles(doc);
    ensureCourtStyles(doc);
    expect(styleElements(doc)).toHaveLength(10);
  });

  test('各スタイルの w:ind が平野晋標準値（cm→twips）と一致する', () => {
    const doc = parse(buildPackage());
    ensureCourtStyles(doc);
    const byId = {};
    styleElements(doc).forEach((s) => { byId[styleId(s)] = s; });

    for (let lv = 1; lv <= STYLE_MAX_LEVEL; lv += 1) {
      const vals = HIRANO_STYLE_CM[lv];

      // 見出し: left + hanging + outlineLvl(lv-1)
      const h = byId[`saibanL${lv}`];
      const hInd = h.getElementsByTagNameNS(W_NS, 'ind')[0];
      expect(attr(hInd, 'left')).toBe(String(cmToTwips(vals.headingLeft)));
      expect(attr(hInd, 'hanging')).toBe(String(cmToTwips(vals.headingHanging)));
      const hOlvl = h.getElementsByTagNameNS(W_NS, 'outlineLvl')[0];
      expect(attr(hOlvl, 'val')).toBe(String(lv - 1));

      // 本文: left + firstLine、outlineLvl は付かない
      const b = byId[`saibanBodyL${lv}`];
      const bInd = b.getElementsByTagNameNS(W_NS, 'ind')[0];
      expect(attr(bInd, 'left')).toBe(String(cmToTwips(vals.bodyLeft)));
      expect(attr(bInd, 'firstLine')).toBe(String(cmToTwips(vals.bodyFirstLine)));
      expect(b.getElementsByTagNameNS(W_NS, 'outlineLvl')).toHaveLength(0);
    }
  });

  test('見出しスタイルは Normal を継承し eastAsia フォントを上書きしない', () => {
    const doc = parse(buildPackage());
    ensureCourtStyles(doc);
    const s = styleElements(doc).find((x) => styleId(x) === 'saibanL1');
    const basedOn = s.getElementsByTagNameNS(W_NS, 'basedOn')[0];
    expect(attr(basedOn, 'val')).toBe('Normal');
    // フォント（rPr/rFonts）はスタイルに書かない＝本文の明朝・サイズを壊さない
    expect(s.getElementsByTagNameNS(W_NS, 'rFonts')).toHaveLength(0);
  });
});

// ------------------------------------------------------------
// (b) 見出し段落 pPr に w:pStyle が付く
// ------------------------------------------------------------

describe('段落スタイル方式: 見出し/本文段落への w:pStyle 割当', () => {
  test('見出し段落に対応する裁判L{n}スタイルが pStyle で付く', () => {
    const out = transformBodyOoxml(buildPackage(), { font: false, indent: true, zenkaku: false });
    const doc = parse(out);
    const paras = bodyParagraphs(doc);

    // p0=第１→saibanL1, p1=１→saibanL2, p3=(1)→saibanL3
    const byText = {};
    paras.forEach((p) => { byText[paragraphText(p)] = p; });

    const h1 = paras.find((p) => paragraphText(p).includes('総論'));
    const h2 = paras.find((p) => paragraphText(p).includes('事実関係'));
    const h3 = paras.find((p) => paragraphText(p).includes('細目'));

    expect(paragraphPStyle(h1)).toBe('saibanL1');
    expect(paragraphPStyle(h2)).toBe('saibanL2');
    expect(paragraphPStyle(h3)).toBe('saibanL3');
  });

  test('本文段落には直前見出しレベルの裁判本文L{n}が付く', () => {
    const out = transformBodyOoxml(buildPackage(), { font: false, indent: true, zenkaku: false });
    const doc = parse(out);
    const body = bodyParagraphs(doc).find((p) => paragraphText(p).includes('本文の段落'));
    // 直前見出しは「１」= level2 → 裁判本文L2
    expect(paragraphPStyle(body)).toBe('saibanBodyL2');
  });

  test('見出し段落には outlineLvl も段落側に残る（目次堅牢性）', () => {
    const out = transformBodyOoxml(buildPackage(), { font: false, indent: true, zenkaku: false });
    const doc = parse(out);
    const h1 = bodyParagraphs(doc).find((p) => paragraphText(p).includes('総論'));
    const pPr = directChild(h1, 'pPr');
    const olvl = directChild(pPr, 'outlineLvl');
    expect(olvl).not.toBeNull();
    expect(attr(olvl, 'val')).toBe('0');
  });
});

// ------------------------------------------------------------
// (c) スタイル割当段落に直接 w:ind が残らない
// ------------------------------------------------------------

describe('段落スタイル方式: 直接 w:ind を書かない', () => {
  test('スタイルを割り当てた全段落に直接 w:ind が無い', () => {
    const out = transformBodyOoxml(buildPackage(), { font: false, indent: true, zenkaku: false });
    const doc = parse(out);
    const paras = bodyParagraphs(doc);

    for (const p of paras) {
      const pStyle = paragraphPStyle(p);
      if (pStyle && pStyle.startsWith('saiban')) {
        expect(paragraphDirectInd(p)).toBeNull();
      }
    }
  });

  test('入力に存在した直接 w:ind（第１段落の w:left=999）が除去される', () => {
    // 入力では p0 に w:ind w:left="999" がある
    const inputDoc = parse(buildPackage());
    const inP0 = bodyParagraphs(inputDoc)[0];
    expect(paragraphDirectInd(inP0)).not.toBeNull();

    const out = transformBodyOoxml(buildPackage(), { font: false, indent: true, zenkaku: false });
    const doc = parse(out);
    const h1 = bodyParagraphs(doc).find((p) => paragraphText(p).includes('総論'));
    expect(paragraphPStyle(h1)).toBe('saibanL1');
    expect(paragraphDirectInd(h1)).toBeNull();
  });
});

// ------------------------------------------------------------
// 既存機能の非破壊（全角変換が同時に効くか）
// ------------------------------------------------------------

describe('既存機能との両立', () => {
  test('スタイル方式でも全角変換が同時に効く', () => {
    // 半角数字を含む見出し → 全角化 + スタイル割当
    const pkg = `<pkg:package xmlns:pkg="${PKG_NS}"><pkg:part pkg:name="/word/document.xml" pkg:contentType="x"><pkg:xmlData><w:document xmlns:w="${W_NS}"><w:body><w:p><w:r><w:t>abc123</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
    const out = transformBodyOoxml(pkg, { font: false, indent: false, zenkaku: true });
    expect(out).toContain('ａｂｃ１２３');
  });
});
