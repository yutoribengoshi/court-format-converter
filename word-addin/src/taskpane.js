/**
 * taskpane.js — Word アドイン UI制御
 */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById('btnConvert').addEventListener('click', onConvert);
  }
});

async function onConvert() {
  const btn = document.getElementById('btnConvert');
  const status = document.getElementById('status');

  // オプション取得
  const options = {
    page: document.getElementById('optPage').checked,
    font: document.getElementById('optFont').checked,
    indent: document.getElementById('optIndent').checked,
    zenkaku: document.getElementById('optZenkaku').checked,
    footer: document.getElementById('optFooter').checked,
  };

  // 何も選択されていない場合
  if (!options.page && !options.font && !options.indent && !options.zenkaku && !options.footer) {
    status.textContent = '変換項目を1つ以上選択してください。';
    status.className = 'error';
    return;
  }

  btn.disabled = true;
  btn.textContent = '変換中...';
  status.textContent = '';
  status.className = '';

  try {
    await convertDocument(options);

    const items = [];
    if (options.page) items.push('ページ設定');
    if (options.font) items.push('フォント');
    if (options.indent) items.push('インデント');
    if (options.zenkaku) items.push('全角変換');
    if (options.footer) items.push('ページ番号');

    status.textContent = `変換完了: ${items.join('・')}`;
    status.className = 'success';
  } catch (error) {
    console.error('変換エラー:', error);
    status.textContent = `エラー: ${error.message}`;
    status.className = 'error';
  } finally {
    btn.disabled = false;
    btn.textContent = '変換を実行';
  }
}
