// GAS(gas/Code.gs)のロジックを、Google側をモックにして確認するテスト。
//   使い方: node gas/test/run.js     (Node.js のみ。外部ライブラリ・APIキーは不要)
//   確認する内容: 登録時の重複/保留の判定 / 取り込み / 「取り込みを進める」の段階判定 / 作業シートの後片付け
'use strict';
const fs = require('fs');
const path = require('path');
const assert = require('assert');

const SRC = fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8').replace(/^const /gm, 'var ');

// ---- シートのモック(1始まりの2次元配列) ----
class Sheet {
  constructor(name, rows) { this.name = name; this.g = rows.map(r => r.slice()); }
  getName() { return this.name; }
  getLastRow() { return this.g.length; }
  getMaxRows() { return Math.max(this.g.length, 1000); }
  getMaxColumns() { return 26; }
  getLastColumn() { return Math.max(1, ...this.g.map(r => r.length)); }
  insertColumnsAfter() {}
  setFrozenRows() {}
  clearContents() { this.g = []; }
  getRange(r, c, nr = 1, nc = 1) {
    const t = this;
    return {
      getValues() {
        const out = [];
        for (let i = 0; i < nr; i++) {
          const row = [];
          for (let j = 0; j < nc; j++) {
            const v = (t.g[r - 1 + i] || [])[c - 1 + j];
            row.push(v === undefined ? '' : v);
          }
          out.push(row);
        }
        return out;
      },
      getValue() { const v = (t.g[r - 1] || [])[c - 1]; return v === undefined ? '' : v; },
      setValues(a) {
        a.forEach((row, i) => row.forEach((v, j) => {
          while (t.g.length < r + i) t.g.push([]);
          t.g[r - 1 + i][c - 1 + j] = v;
        }));
      },
      setValue(v) { while (t.g.length < r) t.g.push([]); t.g[r - 1][c - 1] = v; },
      clearContent() { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) if (t.g[r - 1 + i]) t.g[r - 1 + i][c - 1 + j] = ''; },
      setDataValidation() {}
    };
  }
}

const H = ['type', 'title', 'img', 'category', 'comment', 'website', 'address', 'hide', 'map_url', 'age_limit', 'gtypes', 'status', 'lat', 'lng'];
const HS = H.concat(['登録状態']);
const IH = ['タイトル', 'メモ', 'URL', 'タグ', 'コメント', '取込状態'];
const mk = (o) => {
  const a = new Array(15).fill('');
  for (const k of Object.keys(o)) a[k === 'reg' ? 14 : H.indexOf(k)] = o[k];
  return a;
};
const fid = (x, name = 'n') => `https://www.google.com/maps/place/${name}/data=!4m2!3m1!1s0x1:0x${x}`;
const dLat = (m) => m / 111320; // メートル → 緯度の度数

// ---- GAS 環境をモックにして Code.gs を読み込む ----
function load(sheets, ui = {}) {
  const log = [];
  const env = {
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({
        getSheetByName: n => sheets[n] || null,
        getActiveSheet: () => { const sh = sheets[ui.active.sheet]; sh.getActiveRange = () => ({ getRow: () => ui.active.row }); return sh; },
        insertSheet: n => (sheets[n] = new Sheet(n, [])),
        toast: m => log.push('toast:' + m)
      }),
      flush() {},
      newDataValidation: () => ({ requireValueInList: () => ({ build: () => ({}) }) }),
      getUi: () => ({
        alert: (t) => { log.push('alert:' + t); return ui.answer || 'YES'; },
        showModalDialog: (out, title) => { log.push('dialog:' + title); log.html = out.html; },
        ButtonSet: { YES_NO: 1, OK: 0 }, Button: { YES: 'YES' }
      })
    },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k === 'MAPS_API_KEY' ? 'K' : null), setProperty() {}, deleteProperty() {} }) },
    ScriptApp: { getProjectTriggers: () => [], deleteTrigger() {}, newTrigger: () => ({ timeBased: () => ({ after: () => ({ create() {} }) }) }) },
    Utilities: { sleep() {} },
    UrlFetchApp: { fetch: (u, o) => (ui.fetch ? ui.fetch(u, o) : { getResponseCode: () => 200, getHeaders: () => ({}), getContentText: () => '{}' }) },
    DriveApp: ui.drive || {},
    HtmlService: { createHtmlOutput: h => ({ html: h, setWidth() { return this; }, setHeight() { return this; } }) },
    Browser: { msgBox: m => log.push('msg:' + m) },
    hooks: ui.hooks || null
  };
  const stubs = ui.hooks
    ? Object.keys(ui.hooks).map(k => `${k}=hooks.${k};`).join('\n')
    : '';
  const fn = new Function(...Object.keys(env),
    SRC + '\n' + stubs + '\nreturn {registerFromSorting, importSavedList, importWizard, clearWorkSheets, findImageCandidates, saveSelectedImages, extractImageUrls_, resolveUrl_, openPhotoPicker, buildPhotoPickerHtml_};');
  return { api: fn(...Object.values(env)), log };
}

let passed = 0;
function test(name, f) { f(); passed++; console.log('  ok  ' + name); }

// ================= 登録(③)の重複・保留の判定 =================
test('③ 登録: ID一致・同名の距離判定(50m)と保留', () => {
  const main = new Sheet('シート1', [H,
    mk({ type: 'eat_out', title: 'ラーメン太郎', img: 'https://i/1', map_url: fid('aaa'), status: '成功', lat: 35.0, lng: 139.0 }),
    mk({ type: 'eat_out', title: 'ラーメン次郎', img: 'https://i/2', map_url: 'https://example.com/none', status: '成功' }),
    mk({ type: 'eat_out', title: '既存の店', img: 'https://i/3', map_url: fid('bbb'), status: '成功', lat: 35.1, lng: 139.1 })]);
  const ok = (o) => mk(Object.assign({ type: 'eat_out', img: 'https://i/x', status: '成功' }, o));
  const sort = new Sheet('仕分け', [HS,
    ok({ title: 'ユニーク店', map_url: fid('c1'), lat: 36, lng: 140 }),                                  // 登録済
    ok({ title: 'IDは同じで遠い', map_url: fid('bbb'), lat: 36, lng: 140 }),                              // 登録済(IDが同じでも50m以上)
    ok({ title: 'IDは同じで30m', map_url: fid('bbb'), lat: 35.1 + dLat(30), lng: 139.1 }),                // 重複(ID)
    ok({ title: 'IDは同じ座標なし', map_url: fid('bbb') }),                                              // 重複(ID・座標不明)
    ok({ title: 'ラーメン太郎', map_url: fid('c2'), lat: 35.0 + dLat(40), lng: 139.0 }),                  // 重複(同名)
    ok({ title: 'ラーメン太郎', map_url: fid('c3'), lat: 35.0 + dLat(300), lng: 139.0 }),                 // 登録済
    ok({ title: 'ラーメン次郎', map_url: fid('c4'), lat: 35.5, lng: 139.5 }),                             // 保留(既存の座標が不明)
    ok({ type: '', title: 'type未設定', map_url: fid('c5') }),                                           // 保留
    ok({ type: 'go_out', title: '写真なし店', img: '写真なし', map_url: fid('c6'), status: '成功(写真なし)' }), // 保留
    ok({ type: 'go_out', title: 'ラーメン太郎', map_url: fid('c7'), lat: 35.0 + dLat(49), lng: 139.0 }),  // 重複(49m)
    ok({ type: 'go_out', title: 'ラーメン太郎', map_url: fid('c8'), lat: 35.0 + dLat(51), lng: 139.0 })   // 登録済(51m)
  ]);
  const { api } = load({ 'シート1': main, '仕分け': sort });
  api.registerFromSorting();
  // 「登録状態」は ':' より前だけ比べる(後ろは相手の行番号など)
  const reg = sort.g.slice(1).map(r => String(r[14]).split(':')[0]);
  assert.deepStrictEqual(reg, [
    '登録済', '登録済', '重複(ID・30m', '重複(ID・座標不明', '重複(同名・40m', '登録済',
    '保留', '保留', '保留', '重複(同名・49m', '登録済'
  ]);
  assert.strictEqual(main.g.length - 1, 3 + 4); // 既存3 + 登録4
});

// ================= 取り込み(①) =================
test('① 取り込み: 同じURLだけ除外・IDが同じでURLが違えば読み込む', () => {
  const main = new Sheet('シート1', [H, mk({ type: 'eat_out', title: '既存', img: 'https://i', map_url: fid('aaa'), status: '成功' })]);
  const imp = new Sheet('取り込み', [IH, [],
    ['既存と同URL', '', fid('aaa')],
    ['IDだけ同じ', 'メモ1', fid('aaa', '別名')],
    ['新規', 'メモ2', fid('zzz')],
    ['新規と同URL', '', fid('zzz')]]);
  const sheets = { 'シート1': main, '取り込み': imp };
  const { api } = load(sheets);
  api.importSavedList();
  const marks = imp.g.slice(1).map(r => r[5] || '');
  assert.strictEqual(marks[1], '重複(URL:シート1 2行目)');
  assert.strictEqual(marks[2], '取込済');
  assert.strictEqual(marks[3], '取込済');
  assert.strictEqual(marks[4], '重複(URL:取り込み 5行目)');
  const sortTitles = sheets['仕分け'].g.slice(1).map(r => [r[1], r[4]]);
  assert.deepStrictEqual(sortTitles, [['IDだけ同じ', 'メモ1'], ['新規', 'メモ2']]);
  assert.strictEqual(main.g.length, 2); // シート1は変わらない
});

// ================= 取り込みを進める(段階の判定) =================
const wizardCases = [
  ['未読み込みあり', [IH, ['店', '', 'u1']], [HS], 'importSavedList'],
  ['取得待ちあり', [IH, ['店', '', 'u1', '', '', '取込済']], [HS, mk({ title: '店', map_url: 'u' })], 'fetchPlaceData'],
  ['登録待ち(未判定)', [IH, ['店', '', 'u1', '', '', '取込済']], [HS, mk({ title: '店', img: 'https://i', map_url: 'u', status: '成功' })], 'registerFromSorting'],
  ['保留のみ', [IH], [HS, mk({ title: '店', img: 'https://i', map_url: 'u', status: '成功', reg: '保留:typeが未設定' })], 'registerFromSorting'],
  ['すべて解決済み', [IH, ['店', '', 'u1', '', '', '取込済']], [HS, mk({ title: '店', img: 'https://i', map_url: 'u', status: '成功', reg: '登録済' }), mk({ title: '店2', img: 'https://i', map_url: 'u2', status: '成功', reg: '重複(ID・30m:x)' })], 'clearWorkSheets'],
  ['取り込みだけ残っている', [IH, ['店', '', 'u1', '', '', '取込済']], [HS], 'clearWorkSheets'],
  ['未読み込み+仕分け途中は読み込みが先', [IH, ['店', '', 'u1']], [HS, mk({ title: '店', img: 'https://i', map_url: 'u', status: '成功' })], 'importSavedList'],
  ['両方空', [IH], [HS], null],
  ['見出しなし', [['a', 'b'], ['x', 'y']], [HS], null]
];
for (const [name, impRows, sortRows, expected] of wizardCases) {
  test(`▶ 取り込みを進める: ${name}`, () => {
    const called = [];
    const hooks = {};
    for (const k of ['importSavedList', 'fetchPlaceData', 'registerFromSorting', 'clearWorkSheets']) hooks[k] = () => called.push(k);
    const sheets = { '取り込み': new Sheet('取り込み', impRows), '仕分け': new Sheet('仕分け', sortRows), 'シート1': new Sheet('シート1', [H]) };
    const { api } = load(sheets, { hooks });
    api.importWizard();
    assert.deepStrictEqual(called, expected ? [expected] : []);
  });
}

// ================= 後片付け(④) =================
test('④ 後片付け: 取り込みは見出しごと、仕分けは見出し行を残して空にする(未解決は2回確認)', () => {
  const sheets = {
    '取り込み': new Sheet('取り込み', [IH, ['店', '', 'u1', '', '', '取込済']]),
    '仕分け': new Sheet('仕分け', [HS, mk({ title: '店', img: 'https://i', map_url: 'u', status: '成功' })])
  };
  const { api, log } = load(sheets);
  api.clearWorkSheets();
  assert.strictEqual(sheets['取り込み'].g.length, 0);
  assert.strictEqual(sheets['仕分け'].g[0][0], 'type');
  assert.ok(sheets['仕分け'].g[1].every(v => v === ''));
  assert.strictEqual(log.filter(l => l.startsWith('alert:')).length, 2);
});

test('④ 後片付け: 「いいえ」なら何も消さない', () => {
  const sheets = {
    '取り込み': new Sheet('取り込み', [IH, ['店', '', 'u1']]),
    '仕分け': new Sheet('仕分け', [HS, mk({ title: '店', map_url: 'u', reg: '登録済' })])
  };
  const { api } = load(sheets, { answer: 'NO' });
  api.clearWorkSheets();
  assert.strictEqual(sheets['取り込み'].g.length, 2);
  assert.strictEqual(sheets['仕分け'].g[1][1], '店');
});


// ================= 写真を探す =================
const PAGE = `<html><head>
<meta property="og:image" content="/img/top.jpg?a=1&amp;b=2">
<meta name="twitter:image" content="https://cdn.example.com/tw.png">
<link rel="image_src" href="../rel/linked.jpg">
<script type="application/ld+json">{"@type":"Restaurant","image":["https://cdn.example.com/ld1.jpg",{"url":"https://cdn.example.com/ld2.jpg"}]}</script>
</head><body>
<img src="/assets/logo.png" width="300" height="80">
<img src="/small.jpg" width="40" height="40">
<img src="/photos/a.jpg" srcset="/photos/a-480.jpg 480w, /photos/a-960.jpg 960w">
<img data-src="/lazy/b.jpg" src="/placeholder.gif">
<img src="/icons/star.svg">
<img src="/photos/a-960.jpg">
</body></html>`;

test('写真: 相対URLの解決', () => {
  const { api } = load({});
  const b = 'https://shop.example.com/menu/index.html?x=1';
  assert.strictEqual(api.resolveUrl_(b, 'https://a.com/x.jpg'), 'https://a.com/x.jpg');
  assert.strictEqual(api.resolveUrl_(b, '//cdn.com/x.jpg'), 'https://cdn.com/x.jpg');
  assert.strictEqual(api.resolveUrl_(b, '/img/x.jpg'), 'https://shop.example.com/img/x.jpg');
  assert.strictEqual(api.resolveUrl_(b, 'x.jpg'), 'https://shop.example.com/menu/x.jpg');
  assert.strictEqual(api.resolveUrl_(b, '../x.jpg'), 'https://shop.example.com/x.jpg');
  assert.strictEqual(api.resolveUrl_(b, 'data:image/png;base64,AAAA'), '');
});

test('写真: ページから画像の候補を集める(og → JSON-LD → img、ロゴ・小さい画像・重複は除く)', () => {
  const { api } = load({});
  const urls = api.extractImageUrls_(PAGE, 'https://shop.example.com/menu/index.html').map(c => c.url);
  assert.deepStrictEqual(urls, [
    'https://shop.example.com/img/top.jpg?a=1&b=2',   // og:image(&amp; は戻す)
    'https://cdn.example.com/tw.png',                 // twitter:image
    'https://shop.example.com/rel/linked.jpg',        // link rel=image_src
    'https://cdn.example.com/ld1.jpg',                // JSON-LD
    'https://cdn.example.com/ld2.jpg',
    'https://shop.example.com/photos/a-960.jpg',      // srcset は一番大きいもの(重複する単独のimgは除く)
    'https://shop.example.com/lazy/b.jpg'             // data-src を優先
  ]);
});

test('写真: 候補の取得(ドライブのリンク・画像そのもの・HTTPエラー・画像なし)', () => {
  const resp = (code, headers, body) => ({ getResponseCode: () => code, getHeaders: () => headers, getContentText: () => body });
  const { api } = load({}, { fetch: u => {
    if (u.includes('image.jpg')) return resp(200, { 'Content-Type': 'image/jpeg' }, '');
    if (u.includes('forbidden')) return resp(403, {}, '');
    if (u.includes('empty')) return resp(200, { 'Content-Type': 'text/html' }, '<html></html>');
    return resp(200, { 'Content-Type': 'text/html' }, PAGE);
  } });
  assert.deepStrictEqual(api.findImageCandidates('https://drive.google.com/file/d/1AbCdEfGhIjKlMnOp/view'), [{ url: 'https://drive.google.com/thumbnail?sz=w800&id=1AbCdEfGhIjKlMnOp', source: 'drive' }]);
  assert.deepStrictEqual(api.findImageCandidates('https://x.com/image.jpg'), [{ url: 'https://x.com/image.jpg', source: 'direct' }]);
  assert.ok(api.findImageCandidates('https://shop.example.com/').length >= 5);
  assert.throws(() => api.findImageCandidates('https://x.com/forbidden'), /HTTP 403/);
  assert.throws(() => api.findImageCandidates('https://x.com/empty'), /見つかりませんでした/);
  assert.throws(() => api.findImageCandidates('ftp://x'), /http/);
});

test('写真: 選んだ画像をドライブに保存し、画像欄とstatusを更新する(置き換え/追加)', () => {
  const files = [];
  const drive = {
    Access: { ANYONE_WITH_LINK: 'A' }, Permission: { VIEW: 'V' },
    getFoldersByName: () => ({ hasNext: () => false }),
    createFolder: () => ({ createFile: (blob) => { const f = { id: 'FILE' + files.length, name: blob.name, setSharing() {}, getId() { return this.id; } }; files.push(f); return f; } })
  };
  const blobOf = (type, size) => ({ getContentType: () => type, getBytes: () => new Array(size), setName(n) { this.name = n; }, get name() { return this._n; }, set name(v) { this._n = v; } });
  const fetch = (u) => {
    if (u.includes('ok')) return { getResponseCode: () => 200, getBlob: () => blobOf('image/jpeg', 20000) };
    if (u.includes('tiny')) return { getResponseCode: () => 200, getBlob: () => blobOf('image/png', 100) };
    if (u.includes('html')) return { getResponseCode: () => 200, getBlob: () => blobOf('text/html', 20000) };
    return { getResponseCode: () => 404, getBlob: () => blobOf('', 0) };
  };
  const sort = new Sheet('仕分け', [HS, mk({ type: 'eat_out', title: '写真なし店', img: '写真なし', map_url: fid('p1'), status: '失敗:場所を特定できず' })]);
  const { api } = load({ '仕分け': sort }, { fetch, drive });
  const r = api.saveSelectedImages('仕分け', 2, ['https://x.com/ok1.jpg', 'https://x.com/tiny.png', 'https://x.com/html', 'https://x.com/missing', 'ftp://bad'], 'replace');
  assert.strictEqual(r.saved, 1);
  assert.strictEqual(r.failed.length, 3);
  assert.strictEqual(sort.g[1][2], 'https://drive.google.com/thumbnail?sz=w800&id=FILE0');
  assert.strictEqual(sort.g[1][11], '成功(写真を手動で追加)'); // 「失敗」で始まらない → ③で消されない
  // 追加モード
  const r2 = api.saveSelectedImages('仕分け', 2, ['https://x.com/ok2.jpg'], 'append');
  assert.strictEqual(sort.g[1][2].split(',').length, 2);
  assert.strictEqual(r2.saved, 1);
  // 不正な入力
  assert.throws(() => api.saveSelectedImages('別シート', 2, ['https://x.com/ok.jpg'], 'replace'), /シート/);
  assert.throws(() => api.saveSelectedImages('仕分け', 1, ['https://x.com/ok.jpg'], 'replace'), /行/);
  assert.throws(() => api.saveSelectedImages('仕分け', 2, [], 'replace'), /選ばれていません/);
});

test('写真: 選択画面を開く(対象の行の情報が入る・画面のスクリプトが構文として正しい)', () => {
  const main = new Sheet('シート1', [H, mk({ type: 'go_out', title: '</script>危ない"名前', website: 'https://shop.example.com/', img: '写真なし' })]);
  const { api, log } = load({ 'シート1': main }, { active: { sheet: 'シート1', row: 2 } });
  api.openPhotoPicker();
  assert.ok(log.includes('dialog:写真を探す'));
  const html = log.html;
  assert.ok(!/<\/script>危ない/.test(html), 'タイトル内の </script> でHTMLが壊れないこと');
  const script = html.match(/<script>([\s\S]*)<\/script>/)[1];
  new Function('document', 'google', script); // 構文チェック(実行はしない)
  assert.ok(html.includes('"hasPhoto":false'));
  // 対象外のシート・見出し行
  const { api: api2, log: log2 } = load({ '取り込み': new Sheet('取り込み', [IH]) }, { active: { sheet: '取り込み', row: 2 } });
  api2.openPhotoPicker();
  assert.ok(log2.some(l => l.startsWith('alert:')) && !log2.some(l => l.startsWith('dialog:')));
});

console.log(`\n${passed} tests passed`);
