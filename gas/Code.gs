// GAS Version 1.35 (v1.34ベース: 保存リスト取り込み・成功/失敗マーク・6分制限対策・二重実行ロック)
// v1.34: 座標を地図の表示中心ではなく場所のピンから取得
//
// シート「シート1」の列構成:
//   A type / B title / C img / D category / E comment / F website / G address
//   H hide / I map_url / J age_limit / K gtypes(Google元カテゴリ) / L status(処理結果)
// ※ K・L は公開CSVにも出る。秘密や個人情報を書かないこと。
// ※ APIキーは「スクリプトプロパティ」(MAPS_API_KEY / SEARCH_ENGINE_ID)にのみ置く。

const SHEET_MAIN = 'シート1';
const SHEET_IMPORT = '取り込み';
const COL = { TYPE: 1, TITLE: 2, IMG: 3, CATEGORY: 4, COMMENT: 5, WEBSITE: 6, ADDRESS: 7, HIDE: 8, MAP_URL: 9, AGE: 10, GTYPES: 11, STATUS: 12 };
const NUM_COLS = 12;
const TIME_LIMIT_MS = 270000;          // 1回の実行で新しい行に着手してよい時間(上限6分の手前で止める)
const CONTINUE_HANDLER = 'fetchPlaceDataContinue';
const MAX_CONTINUES = 40;              // 続き実行の安全上限(無限ループ防止)

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🌟プレゼン準備')
    .addItem('空欄のデータを自動取得（全部入り）', 'fetchPlaceData')
    .addItem('📥 保存リストを取り込み', 'importSavedList')
    .addItem('🔁 失敗行を再試行', 'retryFailedRows')
    .addItem('🔧 リンク切れ画像を自動修復', 'fixBrokenImages')
    .addToUi();
}

// ==========================================
// 共通ヘルパー
// ==========================================
function getDistance(lat1, lng1, lat2, lng2) {
  const R = 6371000;
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLng = (lng2 - lng1) * Math.PI / 180;
  const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
            Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
            Math.sin(dLng/2) * Math.sin(dLng/2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
  return R * c;
}

// URLから座標を取り出す(場所のピン !3d!4d を優先、無ければ地図の表示中心 @ / ll=)
function extractLatLng(url) {
  const pin = url.match(/!3d(-?[0-9.]+)!4d(-?[0-9.]+)/);
  if (pin) return pin[1] + ',' + pin[2];
  const view = url.match(/@([0-9.,\-]+)/) || url.match(/[?&]ll=([0-9.,\-]+)/);
  return view ? view[1].split(',').slice(0, 2).join(',') : '';
}

// K・L列のヘッダーが無ければ付ける
function ensureHeaders_(sheet) {
  if (!sheet.getRange(1, COL.GTYPES).getValue()) sheet.getRange(1, COL.GTYPES).setValue('gtypes');
  if (!sheet.getRange(1, COL.STATUS).getValue()) sheet.getRange(1, COL.STATUS).setValue('status');
}

// 検索用URL(query=やmaps/search)は地図URLとして使わない
function cleanMapUrl_(raw) {
  const s = raw ? String(raw) : '';
  if (s.includes('query=') || s.includes('maps/search')) return '';
  return s;
}

// 写真アップロード済み(Drive画像)で住所が未取得の施設は、website/addressだけ補完する
function isInfoOnly_(r) {
  const type = r[COL.TYPE - 1];
  return String(r[COL.IMG - 1]).includes('drive.google.com') && !r[COL.ADDRESS - 1] && (type === 'go_out' || type === 'eat_out');
}

// 取得対象の行か(未処理で、名前かURLがあり、画像が空 または 住所だけ補完する行)
function isPendingRow_(r) {
  if (r[COL.STATUS - 1]) return false;
  if (!(r[COL.TITLE - 1] || cleanMapUrl_(r[COL.MAP_URL - 1]))) return false;
  return !r[COL.IMG - 1] || isInfoOnly_(r);
}

// Google側の致命的エラー(キー無効・上限超過)は FATAL: で投げ、全行を失敗扱いにしない
function fetchJson_(url) {
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const json = JSON.parse(res.getContentText());
  if (json.error && (json.error.code === 403 || json.error.code === 429)) {
    throw new Error('FATAL:customsearch ' + json.error.code);
  }
  if (['REQUEST_DENIED', 'OVER_QUERY_LIMIT', 'OVER_DAILY_LIMIT'].indexOf(json.status) >= 0) {
    throw new Error('FATAL:' + json.status);
  }
  return json;
}

function clearContinueTriggers_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === CONTINUE_HANDLER) ScriptApp.deleteTrigger(t);
  });
}

// 未処理の行が残っていれば、1分後に続きを実行するトリガーを1つだけ作る
function scheduleContinueIfPending_(sheet, fatal) {
  const props = PropertiesService.getScriptProperties();
  clearContinueTriggers_();
  if (fatal) { props.deleteProperty('CONT_COUNT'); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const v = sheet.getRange(2, 1, lastRow - 1, NUM_COLS).getValues();
  const pending = v.filter(isPendingRow_).length;
  if (pending === 0) { props.deleteProperty('CONT_COUNT'); return; }

  const cont = Number(props.getProperty('CONT_COUNT') || 0);
  if (cont >= MAX_CONTINUES) {
    console.error('続き実行の上限に達したため停止しました');
    props.deleteProperty('CONT_COUNT');
    return;
  }
  props.setProperty('CONT_COUNT', String(cont + 1));
  ScriptApp.newTrigger(CONTINUE_HANDLER).timeBased().after(60 * 1000).create();
}

function fetchPlaceDataContinue() {
  fetchPlaceData(true);
}

// ==========================================
// 空欄データの自動取得(時間予算つき・ロックつき・成功/失敗マークつき)
// ==========================================
function fetchPlaceData(isBackground = false) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) { console.log('別の取得処理が実行中のためスキップ'); return; }

  let fatal = false;
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MAIN);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    ensureHeaders_(sheet);

    const apiKey = PropertiesService.getScriptProperties().getProperty('MAPS_API_KEY');
    const cx = PropertiesService.getScriptProperties().getProperty('SEARCH_ENGINE_ID');
    if (!apiKey) { console.error('MAPS_API_KEY が未設定です'); return; }

    const data = sheet.getRange(2, 1, lastRow - 1, NUM_COLS).getValues();
    const start = Date.now();

    for (let i = 0; i < data.length; i++) {
      if (!isPendingRow_(data[i])) continue;
      if (Date.now() - start > TIME_LIMIT_MS) break; // 残りは続き実行へ

      const row = i + 2;
      let result;
      try {
        result = processRow_(sheet, row, data[i], apiKey, cx);
      } catch (err) {
        const msg = String(err && err.message ? err.message : err);
        if (msg.indexOf('FATAL:') === 0) {
          fatal = true;
          console.error(msg); // 失敗マークは付けない(全行が失敗扱いになるのを防ぐ)
          break;
        }
        console.error(`行${row}: ${msg}`);
        result = { status: '失敗:例外' }; // エラー文にURL(キー)が混ざる恐れがあるためシートには書かない
      }

      sheet.getRange(row, COL.STATUS).setValue(result.status);
      if (result.gtypes) sheet.getRange(row, COL.GTYPES).setValue(result.gtypes);
      Utilities.sleep(500);
    }

    SpreadsheetApp.flush();
    scheduleContinueIfPending_(sheet, fatal);
  } finally {
    lock.releaseLock();
  }
}

// 1行ぶんの取得処理。r はシートの1行(配列)。{status, gtypes} を返す。致命的エラーは throw
function processRow_(sheet, row, r, apiKey, cx) {
  const type = r[COL.TYPE - 1];
  const title = r[COL.TITLE - 1];
  const mapUrl = cleanMapUrl_(r[COL.MAP_URL - 1]);
  const infoOnly = isInfoOnly_(r);

  // ---- 料理系(eat / lunch): 画像検索のみ ----
  if (type === 'eat' || type === 'lunch') {
    if (!title) return { status: '失敗:なまえが空欄' };
    if (!cx) return { status: '失敗:SEARCH_ENGINE_ID未設定' };

    const searchUrl = `https://customsearch.googleapis.com/customsearch/v1?cx=${cx}&q=${encodeURIComponent(title)}&searchType=image&key=${apiKey}&num=3`;
    const json = fetchJson_(searchUrl);
    if (json.items && json.items.length > 0) {
      sheet.getRange(row, COL.IMG).setValue(json.items.map(item => item.link).join(','));
      sheet.getRange(row, COL.CATEGORY).setValue('ごはん');
      return { status: '成功' };
    }
    return { status: '失敗:画像検索ヒットなし' };
  }

  // ---- それ以外(go_out / eat_out / 未仕分けの空type など): Google Maps で取得 ----
  let latLngStr = '';
  let urlName = '';

  if (mapUrl && mapUrl.includes('goo.gl')) {
    try {
      const options = { "method": "get", "followRedirects": false, "muteHttpExceptions": true };
      const res = UrlFetchApp.fetch(mapUrl, options);
      const locationUrl = res.getHeaders()['Location'];
      if (locationUrl) {
        latLngStr = extractLatLng(locationUrl);

        let match = locationUrl.match(/\/maps\/place\/([^\/?]+)/) || locationUrl.match(/\/maps\/search\/([^\/?]+)/) || locationUrl.match(/[?&]q=([^&]+)/);
        if (match && match[1]) {
          let decoded = decodeURIComponent(match[1].replace(/\+/g, ' '));
          if (!decoded.match(/^[0-9.,\-]+$/) && !decoded.includes('°')) urlName = decoded;
        }
      }
    } catch (e) {}
  }
  // 通常形式のGoogleマップURL(Takeoutの保存リストのURLを含む)からも座標と施設名を取り出す
  if (mapUrl && mapUrl.includes('/maps/place/')) {
    latLngStr = extractLatLng(mapUrl);

    const fullName = mapUrl.match(/\/maps\/place\/([^\/?]+)/);
    if (fullName && fullName[1]) {
      try {
        const decoded = decodeURIComponent(fullName[1].replace(/\+/g, ' '));
        if (!decoded.match(/^[0-9.,\-]+$/) && !decoded.includes('°')) urlName = decoded;
      } catch (e) {}
    }
  }

  let apiSearchQuery = urlName ? urlName : title;
  if (!apiSearchQuery && latLngStr) apiSearchQuery = latLngStr;

  let fallbackMode = false;
  let placeId = null;
  let cleanName = urlName || title;

  if (apiSearchQuery) {
    let searchJson = { results: [] };

    if (latLngStr) {
      const nsUrl = `https://maps.googleapis.com/maps/api/place/nearbysearch/json?location=${latLngStr}&radius=3000&keyword=${encodeURIComponent(apiSearchQuery)}&language=ja&key=${apiKey}`;
      searchJson = fetchJson_(nsUrl);
    }

    if (!searchJson.results || searchJson.results.length === 0) {
      let tsUrl = `https://maps.googleapis.com/maps/api/place/textsearch/json?query=${encodeURIComponent(apiSearchQuery)}&language=ja&key=${apiKey}`;
      if (latLngStr) tsUrl += `&location=${latLngStr}&radius=3000`;
      searchJson = fetchJson_(tsUrl);
    }

    if (searchJson.results && searchJson.results.length > 0) {
      if (latLngStr) {
        let minDistance = Infinity;
        let closestPlaceId = null;
        let closestName = '';

        const coords = latLngStr.split(',');
        const urlLat = parseFloat(coords[0]);
        const urlLng = parseFloat(coords[1]);

        for (let res of searchJson.results) {
          if (res.geometry && res.geometry.location) {
            const dist = getDistance(urlLat, urlLng, res.geometry.location.lat, res.geometry.location.lng);
            if (dist < minDistance) {
              minDistance = dist;
              closestPlaceId = res.place_id;
              closestName = res.name;
            }
          }
        }

        if (minDistance > 10000) fallbackMode = true;
        else if (closestPlaceId) { placeId = closestPlaceId; cleanName = closestName; }
      } else {
        placeId = searchJson.results[0].place_id;
        cleanName = searchJson.results[0].name;
      }
    } else {
      fallbackMode = true;
    }
  } else {
    fallbackMode = true;
  }

  // ---- 場所が特定できなかった場合 ----
  if (fallbackMode || !placeId) {
    let addressStr = '';
    if (latLngStr) {
      try {
        const geoUrl = `https://maps.googleapis.com/maps/api/geocode/json?latlng=${latLngStr}&language=ja&key=${apiKey}`;
        const geoJson = fetchJson_(geoUrl);
        if (geoJson.results && geoJson.results.length > 0) {
          addressStr = geoJson.results[0].formatted_address.replace(/^日本、\s*(〒\d{3}-\d{4}\s*)?/, '');
          const match = addressStr.match(/^(...?[都道府県]).*?[市区町村]/);
          if (match) addressStr = match[0];
        }
      } catch (e) {
        if (String(e.message).indexOf('FATAL:') === 0) throw e;
      }
    }
    if (infoOnly) {
      if (addressStr) {
        sheet.getRange(row, COL.ADDRESS).setValue(addressStr);
        return { status: '成功' };
      }
      return { status: '失敗:住所を取得できず' };
    }

    let categoryStr = '';
    if (type === 'eat_out') categoryStr = 'ごはん';
    else if (type) categoryStr = 'おでかけスポット'; // type未仕分けの行はカテゴリを空にして後で手動仕分け

    if (!title) sheet.getRange(row, COL.TITLE).setValue(cleanName || "名称不明");
    sheet.getRange(row, COL.IMG).setValue('写真なし');
    sheet.getRange(row, COL.CATEGORY).setValue(categoryStr);
    sheet.getRange(row, COL.ADDRESS).setValue(addressStr);
    sheet.getRange(row, COL.MAP_URL).setValue(mapUrl || '');
    return { status: '失敗:場所を特定できず' };
  }

  // ---- 場所が特定できた場合: 詳細取得 ----
  const detailsUrl = `https://maps.googleapis.com/maps/api/place/details/json?place_id=${placeId}&fields=name,photos,types,editorial_summary,website,formatted_address,url&language=ja&key=${apiKey}`;
  const detailsJson = fetchJson_(detailsUrl);

  if (!detailsJson.result) {
    return { status: '失敗:詳細取得失敗(' + (detailsJson.status || '不明') + ')' };
  }

  const result = detailsJson.result;

  let photoUrls = [];
  if (!infoOnly && result.photos && result.photos.length > 0) {
    const maxPhotos = Math.min(result.photos.length, 3);
    for (let p = 0; p < maxPhotos; p++) {
      const photoRef = result.photos[p].photo_reference;
      const photoApiUrl = `https://maps.googleapis.com/maps/api/place/photo?maxwidth=800&photoreference=${photoRef}&key=${apiKey}`;
      const photoRes = UrlFetchApp.fetch(photoApiUrl, { followRedirects: false, muteHttpExceptions: true });
      const finalImageUrl = photoRes.getHeaders()['Location'];
      if (finalImageUrl) photoUrls.push(finalImageUrl);
    }
  }

  let categoryStr = '';
  if (result.types && result.types.length > 0) {
    const typeMap = { 'zoo': 'どうぶつえん', 'aquarium': 'すいぞくかん', 'amusement_park': 'ゆうえんち', 'park': 'こうえん', 'shopping_mall': 'おかいもの', 'restaurant': 'ごはん', 'cafe': 'カフェ', 'museum': 'はくぶつかん', 'tourist_attraction': 'おでかけスポット' };
    for (let t of result.types) { if (typeMap[t]) { categoryStr = typeMap[t]; break; } }
  }

  let addressStr = result.formatted_address || '';
  if (addressStr) {
    addressStr = addressStr.replace(/^日本、\s*(〒\d{3}-\d{4}\s*)?/, '');
    const match = addressStr.match(/^(...?[都道府県]).*?[市区町村]/);
    if (match) addressStr = match[0];
  }

  const gtypes = (result.types || []).join(',');

  if (infoOnly) {
    if (!r[COL.WEBSITE - 1] && result.website) sheet.getRange(row, COL.WEBSITE).setValue(result.website);
    sheet.getRange(row, COL.ADDRESS).setValue(addressStr);
    return { status: '成功', gtypes: gtypes };
  }

  if (!title || String(title).includes('http') || String(title).includes('〒')) sheet.getRange(row, COL.TITLE).setValue(cleanName);
  sheet.getRange(row, COL.IMG).setValue(photoUrls.length > 0 ? photoUrls.join(',') : '写真なし');
  if (categoryStr) sheet.getRange(row, COL.CATEGORY).setValue(categoryStr);
  if (result.website) sheet.getRange(row, COL.WEBSITE).setValue(result.website);
  sheet.getRange(row, COL.ADDRESS).setValue(addressStr);
  sheet.getRange(row, COL.MAP_URL).setValue(mapUrl ? mapUrl : (result.url || ''));

  return {
    status: photoUrls.length > 0 ? '成功' : '成功(写真なし)',
    gtypes: gtypes
  };
}

// ==========================================
// 保存リスト(Takeout CSV)の取り込み
//   シート「取り込み」に CSV(タイトル/メモ/URL/タグ/コメント)を貼って実行。
//   シート1に無いものだけ追加し、取り込み側に「取込済」「重複」を記録する。
// ==========================================
function extractFid_(url) {
  const m = String(url || '').match(/!1s(0x[0-9a-f]+:0x[0-9a-f]+)/i);
  return m ? m[1].toLowerCase() : '';
}

function normalizeTitle_(s) {
  return String(s || '').normalize('NFKC').toLowerCase().replace(/[\s・\-‐–—'’"]/g, '');
}

function importSavedList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(SHEET_IMPORT);
  const main = ss.getSheetByName(SHEET_MAIN);
  if (!src || !main) { Browser.msgBox('シート「取り込み」または「シート1」が見つかりません。'); return; }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) { Browser.msgBox('別の処理が実行中です。少し待ってからやり直してください。'); return; }

  let added = 0, dup = 0;
  try {
    ensureHeaders_(main);
    const srcLast = src.getLastRow();
    if (srcLast < 2) { Browser.msgBox('取り込み用シートにデータがありません。'); return; }

    const lastCol = Math.max(src.getLastColumn(), 1);
    const values = src.getRange(1, 1, srcLast, lastCol).getValues();
    const header = values[0].map(h => String(h).trim());
    const iTitle = header.indexOf('タイトル');
    const iMemo = header.indexOf('メモ');
    const iUrl = header.indexOf('URL');
    const iComment = header.indexOf('コメント');
    if (iTitle < 0 || iUrl < 0) { Browser.msgBox('ヘッダーに「タイトル」「URL」が見つかりません。'); return; }

    let iMark = header.indexOf('取込状態');
    if (iMark < 0) { iMark = lastCol; src.getRange(1, iMark + 1).setValue('取込状態'); }

    // シート1の既存キー(場所の固有ID と 正規化した名称)
    const fids = new Set();
    const titles = new Set();
    const mainLast = main.getLastRow();
    if (mainLast >= 2) {
      const mv = main.getRange(2, 1, mainLast - 1, NUM_COLS).getValues();
      for (const r of mv) {
        const f = extractFid_(r[COL.MAP_URL - 1]);
        if (f) fids.add(f);
        const t = normalizeTitle_(r[COL.TITLE - 1]);
        if (t) titles.add(t);
      }
    }

    const newRows = [];
    const marks = [];
    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const title = String(row[iTitle] || '').trim();
      const url = String(row[iUrl] || '').trim();
      const already = String(row[iMark] === undefined ? '' : row[iMark]).trim();

      if (already || (!title && !url)) { marks.push([already]); continue; } // 処理済み・空行は触らない

      const fid = extractFid_(url);
      const tkey = normalizeTitle_(title);

      if (fid && fids.has(fid)) { marks.push(['重複(ID)']); dup++; continue; }
      if (tkey && titles.has(tkey)) { marks.push(['重複(名称)']); dup++; continue; }

      const memo = iMemo >= 0 ? String(row[iMemo] || '').trim() : '';
      const cmt = iComment >= 0 ? String(row[iComment] || '').trim() : '';
      newRows.push([
        '',           // A type(後で手動仕分け)
        title,        // B title
        '',           // C img
        '',           // D category
        memo || cmt,  // E comment
        '',           // F website
        '',           // G address
        '',           // H hide
        url,          // I map_url
        '',           // J age_limit
        '',           // K gtypes
        ''            // L status
      ]);
      if (fid) fids.add(fid);
      if (tkey) titles.add(tkey);
      marks.push(['取込済']);
      added++;
    }

    if (newRows.length > 0) {
      main.getRange(Math.max(mainLast, 1) + 1, 1, newRows.length, NUM_COLS).setValues(newRows);
    }
    if (marks.length > 0) {
      src.getRange(2, iMark + 1, marks.length, 1).setValues(marks);
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  ss.toast(`${added}件を追加、${dup}件は重複でした。続けて情報を取得します。`, '取り込み', 8);
  if (added > 0) fetchPlaceData(true);
}

// ==========================================
// 失敗行の再試行(statusが「失敗」で始まる行の画像とstatusを空にして再取得)
// ==========================================
function retryFailedRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MAIN);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const v = sheet.getRange(2, 1, lastRow - 1, NUM_COLS).getValues();
  let n = 0;
  for (let i = 0; i < v.length; i++) {
    if (String(v[i][COL.STATUS - 1]).indexOf('失敗') === 0) {
      // 写真アップロード済み(Drive画像)の行は画像を消さない(住所補完のやり直しのため status だけ空にする)
      if (!String(v[i][COL.IMG - 1]).includes('drive.google.com')) sheet.getRange(i + 2, COL.IMG).setValue('');
      sheet.getRange(i + 2, COL.STATUS).setValue('');
      n++;
    }
  }
  SpreadsheetApp.flush();
  if (n === 0) { Browser.msgBox('失敗した行はありませんでした。'); return; }
  SpreadsheetApp.getActiveSpreadsheet().toast(`${n}件を再試行します。`, '再試行', 5);
  fetchPlaceData(true);
}

// ==========================================
// フォーム送信時の処理
// ==========================================
function onFormSubmit(e) {
  try {
    const mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MAIN);
    const responses = e.namedValues;
    if (!responses) return;

    const choice = responses['種類'] ? responses['種類'][0] : '';
    const title = responses['なまえ'] ? responses['なまえ'][0] : '';
    const mapUrl = responses['マップURL'] ? responses['マップURL'][0] : '';
    const photoRawUrl = responses['写真'] ? responses['写真'][0] : '';
    // フォームから送信された「対象年齢」を取得する
    const ageLimit = responses['対象年齢'] ? responses['対象年齢'][0] : '';
    const comment = responses['コメント'] ? responses['コメント'][0] : '';

    let type = 'eat';
    if (choice.includes('昼食')) type = 'lunch';
    else if (choice.includes('お出かけ')) type = 'go_out';
    else if (choice.includes('外食')) type = 'eat_out';

    let finalPhotoUrl = '';
    let category = '';
    if (type === 'eat' || type === 'eat_out' || type === 'lunch') category = 'ごはん';

    if (photoRawUrl) {
      let fileId = '';
      const idMatch = photoRawUrl.match(/id=([^&]+)/);
      const dMatch = photoRawUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
      if (idMatch) fileId = idMatch[1];
      else if (dMatch) fileId = dMatch[1];
      if (fileId) {
        try {
          const file = DriveApp.getFileById(fileId);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          finalPhotoUrl = 'https://drive.google.com/thumbnail?sz=w800&id=' + fileId;
        } catch (error) { finalPhotoUrl = 'エラー: 画像の権限変更に失敗しました'; }
      }
    }

    // 列構成 A〜L に合わせる（H=hideは空、I=map_url）
    mainSheet.appendRow([
      type,           // A type
      title,          // B title
      finalPhotoUrl,  // C img
      category,       // D category
      comment,        // E comment
      '',             // F website
      '',             // G address
      '',             // H hide
      mapUrl,         // I map_url
      ageLimit,       // J age_limit
      '',             // K gtypes
      ''              // L status
    ]);
    SpreadsheetApp.flush();

    // 写真ありの施設も、住所・公式サイトを補完するため取得を走らせる
    if (title || mapUrl) {
      try { fetchPlaceData(true); } catch (fetchError) { console.error(String(fetchError)); }
    }
  } catch (e) {
    console.error(String(e));
  }
}

// ==========================================
// リンク切れ画像を自動検知して修復するプログラム
// ==========================================
function fixBrokenImages() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MAIN);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // A列からC列（写真URL）までを取得
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 3);
  const data = dataRange.getValues();

  let brokenCount = 0;

  // 1. リンク切れの画像をチェックして空欄にする(再取得できるよう status も空にする)
  for (let i = 0; i < data.length; i++) {
    const imgStr = data[i][2]; // C列（写真）

    // Googleの画像URLが入っている場合のみチェック
    if (imgStr && imgStr.includes('googleusercontent.com')) {
      const firstUrl = imgStr.split(',')[0].trim(); // 複数ある場合は最初の1枚目でテスト

      try {
        const res = UrlFetchApp.fetch(firstUrl, { muteHttpExceptions: true });

        // 400エラー等（リンク切れ）が返ってきた場合
        if (res.getResponseCode() >= 400) {
          console.log(`リンク切れ検知 (${i + 2}行目): ${firstUrl}`);
          sheet.getRange(i + 2, COL.IMG).setValue('');
          sheet.getRange(i + 2, COL.STATUS).setValue('');
          brokenCount++;
        }
      } catch (e) {
        // 接続エラー等の場合も空欄にする
        console.log(`アクセスエラー (${i + 2}行目)`);
        sheet.getRange(i + 2, COL.IMG).setValue('');
        sheet.getRange(i + 2, COL.STATUS).setValue('');
        brokenCount++;
      }
    }
  }

  // 2. 空欄にしたデータを再取得する
  if (brokenCount > 0) {
    SpreadsheetApp.flush();
    fetchPlaceData(true);
    Browser.msgBox(`完了！\n${brokenCount}件のリンク切れ画像を検知し、最新の画像に更新しました。\n(件数が多い場合は、残りが自動で続けて処理されます)`);
  } else {
    Browser.msgBox('リンク切れの画像はありませんでした。すべて正常です！');
  }
}

// ==========================================
// デバッグ用(手動実行・ログのみ。シートは変更しない)
// ==========================================
function debugYasuragi() {
  const key = PropertiesService.getScriptProperties().getProperty('MAPS_API_KEY');
  const loc = '35.4652664,139.4934985';
  const q = encodeURIComponent('やすらぎのみち');
  const base = 'https://maps.googleapis.com/maps/api';
  const urls = {
    nearby: `${base}/place/nearbysearch/json?location=${loc}&radius=3000&keyword=${q}&language=ja&key=${key}`,
    text: `${base}/place/textsearch/json?query=${q}&location=${loc}&radius=3000&language=ja&key=${key}`,
    geocode: `${base}/geocode/json?latlng=${loc}&language=ja&key=${key}`
  };
  for (const name in urls) {
    const json = JSON.parse(UrlFetchApp.fetch(urls[name], {muteHttpExceptions: true}).getContentText());
    const top = (json.results || []).slice(0, 3).map(r => r.name || r.formatted_address);
    console.log(name, json.status, json.error_message || '', JSON.stringify(top));
  }
}
