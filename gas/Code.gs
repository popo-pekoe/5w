// GAS Version 1.44 (v1.44: 写真保存(v1.41〜v1.43)とGeocoding修正を統合 / v1.41: Geocodingの拒否で取得全体を止めない(リモート側) / 写真v1.43: 新規取得の写真をその場でドライブに保存(URL失効対策・枚数は PLACE_PHOTO_COUNT) / 写真v1.42: 写真移行でドライブ容量不足・1日の上限を見分けて即停止 / 写真v1.41: Google写真URLをドライブへ移行 / v1.36: 仕分けシート・2段階の取り込み・座標(lat/lng)による重複判定 / v1.37: 距離50m(ID一致も距離で判定)・作業シートを空にするメニュー / v1.38: 取り込みを進める(途中なら続きから) / v1.39: 写真を探す(URLから画像の候補を選んで保存) / v1.40: メモとコメントを両方残す・コメント補完)
// v1.34: 座標を地図の表示中心ではなく場所のピンから取得
//
// 取り込みの流れ:
//   ① 「取り込み」シートにTakeoutのCSVを貼る → メニュー①で「仕分け」へ読み込み＋不足データを自動取得
//   ② 「仕分け」シートで人力の整理(type仕分け・データ修正・不要行の削除)
//   ③ メニュー③で「シート1」へ登録(失敗行の再取得・重複/保留の判定つき)
//   ④ 全部解決したら「作業シートを空にする」
//   ※ 作業シートにデータが残っている間は作業の途中とみなす。メニュー「▶ 取り込みを進める」は、
//      今どの段階かを判定して、その続きから実行する(①〜④を個別に実行することもできる)。
//
// シート「シート1」「仕分け」の列構成:
//   A type / B title / C img / D category / E comment / F website / G address
//   H hide / I map_url / J age_limit / K gtypes(Google元カテゴリ) / L status(処理結果)
//   M lat / N lng (場所の座標。重複判定用)   ※「仕分け」のみ O 登録状態
// ※ K〜N は公開CSVにも出る。秘密や個人情報を書かないこと。
// ※ APIキーは「スクリプトプロパティ」(MAPS_API_KEY / SEARCH_ENGINE_ID)にのみ置く。

const SHEET_MAIN = 'シート1';
const SHEET_IMPORT = '取り込み';
const SHEET_SORT = '仕分け';
const COL = { TYPE: 1, TITLE: 2, IMG: 3, CATEGORY: 4, COMMENT: 5, WEBSITE: 6, ADDRESS: 7, HIDE: 8, MAP_URL: 9, AGE: 10, GTYPES: 11, STATUS: 12, LAT: 13, LNG: 14, REG: 15 };
const NUM_COLS = 14;                   // シート1・仕分け共通のデータ列数(A〜N)
const SORT_COLS = 15;                  // 仕分けは O 列(登録状態)まで
const MAIN_HEADERS = ['type', 'title', 'img', 'category', 'comment', 'website', 'address', 'hide', 'map_url', 'age_limit', 'gtypes', 'status', 'lat', 'lng'];
const TYPE_LIST = ['go_out', 'eat_out', 'eat', 'lunch'];
const TIME_LIMIT_MS = 270000;          // 1回の実行で新しい行に着手してよい時間(上限6分の手前で止める)
const CONTINUE_HANDLER = 'fetchPlaceDataContinue';
const MAX_CONTINUES = 40;              // 続き実行の安全上限(無限ループ防止)
const DUP_DISTANCE_M = 50;             // 場所IDまたは店名が同じ店が、この距離(m)未満なら重複、以上なら別の店とみなす
const PLACE_PHOTO_COUNT = 3;           // 場所1件あたりに取得する写真の最大枚数(Places Photo は1枚ごとに課金。減らせば料金も減る)
const REDIRECT_FETCH_LIMIT = 30;       // 重複判定で短縮URLの座標を調べる最大回数(1回の実行あたり)

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🌟プレゼン準備')
    .addItem('▶ 取り込みを進める(途中なら続きから)', 'importWizard')
    .addSeparator()
    .addItem('① 保存リストを「仕分け」へ読み込み', 'importSavedList')
    .addItem('③ 「仕分け」をシート1へ登録', 'registerFromSorting')
    .addSeparator()
    .addItem('空欄のデータを自動取得（全部入り）', 'fetchPlaceData')
    .addItem('🔁 失敗行を再試行', 'retryFailedRows')
    .addItem('🔧 リンク切れ画像を自動修復', 'fixBrokenImages')
    .addItem('🖼 写真を探す(選択中の行・URLから候補を選ぶ)', 'openPhotoPicker')
    .addItem('📦 Googleの写真URLをドライブへ移行', 'migratePhotosToDrive')
    .addItem('📝 コメント補完(取り込み元のメモ+コメント)', 'restoreCommentsFromImport')
    .addSeparator()
    .addItem('🧹 作業シートを空にする(取り込み・仕分け)', 'clearWorkSheets')
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

// K〜N列(と仕分けのO列)のヘッダーが無ければ付け、列が足りなければ増やす
function ensureHeaders_(sheet, cols) {
  const need = cols || NUM_COLS;
  if (sheet.getMaxColumns() < need) sheet.insertColumnsAfter(sheet.getMaxColumns(), need - sheet.getMaxColumns());
  for (let c = COL.GTYPES; c <= COL.LNG; c++) {
    if (!sheet.getRange(1, c).getValue()) sheet.getRange(1, c).setValue(MAIN_HEADERS[c - 1]);
  }
  if (need >= SORT_COLS && !sheet.getRange(1, COL.REG).getValue()) sheet.getRange(1, COL.REG).setValue('登録状態');
}

// シート「仕分け」が無ければ作る(シート1と同じ列構成+登録状態、typeはプルダウン)
function ensureSortSheet_(ss) {
  let sh = ss.getSheetByName(SHEET_SORT);
  if (!sh) {
    sh = ss.insertSheet(SHEET_SORT);
    sh.getRange(1, 1, 1, SORT_COLS).setValues([MAIN_HEADERS.concat(['登録状態'])]);
    sh.setFrozenRows(1);
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(TYPE_LIST, true).build();
    sh.getRange(2, COL.TYPE, 1000, 1).setDataValidation(rule);
  }
  ensureHeaders_(sh, SORT_COLS);
  return sh;
}

function writeLatLng_(sheet, row, lat, lng) {
  const a = parseFloat(lat), b = parseFloat(lng);
  if (!isFinite(a) || !isFinite(b)) return;
  sheet.getRange(row, COL.LAT).setValue(a);
  sheet.getRange(row, COL.LNG).setValue(b);
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

// 未処理の行の数(シート1・仕分けの合計)
function pendingCount_(names) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let n = 0;
  for (const name of names) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) continue;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;
    n += sheet.getRange(2, 1, lastRow - 1, NUM_COLS).getValues().filter(isPendingRow_).length;
  }
  return n;
}

// 未処理の行が残っていれば、1分後に続きを実行するトリガーを1つだけ作る
function scheduleContinueIfPending_(fatal) {
  const props = PropertiesService.getScriptProperties();
  clearContinueTriggers_();
  if (fatal) { props.deleteProperty('CONT_COUNT'); return; }

  const pending = pendingCount_([SHEET_MAIN, SHEET_SORT]);
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
//   対象: シート1 と 仕分け。ロックは呼び出し側で取る。
// ==========================================
function enrichSheets_(names, start, apiKey, cx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let fatal = false, timedOut = false;

  for (const name of names) {
    if (fatal || timedOut) break;
    const sheet = ss.getSheetByName(name);
    if (!sheet) continue;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;
    ensureHeaders_(sheet, name === SHEET_SORT ? SORT_COLS : NUM_COLS);

    const data = sheet.getRange(2, 1, lastRow - 1, NUM_COLS).getValues();
    for (let i = 0; i < data.length; i++) {
      if (!isPendingRow_(data[i])) continue;
      if (Date.now() - start > TIME_LIMIT_MS) { timedOut = true; break; } // 残りは続き実行へ

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
        console.error(`${name} 行${row}: ${msg}`);
        result = { status: '失敗:例外' }; // エラー文にURL(キー)が混ざる恐れがあるためシートには書かない
      }

      sheet.getRange(row, COL.STATUS).setValue(result.status);
      if (result.gtypes) sheet.getRange(row, COL.GTYPES).setValue(result.gtypes);
      Utilities.sleep(500);
    }
  }
  SpreadsheetApp.flush();
  return { fatal: fatal, timedOut: timedOut };
}

function fetchPlaceData(isBackground = false) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) { console.log('別の取得処理が実行中のためスキップ'); return; }
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('MAPS_API_KEY');
    const cx = PropertiesService.getScriptProperties().getProperty('SEARCH_ENGINE_ID');
    if (!apiKey) { console.error('MAPS_API_KEY が未設定です'); return; }

    const r = enrichSheets_([SHEET_MAIN, SHEET_SORT], Date.now(), apiKey, cx);
    scheduleContinueIfPending_(r.fatal);
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
        // Geocodingが未許可でも住所なしで続ける(他の行の処理を止めない)
        const msg = String(e.message);
        if (msg.indexOf('FATAL:') === 0 && msg !== 'FATAL:REQUEST_DENIED') throw e;
        console.error('Geocoding: ' + msg);
      }
    }
    if (latLngStr) {
      const ll = latLngStr.split(',');
      writeLatLng_(sheet, row, ll[0], ll[1]);
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
  const detailsUrl = `https://maps.googleapis.com/maps/api/place/details/json?place_id=${placeId}&fields=name,photos,types,editorial_summary,website,formatted_address,url,geometry&language=ja&key=${apiKey}`;
  const detailsJson = fetchJson_(detailsUrl);

  if (!detailsJson.result) {
    return { status: '失敗:詳細取得失敗(' + (detailsJson.status || '不明') + ')' };
  }

  const result = detailsJson.result;

  let photoUrls = [];
  if (!infoOnly && result.photos && result.photos.length > 0) {
    const maxPhotos = Math.min(result.photos.length, PLACE_PHOTO_COUNT);
    for (let p = 0; p < maxPhotos; p++) {
      const photoRef = result.photos[p].photo_reference;
      const photoApiUrl = `https://maps.googleapis.com/maps/api/place/photo?maxwidth=800&photoreference=${photoRef}&key=${apiKey}`;
      const photoRes = UrlFetchApp.fetch(photoApiUrl, { followRedirects: false, muteHttpExceptions: true });
      const finalImageUrl = photoRes.getHeaders()['Location'];
      if (!finalImageUrl) continue;
      // Googleの写真URLは一定期間で変わる/切れるため、その場でドライブに保存してそのURLを使う(Photo APIの追加呼び出しはなし)。
      // 保存できなかったとき(容量不足など)は、従来どおりGoogleのURLを入れておく(後で「写真をドライブへ移行」で保存できる)
      photoUrls.push(savePhotoUrlToDrive_(finalImageUrl, (cleanName || title || '写真') + '_' + (p + 1)) || finalImageUrl);
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
  if (result.geometry && result.geometry.location) {
    writeLatLng_(sheet, row, result.geometry.location.lat, result.geometry.location.lng);
  }

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
//   同じURLが無いものだけ「仕分け」へ追加し、取り込み側に「取込済」「重複」を記録する。
//   メモ→comment列(コメントもあれば「メモ / コメント」)。タグは使わない。
// ==========================================
function extractFid_(url) {
  const m = String(url || '').match(/!1s(0x[0-9a-f]+:0x[0-9a-f]+)/i);
  return m ? m[1].toLowerCase() : '';
}

// Takeoutの「メモ」と「コメント」を1つにまとめる(両方あれば「メモ / コメント」。同じ内容なら1つ)
function joinComment_(memo, cmt) {
  memo = String(memo || '').trim();
  cmt = String(cmt || '').trim();
  if (memo && cmt && memo !== cmt) return memo + ' / ' + cmt;
  return memo || cmt;
}

function normalizeTitle_(s) {
  return String(s || '').normalize('NFKC').toLowerCase().replace(/[\s・\-‐–—'’"]/g, '');
}

// ① 取り込みシートのCSVを「仕分け」へ読み込む(同じURLが既にあるものは除く)。続けて不足データを取得する
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
    const sort = ensureSortSheet_(ss);
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

    // 既にシート1・仕分けにある同じURL(貼り直し・二重取り込みの防止)。
    // 場所IDが同じでもURLが違えば読み込み、重複かどうかは③の登録時に座標の距離で判定する
    const urls = new Map(); // URL → どこにあるか(印に書いて、重複の相手が分かるようにする)
    for (const sh of [main, sort]) {
      const last = sh.getLastRow();
      if (last < 2) continue;
      const vals = sh.getRange(2, 1, last - 1, NUM_COLS).getValues();
      for (let k = 0; k < vals.length; k++) {
        const u = String(vals[k][COL.MAP_URL - 1] || '').trim();
        if (u && !urls.has(u)) urls.set(u, `${sh.getName()} ${k + 2}行目`);
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

      if (url && urls.has(url)) { marks.push([`重複(URL:${urls.get(url)})`]); dup++; continue; }

      const memo = iMemo >= 0 ? String(row[iMemo] || '').trim() : '';
      const cmt = iComment >= 0 ? String(row[iComment] || '').trim() : '';
      const out = new Array(NUM_COLS).fill('');
      out[COL.TITLE - 1] = title;
      out[COL.COMMENT - 1] = joinComment_(memo, cmt);
      out[COL.MAP_URL - 1] = url;
      newRows.push(out);
      if (url) urls.set(url, `取り込み ${r + 1}行目`);
      marks.push(['取込済']);
      added++;
    }

    if (newRows.length > 0) {
      sort.getRange(Math.max(sort.getLastRow(), 1) + 1, 1, newRows.length, NUM_COLS).setValues(newRows);
    }
    if (marks.length > 0) {
      src.getRange(2, iMark + 1, marks.length, 1).setValues(marks);
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  ss.toast(`${added}件を「仕分け」に追加、${dup}件は重複でした。続けて情報を取得します。`, '取り込み', 8);
  if (added > 0) fetchPlaceData(true);
}

// ==========================================
// 📝 コメント補完: 古い版の取り込みで捨てられた「コメント」を、「取り込み」シートの元データから補う
//   「取り込み」に残っているCSV(または貼り直したCSV)を使い、同じURLの行(仕分け・シート1)の comment が
//   「メモだけ」のままのものを「メモ / コメント」に直す。手で書き換えた行は触らない。
// ==========================================
function restoreCommentsFromImport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const imp = ss.getSheetByName(SHEET_IMPORT);
  if (!imp || imp.getLastRow() < 2) { ui.alert('「取り込み」シートに元のCSVがありません。TakeoutのCSVを貼ってから実行してください。'); return; }

  const values = imp.getRange(1, 1, imp.getLastRow(), Math.max(imp.getLastColumn(), 1)).getValues();
  const header = values[0].map(h => String(h).trim());
  const iMemo = header.indexOf('メモ'), iUrl = header.indexOf('URL'), iComment = header.indexOf('コメント');
  if (iUrl < 0 || iMemo < 0 || iComment < 0) { ui.alert('「取り込み」の見出しに「メモ」「URL」「コメント」が見つかりません。'); return; }

  const byUrl = new Map();
  for (let r = 1; r < values.length; r++) {
    const url = String(values[r][iUrl] || '').trim();
    const memo = String(values[r][iMemo] || '').trim();
    const cmt = String(values[r][iComment] || '').trim();
    if (url && memo && cmt && memo !== cmt) byUrl.set(url, { memo: memo, joined: joinComment_(memo, cmt) });
  }
  if (byUrl.size === 0) { ui.alert('メモとコメントの両方が入っている行がありませんでした。補完するものはありません。'); return; }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) { ui.alert('別の処理が実行中です。少し待ってからやり直してください。'); return; }
  let fixed = 0;
  try {
    for (const name of [SHEET_SORT, SHEET_MAIN]) {
      const sheet = ss.getSheetByName(name);
      if (!sheet || sheet.getLastRow() < 2) continue;
      const v = sheet.getRange(2, 1, sheet.getLastRow() - 1, NUM_COLS).getValues();
      for (let i = 0; i < v.length; i++) {
        const info = byUrl.get(String(v[i][COL.MAP_URL - 1] || '').trim());
        if (!info) continue;
        if (String(v[i][COL.COMMENT - 1] || '').trim() !== info.memo) continue; // 手で直した行は触らない
        sheet.getRange(i + 2, COL.COMMENT).setValue(info.joined);
        fixed++;
      }
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  ss.toast(`${fixed}件のコメントを補いました。`, 'コメント補完', 6);
}

// ==========================================
// 失敗行の再試行(statusが「失敗」で始まる行の画像とstatusを空にして再取得)
//   対象: シート1 と 仕分け(登録済・重複の行は除く)
// ==========================================
function retryFailedRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let n = 0;
  for (const name of [SHEET_MAIN, SHEET_SORT]) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) continue;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;
    const v = sheet.getRange(2, 1, lastRow - 1, name === SHEET_SORT ? SORT_COLS : NUM_COLS).getValues();
    for (let i = 0; i < v.length; i++) {
      const reg = name === SHEET_SORT ? String(v[i][COL.REG - 1]) : '';
      if (reg === '登録済' || reg.indexOf('重複') === 0) continue;
      if (String(v[i][COL.STATUS - 1]).indexOf('失敗') === 0) {
        // 写真アップロード済み(Drive画像)の行は画像を消さない(住所補完のやり直しのため status だけ空にする)
        if (!String(v[i][COL.IMG - 1]).includes('drive.google.com')) sheet.getRange(i + 2, COL.IMG).setValue('');
        sheet.getRange(i + 2, COL.STATUS).setValue('');
        n++;
      }
    }
  }
  SpreadsheetApp.flush();
  if (n === 0) { Browser.msgBox('失敗した行はありませんでした。'); return; }
  ss.toast(`${n}件を再試行します。`, '再試行', 5);
  fetchPlaceData(true);
}

// ==========================================
// ③ 「仕分け」→「シート1」への登録
//   1. 失敗行を再取得  2. 重複・保留を判定  3. 登録できる行だけをシート1へ追加
//   「登録状態」列: 登録済 / 重複(…) / 保留:理由   (保留は直して再実行すれば再判定される)
// ==========================================
function noPhoto_(img) {
  const s = String(img || '').trim();
  return !s || s === '写真なし' || s.indexOf('エラー') === 0;
}

// 行の座標を返す(lat/lng列 → URLのピン → 短縮URLのリダイレクト先)。分からなければ null
function makeCoordResolver_() {
  const cache = new Map();
  let budget = REDIRECT_FETCH_LIMIT;
  return function (r) {
    const lat = r[COL.LAT - 1], lng = r[COL.LNG - 1];
    if (lat !== '' && lng !== '' && isFinite(Number(lat)) && isFinite(Number(lng))) {
      return { lat: Number(lat), lng: Number(lng) };
    }
    const url = String(r[COL.MAP_URL - 1] || '');
    if (!url) return null;
    if (cache.has(url)) return cache.get(url);

    let ll = extractLatLng(url);
    let skipped = false;
    if (!ll && url.includes('goo.gl')) {
      if (budget > 0) {
        budget--;
        try {
          const res = UrlFetchApp.fetch(url, { method: 'get', followRedirects: false, muteHttpExceptions: true });
          const loc = res.getHeaders()['Location'];
          if (loc) ll = extractLatLng(loc);
        } catch (e) {}
      } else {
        skipped = true; // 回数上限で調べられなかった(次回に持ち越し)
      }
    }
    let out = null;
    if (ll) {
      const parts = ll.split(',');
      const a = parseFloat(parts[0]), b = parseFloat(parts[1]);
      if (isFinite(a) && isFinite(b)) out = { lat: a, lng: b };
    }
    if (!skipped) cache.set(url, out);
    return out;
  };
}

function registerFromSorting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sort = ss.getSheetByName(SHEET_SORT);
  const main = ss.getSheetByName(SHEET_MAIN);
  if (!sort || !main) { Browser.msgBox('シート「仕分け」または「シート1」が見つかりません。'); return; }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) { Browser.msgBox('別の処理が実行中です。少し待ってからやり直してください。'); return; }

  let nReg = 0, nDup = 0, nHold = 0;
  let message = '';
  try {
    ensureHeaders_(main);
    ensureSortSheet_(ss);
    const apiKey = PropertiesService.getScriptProperties().getProperty('MAPS_API_KEY');
    const cx = PropertiesService.getScriptProperties().getProperty('SEARCH_ENGINE_ID');
    if (!apiKey) { Browser.msgBox('MAPS_API_KEY が未設定です。'); return; }

    const sLast0 = sort.getLastRow();
    if (sLast0 < 2) { Browser.msgBox('「仕分け」にデータがありません。'); return; }

    // 1. 失敗行(写真が取れなかった行など)を、直した内容でもう一度取得する
    const v0 = sort.getRange(2, 1, sLast0 - 1, SORT_COLS).getValues();
    for (let i = 0; i < v0.length; i++) {
      const reg = String(v0[i][COL.REG - 1]);
      if (reg === '登録済' || reg.indexOf('重複') === 0) continue;
      if (String(v0[i][COL.STATUS - 1]).indexOf('失敗') === 0) {
        if (!String(v0[i][COL.IMG - 1]).includes('drive.google.com')) sort.getRange(i + 2, COL.IMG).setValue('');
        sort.getRange(i + 2, COL.STATUS).setValue('');
      }
    }
    SpreadsheetApp.flush();
    const e = enrichSheets_([SHEET_SORT], Date.now(), apiKey, cx);
    if (e.fatal || pendingCount_([SHEET_SORT]) > 0) {
      scheduleContinueIfPending_(e.fatal);
      message = e.fatal
        ? 'APIのエラー(キー無効・上限超過など)で取得を止めました。実行ログを確認してください。'
        : '取得が続いています(自動で続きを実行します)。終わってから、もう一度③を実行してください。';
      ss.toast(message, '登録', 10);
      return;
    }

    // 2. 重複・保留の判定
    const sLast = sort.getLastRow();
    const sv = sort.getRange(2, 1, sLast - 1, SORT_COLS).getValues();
    const mLast = main.getLastRow();
    const mv = mLast >= 2 ? main.getRange(2, 1, mLast - 1, NUM_COLS).getValues() : [];

    const coordOf = makeCoordResolver_();
    const fidMap = new Map();   // 場所ID → [{r: 行, label: どこにあるか}]
    const byTitle = new Map();  // 正規化した名称 → [{r: 行, label: どこにあるか}]
    const addPool = (r, label) => {
      const entry = { r: r, label: label };
      const f = extractFid_(r[COL.MAP_URL - 1]);
      if (f) {
        if (!fidMap.has(f)) fidMap.set(f, []);
        fidMap.get(f).push(entry);
      }
      const k = normalizeTitle_(r[COL.TITLE - 1]);
      if (k) {
        if (!byTitle.has(k)) byTitle.set(k, []);
        byTitle.get(k).push(entry);
      }
    };
    mv.forEach((r, k) => addPool(r, `${SHEET_MAIN} ${k + 2}行目`));

    const newRows = [];
    const regs = [];
    for (let i = 0; i < sv.length; i++) {
      const r = sv[i];
      const reg0 = String(r[COL.REG - 1] || '');
      if (reg0 === '登録済' || reg0.indexOf('重複') === 0) { regs.push([reg0]); continue; }

      const type = String(r[COL.TYPE - 1] || '').trim();
      const title = String(r[COL.TITLE - 1] || '').trim();
      const url = String(r[COL.MAP_URL - 1] || '').trim();
      if (!title && !url) { regs.push([reg0]); continue; } // 空行

      let reg = '登録済';
      if (!type) reg = '保留:typeが未設定';
      else if (!title) reg = '保留:なまえが空欄';
      else if (String(r[COL.STATUS - 1]).indexOf('失敗') === 0) reg = '保留:' + r[COL.STATUS - 1];
      else if (noPhoto_(r[COL.IMG - 1])) reg = '保留:写真が未取得';
      else {
        // 同じ場所ID または 同じ店名 の相手を、座標の距離で判定する
        const fid = extractFid_(url);
        const idCands = fid ? (fidMap.get(fid) || []) : [];
        const nameCands = byTitle.get(normalizeTitle_(title)) || [];
        if (idCands.length > 0 || nameCands.length > 0) {
          const me = coordOf(r);
          let dupLabel = '';
          let unknown = false;

          // 場所IDが同じ: 距離が近い、または座標が不明なら重複(IDは同じ場所の強い証拠)。DUP_DISTANCE_M以上離れていれば別の店
          for (const c of idCands) {
            const cc = me ? coordOf(c.r) : null;
            if (!me || !cc) { dupLabel = `重複(ID・座標不明:${c.label})`; break; }
            const d = getDistance(me.lat, me.lng, cc.lat, cc.lng);
            if (d < DUP_DISTANCE_M) { dupLabel = `重複(ID・${Math.round(d)}m:${c.label})`; break; }
          }
          // 店名だけ同じ: 距離が近ければ重複、離れていれば別の店、座標が不明なら保留
          if (!dupLabel) {
            for (const c of nameCands) {
              if (idCands.indexOf(c) >= 0) continue; // IDでも判定済み(離れていると判定された相手)
              const cc = me ? coordOf(c.r) : null;
              if (!me || !cc) { unknown = true; continue; }
              const d = getDistance(me.lat, me.lng, cc.lat, cc.lng);
              if (d < DUP_DISTANCE_M) { dupLabel = `重複(同名・${Math.round(d)}m:${c.label})`; break; }
            }
          }
          if (dupLabel) reg = dupLabel;
          else if (unknown) reg = '保留:同名あり・座標不明';
        }
      }

      if (reg === '登録済') { newRows.push(r.slice(0, NUM_COLS)); addPool(r, `${SHEET_SORT} ${i + 2}行目`); nReg++; }
      else if (reg.indexOf('重複') === 0) nDup++;
      else nHold++;
      regs.push([reg]);
    }

    // 3. 書き込み(シート1 → 仕分けの順)
    if (newRows.length > 0) {
      main.getRange(Math.max(mLast, 1) + 1, 1, newRows.length, NUM_COLS).setValues(newRows);
    }
    sort.getRange(2, COL.REG, regs.length, 1).setValues(regs);
    SpreadsheetApp.flush();
    message = `登録 ${nReg}件 / 重複 ${nDup}件 / 保留 ${nHold}件(「仕分け」の登録状態列を確認)`;
  } finally {
    lock.releaseLock();
  }
  if (message) ss.toast(message, '登録', 10);
}

// ==========================================
// ▶ 取り込みを進める(作業シートの状態から「今どの段階か」を判定して、その続きを実行する)
//   取り込みに未読み込みの行      → ① 読み込み
//   仕分けに取得待ちの行          → ① の続き(情報取得)
//   仕分けに登録待ち・保留の行    → ② 人力の整理のあと ③ 登録
//   すべて解決済み                → ④ 後片付け
//   どちらも空                    → 新規開始の案内
// ==========================================
function summarizeImport_(imp) {
  const s = { total: 0, pending: 0, badHeader: false };
  if (!imp || imp.getLastRow() < 2) return s;
  const lastCol = Math.max(imp.getLastColumn(), 1);
  const values = imp.getRange(1, 1, imp.getLastRow(), lastCol).getValues();
  const header = values[0].map(h => String(h).trim());
  const iTitle = header.indexOf('タイトル');
  const iUrl = header.indexOf('URL');
  const iMark = header.indexOf('取込状態');
  if (iTitle < 0 || iUrl < 0) { s.badHeader = true; return s; }
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (!String(row[iTitle] || '').trim() && !String(row[iUrl] || '').trim()) continue;
    s.total++;
    const mark = iMark >= 0 ? String(row[iMark] === undefined ? '' : row[iMark]).trim() : '';
    if (!mark) s.pending++;
  }
  return s;
}

function summarizeSort_(sort) {
  const s = { total: 0, pendingFetch: 0, undecided: 0, hold: 0, registered: 0, dup: 0 };
  if (!sort || sort.getLastRow() < 2) return s;
  const v = sort.getRange(2, 1, sort.getLastRow() - 1, SORT_COLS).getValues();
  for (const r of v) {
    if (!(r[COL.TITLE - 1] || r[COL.MAP_URL - 1])) continue;
    s.total++;
    const reg = String(r[COL.REG - 1]);
    if (reg === '登録済') s.registered++;
    else if (reg.indexOf('重複') === 0) s.dup++;
    else if (isPendingRow_(r)) s.pendingFetch++;
    else if (reg.indexOf('保留') === 0) s.hold++;
    else s.undecided++;
  }
  return s;
}

function importWizard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const a = summarizeImport_(ss.getSheetByName(SHEET_IMPORT));
  const b = summarizeSort_(ss.getSheetByName(SHEET_SORT));

  if (a.badHeader) {
    ui.alert('取り込み', '「取り込み」シートの見出し(タイトル・URL)が見つかりません。1行目にTakeoutのCSVの見出しがあるか確認してください。', ui.ButtonSet.OK);
    return;
  }

  const status =
    `【取り込み】全${a.total}件(未読み込み ${a.pending}件)\n` +
    `【仕分け】全${b.total}件(取得待ち ${b.pendingFetch} / 未判定 ${b.undecided} / 保留 ${b.hold} / 登録済 ${b.registered} / 重複 ${b.dup})`;
  const yes = (title, msg) => ui.alert(title, status + '\n\n' + msg, ui.ButtonSet.YES_NO) === ui.Button.YES;

  if (a.pending > 0) {
    if (yes('ステップ①: 読み込み', '「取り込み」に未読み込みの行があります。①の読み込みを実行しますか?')) importSavedList();
    return;
  }
  if (b.pendingFetch > 0) {
    if (yes('ステップ①の続き: 情報取得', '「仕分け」に情報が未取得の行があります(取得の途中です)。続きを実行しますか?')) fetchPlaceData(true);
    return;
  }
  if (b.undecided + b.hold > 0) {
    if (yes('ステップ②→③: 登録',
      '「仕分け」に登録待ちの行があります。\n' +
      '②の人力の整理(typeの仕分け・データの修正・不要行の削除)が済んでいれば、③の登録を実行します。\n' +
      '保留の理由は「登録状態」列に出ています。③を実行しますか?')) registerFromSorting();
    return;
  }
  if (a.total > 0 || b.total > 0) {
    if (yes('ステップ④: 後片付け', 'すべて解決済みです。作業シート(取り込み・仕分け)を空にしますか?')) clearWorkSheets();
    return;
  }
  ui.alert('新規開始', '作業シートは空です。「取り込み」シートにTakeoutのCSVを貼ってから、もう一度「▶ 取り込みを進める」を実行してください。', ui.ButtonSet.OK);
}

// ==========================================
// 作業シート(取り込み・仕分け)の中身をすべて消す
//   全部解決したあとの後片付け用。未解決の行があれば警告し、2回確認する。
//   「仕分け」は見出し行・プルダウンを残し、「取り込み」は見出しごと空にする。
// ==========================================
function clearWorkSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const imp = ss.getSheetByName(SHEET_IMPORT);
  const sort = ss.getSheetByName(SHEET_SORT);
  if (!imp && !sort) { ui.alert('「取り込み」「仕分け」のシートが見つかりません。'); return; }

  // 未解決の行(登録済でも重複でもない行)を数える
  let total = 0, unresolved = 0;
  if (sort && sort.getLastRow() >= 2) {
    const v = sort.getRange(2, 1, sort.getLastRow() - 1, SORT_COLS).getValues();
    for (const r of v) {
      if (!(r[COL.TITLE - 1] || r[COL.MAP_URL - 1])) continue;
      total++;
      const reg = String(r[COL.REG - 1]);
      if (reg !== '登録済' && reg.indexOf('重複') !== 0) unresolved++;
    }
  }

  let msg = `「${SHEET_IMPORT}」と「${SHEET_SORT}」の中身をすべて消します。元に戻せません。\n(「${SHEET_SORT}」の見出し行とtypeのプルダウンは残します)\n\n「${SHEET_SORT}」の行数: ${total}件`;
  if (unresolved > 0) msg += `\n⚠ まだ解決していない行(保留・未処理)が ${unresolved}件あります。`;
  msg += '\n\n消しますか?';
  if (ui.alert('作業シートを空にする', msg, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  if (unresolved > 0 &&
      ui.alert('最終確認', `未解決の ${unresolved}件も消えます。本当に消しますか?`, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) { ui.alert('別の処理が実行中です。少し待ってからやり直してください。'); return; }
  try {
    if (imp) imp.clearContents();
    if (sort && sort.getMaxRows() >= 2) {
      sort.getRange(2, 1, sort.getMaxRows() - 1, sort.getMaxColumns()).clearContent();
    }
    SpreadsheetApp.flush();
    scheduleContinueIfPending_(false); // 取得待ちが無くなったので、続き実行のトリガーがあれば片付ける
  } finally {
    lock.releaseLock();
  }
  ss.toast('作業シートを空にしました。', '後片付け', 5);
}

// ==========================================
// 🖼 写真を探す(URLのページから画像の候補を集め、選んだものをドライブに保存して画像欄に入れる)
//   使い方: 「シート1」か「仕分け」で行を選ぶ → メニュー → URL(店の公式サイト・ブログ・画像のURL・ドライブの共有リンク)を入れる
//           → 候補から選ぶ → 保存。ドライブの「えらんで_写真」フォルダに保存し、サムネイルURLを画像欄(C列)に書く。
//   ※ 画像を読み込むのにJavaScriptが必要なページ(SNSなど)・ログインが必要なページ・ボット対策のあるページは、候補が出ないことがある。
//   ※ 他サイトの画像の保存・公開には著作権の問題がある。店の公式サイトや自分で撮った写真を使うこと。
// ==========================================
const PHOTO_FOLDER_NAME = 'えらんで_写真';
const PHOTO_MAX_CANDIDATES = 15;       // 候補として出す最大数
const PHOTO_MAX_SAVE = 5;              // 1行に保存する最大数
const PHOTO_MIN_BYTES = 5 * 1024;      // これより小さい画像は除外(アイコン・計測用の1pxなど)
const PHOTO_MAX_BYTES = 8 * 1024 * 1024;
const FETCH_UA = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36';

function decodeHtmlEntities_(s) {
  return String(s).replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#39;|&apos;/g, "'")
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&#x2F;/gi, '/').replace(/&#47;/g, '/');
}

// タグ文字列から属性の値を取り出す(data-src と src を取り違えないよう、直前は空白に限る)
function getAttr_(tag, name) {
  const re = new RegExp('\\s' + name + '\\s*=\\s*(?:"([^"]*)"|\'([^\']*)\'|([^\\s>]+))', 'i');
  const m = String(tag).match(re);
  if (!m) return '';
  return decodeHtmlEntities_(m[1] !== undefined ? m[1] : (m[2] !== undefined ? m[2] : m[3]));
}

// 相対URLを絶対URLにする(Apps Script には URL クラスが無いので自前で処理)
function resolveUrl_(base, ref) {
  ref = String(ref || '').trim();
  if (!ref || /^(data:|javascript:|#)/i.test(ref)) return '';
  if (/^https?:\/\//i.test(ref)) return ref;
  const m = String(base).match(/^(https?:)\/\/([^\/?#]+)([^?#]*)/i);
  if (!m) return '';
  if (ref.indexOf('//') === 0) return m[1] + ref;
  if (ref.charAt(0) === '/') return m[1] + '//' + m[2] + ref;
  const dir = (m[3] || '/').replace(/[^\/]*$/, '');
  const segs = (dir + ref).split('/');
  const out = [];
  for (const seg of segs) {
    if (seg === '..') { if (out.length > 1) out.pop(); }
    else if (seg !== '.') out.push(seg);
  }
  return m[1] + '//' + m[2] + out.join('/');
}

// srcset("a.jpg 480w, b.jpg 960w")から一番大きいものを選ぶ
function bestFromSrcset_(srcset) {
  let best = '', bestW = -1;
  String(srcset).split(',').forEach(part => {
    const t = part.trim().split(/\s+/);
    if (!t[0]) return;
    const w = t[1] ? parseFloat(t[1]) : 0;
    if (w >= bestW) { bestW = w; best = t[0]; }
  });
  return best;
}

function collectJsonLdImages_(node, depth) {
  const out = [];
  depth = depth || 0;
  if (!node || depth > 6) return out;
  if (Array.isArray(node)) { node.forEach(n => out.push(...collectJsonLdImages_(n, depth + 1))); return out; }
  if (typeof node !== 'object') return out;
  const img = node.image;
  if (typeof img === 'string') out.push(img);
  else if (Array.isArray(img)) img.forEach(i => { if (typeof i === 'string') out.push(i); else if (i && i.url) out.push(i.url); });
  else if (img && img.url) out.push(img.url);
  Object.keys(node).forEach(k => { if (k !== 'image' && node[k] && typeof node[k] === 'object') out.push(...collectJsonLdImages_(node[k], depth + 1)); });
  return out;
}

// imgタグ由来の候補から、ロゴ・アイコン・装飾らしいものを除く
function isJunkImageUrl_(u) {
  return /\.(svg|ico|gif)(\?|#|$)/i.test(u) ||
    /(logo|icon|sprite|favicon|avatar|pixel|spacer|blank|tracking|badge|button|loading|banner_ad)/i.test(u.split('?')[0]);
}

// HTMLから画像の候補を集める。優先順: og:image等 → JSON-LD → imgタグ
function extractImageUrls_(html, baseUrl) {
  const out = [];
  const seen = {};
  const add = (ref, source) => {
    const u = resolveUrl_(baseUrl, decodeHtmlEntities_(ref));
    if (!u || seen[u]) return;
    if (source === 'img' && isJunkImageUrl_(u)) return;
    seen[u] = true;
    out.push({ url: u, source: source });
  };

  (String(html).match(/<meta\b[^>]*>/gi) || []).forEach(tag => {
    const key = (getAttr_(tag, 'property') || getAttr_(tag, 'name')).toLowerCase();
    if (/^(og:image(:secure_url|:url)?|twitter:image(:src)?)$/.test(key)) add(getAttr_(tag, 'content'), 'og');
  });
  (String(html).match(/<link\b[^>]*>/gi) || []).forEach(tag => {
    if (getAttr_(tag, 'rel').toLowerCase() === 'image_src') add(getAttr_(tag, 'href'), 'og');
  });
  (String(html).match(/<script\b[^>]*application\/ld\+json[^>]*>[\s\S]*?<\/script>/gi) || []).forEach(block => {
    const body = block.replace(/^<script[^>]*>/i, '').replace(/<\/script>$/i, '');
    try { collectJsonLdImages_(JSON.parse(body)).forEach(u => add(u, 'json-ld')); } catch (e) {}
  });
  (String(html).match(/<img\b[^>]*>/gi) || []).forEach(tag => {
    const w = parseInt(getAttr_(tag, 'width'), 10), h = parseInt(getAttr_(tag, 'height'), 10);
    if ((w && w < 150) || (h && h < 150)) return; // 小さい画像は除外
    const srcset = getAttr_(tag, 'srcset') || getAttr_(tag, 'data-srcset');
    let ref = srcset ? bestFromSrcset_(srcset) : '';
    if (!ref) ref = getAttr_(tag, 'data-src') || getAttr_(tag, 'data-original') || getAttr_(tag, 'data-lazy-src') || getAttr_(tag, 'src');
    if (ref) add(ref, 'img');
  });
  return out.slice(0, PHOTO_MAX_CANDIDATES);
}

function extractDriveId_(url) {
  if (!/^https?:\/\/(drive|docs)\.google\.com\//i.test(url)) return '';
  const m = url.match(/\/d\/([A-Za-z0-9_-]{10,})/) || url.match(/[?&]id=([A-Za-z0-9_-]{10,})/);
  return m ? m[1] : '';
}

function driveThumb_(id) {
  return 'https://drive.google.com/thumbnail?sz=w800&id=' + id;
}

// ダイアログから呼ばれる: URLのページ(または画像・ドライブのリンク)から画像の候補を返す
function findImageCandidates(pageUrl) {
  const url = String(pageUrl || '').trim();
  if (!/^https?:\/\//i.test(url)) throw new Error('http:// または https:// で始まるURLを入力してください。');

  const driveId = extractDriveId_(url);
  if (driveId) return [{ url: driveThumb_(driveId), source: 'drive' }];

  const res = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true, followRedirects: true,
    headers: { 'User-Agent': FETCH_UA, 'Accept-Language': 'ja,en;q=0.8' }
  });
  const code = res.getResponseCode();
  if (code >= 400) throw new Error('ページを取得できませんでした(HTTP ' + code + ')。');
  const headers = res.getHeaders() || {};
  const ct = String(headers['Content-Type'] || headers['content-type'] || '').toLowerCase();
  if (ct.indexOf('image/') === 0) return [{ url: url, source: 'direct' }];

  const list = extractImageUrls_(res.getContentText(), url);
  if (list.length === 0) throw new Error('このページから画像が見つかりませんでした(SNSなど、画像を後から読み込むページは取れないことがあります)。');
  return list;
}

function getPhotoFolder_() {
  const it = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER_NAME);
}

// ダイアログから呼ばれる: 選んだ画像をドライブに保存し、画像欄(C列)に書く
//   mode: 'replace'=置き換え / 'append'=今ある写真に追加
function saveSelectedImages(sheetName, row, urls, mode) {
  if ([SHEET_MAIN, SHEET_SORT].indexOf(sheetName) < 0) throw new Error('対象のシートが正しくありません。');
  row = parseInt(row, 10);
  if (!(row >= 2)) throw new Error('対象の行が正しくありません。');
  urls = (Array.isArray(urls) ? urls : []).map(u => String(u)).filter(u => /^https?:\/\//i.test(u)).slice(0, PHOTO_MAX_SAVE);
  if (urls.length === 0) throw new Error('保存する写真が選ばれていません。');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('別の処理が実行中です。少し待ってからやり直してください。');
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) throw new Error('シートが見つかりません。');
    const title = String(sheet.getRange(row, COL.TITLE).getValue() || '写真');
    const folder = getPhotoFolder_();

    const saved = [];
    const failed = [];
    urls.forEach((u, i) => {
      try {
        const id = extractDriveId_(u);
        if (id || /^https:\/\/drive\.google\.com\/thumbnail\?/i.test(u)) { // ドライブの画像は保存し直さず、サムネイルURLにそろえる
          saved.push(id ? driveThumb_(id) : u);
          return;
        }
        const res = UrlFetchApp.fetch(u, { muteHttpExceptions: true, followRedirects: true, headers: { 'User-Agent': FETCH_UA } });
        if (res.getResponseCode() >= 400) { failed.push(`${i + 1}枚目: 取得できませんでした(HTTP ${res.getResponseCode()})`); return; }
        const blob = res.getBlob();
        const ct = String(blob.getContentType() || '').toLowerCase();
        const size = blob.getBytes().length;
        if (ct.indexOf('image/') !== 0) { failed.push(`${i + 1}枚目: 画像ではありませんでした`); return; }
        if (size < PHOTO_MIN_BYTES) { failed.push(`${i + 1}枚目: 小さすぎる画像です`); return; }
        if (size > PHOTO_MAX_BYTES) { failed.push(`${i + 1}枚目: 大きすぎる画像です`); return; }
        const ext = (ct.split('/')[1] || 'jpg').replace(/[^a-z0-9]/g, '').replace('jpeg', 'jpg');
        blob.setName(`${title}_${Date.now()}_${i + 1}.${ext}`);
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        saved.push(driveThumb_(file.getId()));
      } catch (e) {
        failed.push(`${i + 1}枚目: 保存に失敗しました`); // エラー文にURL等が混ざらないよう、詳細はログのみ
        console.error(String(e));
      }
    });

    if (saved.length === 0) return { saved: 0, failed: failed, img: String(sheet.getRange(row, COL.IMG).getValue() || '') };

    const current = String(sheet.getRange(row, COL.IMG).getValue() || '');
    const existing = noPhoto_(current) ? [] : current.split(',').map(x => x.trim()).filter(Boolean);
    const all = (mode === 'append' ? existing.concat(saved) : saved).slice(0, PHOTO_MAX_SAVE);
    sheet.getRange(row, COL.IMG).setValue(all.join(','));
    sheet.getRange(row, COL.STATUS).setValue('成功(写真を手動で追加)'); // 登録時に「失敗行」として再取得・消去されないようにする
    SpreadsheetApp.flush();
    return { saved: saved.length, failed: failed, img: all.join(',') };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// Googleの写真URL(lh3.googleusercontent.com)をドライブへ移行(v1.41)
//   Places の写真URLは一定期間で変わる/切れるため、画像を取り込んでドライブに保存し、
//   サムネイルURL(drive.google.com/thumbnail)に置き換える。Places API は呼ばない(追加料金なし)。
//   対象: シート1・仕分けの画像欄(C列)。ドライブ/他サイトのURLはそのまま残す。
//   消えているURL(HTTP 400/403/404/410)は外す。全部消えた行は「写真なし」にする(後で「写真を探す」で付けられる)。
//   一時的なエラー(5xx・429・通信失敗)のURLは残す(再実行で拾う)。6分制限は続き実行トリガーで越える。
// ==========================================
const MIGRATE_HANDLER = 'migratePhotosContinue';
const MAX_MIGRATE_CONTINUES = 150;     // 続き実行の安全上限
const MIGRATE_ABORT_STREAK = 15;       // 一時的なエラーがこの行数つづいたら中断(制限・障害でむだに回らないように)
const GOOGLE_PHOTO_RE = /^https?:\/\/lh\d*\.googleusercontent\.com\//i;
const DEAD_HTTP_CODES = [400, 403, 404, 410];

// ドライブ保存の例外から、容量不足('storage')・1日の上限('daily')を見分ける。それ以外は ''
function driveFatalKind_(e) {
  const m = String(e && e.message ? e.message : e);
  if (/too many times|Service invoked|1日|上限/i.test(m)) return 'daily';
  if (/storage|quota|容量|空き/i.test(m)) return 'storage';
  return '';
}

// 画像のblobを検査してドライブに保存し、サムネイルURLを返す。画像として不適切なら ''(ドライブ側の例外はそのまま投げる)
let photoFolderCache_ = null;
function storeBlobInDrive_(blob, baseName) {
  const ct = String(blob.getContentType() || '').toLowerCase();
  const size = blob.getBytes().length;
  if (ct.indexOf('image/') !== 0 || size < PHOTO_MIN_BYTES || size > PHOTO_MAX_BYTES) return '';
  const ext = (ct.split('/')[1] || 'jpg').replace(/[^a-z0-9]/g, '').replace('jpeg', 'jpg');
  blob.setName(`${baseName}_${Date.now()}.${ext}`);
  if (!photoFolderCache_) photoFolderCache_ = getPhotoFolder_();
  const file = photoFolderCache_.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return driveThumb_(file.getId());
}

// URLの画像を取得してドライブに保存し、サムネイルURLを返す。失敗したら ''(例外は出さない)
function savePhotoUrlToDrive_(url, baseName) {
  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true, headers: { 'User-Agent': FETCH_UA } });
    if (res.getResponseCode() >= 400) return '';
    return storeBlobInDrive_(res.getBlob(), baseName);
  } catch (e) {
    console.error(String(e)); // URL等が混ざりうるのでシートには書かない
    return '';
  }
}

function isGooglePhotoUrl_(u) { return GOOGLE_PHOTO_RE.test(String(u || '').trim()); }
function hasGooglePhoto_(img) { return String(img || '').split(',').some(isGooglePhotoUrl_); }

function clearMigrateTriggers_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === MIGRATE_HANDLER) ScriptApp.deleteTrigger(t);
  });
}

function countGooglePhotoRows_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let n = 0;
  [SHEET_MAIN, SHEET_SORT].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return;
    n += sh.getRange(2, COL.IMG, sh.getLastRow() - 1, 1).getValues().filter(r => hasGooglePhoto_(r[0])).length;
  });
  return n;
}

// 1行ぶんの移行。{img: 新しい画像欄, saved, dead, transient} を返す。ここでは書き込まない
function migrateRowPhotos_(folder, title, img) {
  const urls = String(img || '').split(',').map(x => x.trim()).filter(Boolean);
  const targets = [];
  urls.forEach((u, i) => { if (isGooglePhotoUrl_(u)) targets.push(i); });
  const out = urls.slice();
  const drop = {};
  let saved = 0, dead = 0, transient = 0, fatal = '';

  let responses = [];
  try {
    responses = UrlFetchApp.fetchAll(targets.map(i => ({ url: urls[i], muteHttpExceptions: true, followRedirects: true, headers: { 'User-Agent': FETCH_UA } })));
  } catch (e) {
    console.error(String(e));
    return { img: img, saved: 0, dead: 0, transient: targets.length };
  }

  targets.forEach((idx, k) => {
    if (fatal) { transient++; return; } // 容量不足などが分かったら、残りの保存は試さない
    try {
      const res = responses[k];
      const code = res.getResponseCode();
      if (DEAD_HTTP_CODES.indexOf(code) >= 0) { drop[idx] = true; dead++; return; }
      if (code >= 400) { transient++; return; }
      const thumb = storeBlobInDrive_(res.getBlob(), `${title}_${idx + 1}`);
      if (!thumb) { transient++; return; }
      out[idx] = thumb;
      saved++;
    } catch (e) {
      transient++;
      fatal = driveFatalKind_(e);
      console.error(String(e)); // URL等が混ざりうるのでシートには書かない
    }
  });

  const kept = out.filter((u, i) => !drop[i]);
  return { img: kept.join(','), saved: saved, dead: dead, transient: transient, allGone: urls.length > 0 && kept.length === 0, fatal: fatal };
}

function migrateSheets_(start) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const folder = getPhotoFolder_();
  photoFolderCache_ = folder;
  const total = { saved: 0, dead: 0, transient: 0 };
  let timedOut = false, aborted = false, streak = 0, fatal = '';

  for (const name of [SHEET_MAIN, SHEET_SORT]) {
    if (timedOut || aborted) break;
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) continue;
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, COL.IMG).getValues();
    for (let i = 0; i < data.length; i++) {
      const img = data[i][COL.IMG - 1];
      if (!hasGooglePhoto_(img)) continue;
      if (Date.now() - start > TIME_LIMIT_MS) { timedOut = true; break; }

      const row = i + 2;
      const r = migrateRowPhotos_(folder, String(data[i][COL.TITLE - 1] || '写真'), img);
      total.saved += r.saved; total.dead += r.dead; total.transient += r.transient;

      if (r.saved > 0 || r.dead > 0) {
        if (r.allGone) {
          sheet.getRange(row, COL.IMG).setValue('写真なし');
          sheet.getRange(row, COL.STATUS).setValue('成功(写真なし)'); // 「失敗」で始めない(再試行でAPI課金が走らないように)
        } else {
          sheet.getRange(row, COL.IMG).setValue(r.img);
        }
      }
      if (r.fatal) { fatal = r.fatal; aborted = true; break; } // この行のぶんは書き込み済み。ここで止める
      streak = (r.saved === 0 && r.dead === 0 && r.transient > 0) ? streak + 1 : 0;
      if (streak >= MIGRATE_ABORT_STREAK) { aborted = true; break; }
    }
  }
  SpreadsheetApp.flush();
  return { total: total, timedOut: timedOut, aborted: aborted, fatal: fatal };
}

function migratePhotosContinue() {
  migratePhotosToDrive(true);
}

function migratePhotosToDrive(isBackground) {
  const background = isBackground === true;
  const ui = SpreadsheetApp.getUi();
  if (!background) {
    const n = countGooglePhotoRows_();
    if (n === 0) { Browser.msgBox('移行する写真(Googleの写真URL)はありませんでした。'); return; }
    const ans = ui.alert('📦 写真をドライブへ移行',
      `Googleの写真URLが残っている行: ${n}行\n\n画像をドライブ(「${PHOTO_FOLDER_NAME}」フォルダ)に保存し、画像欄をドライブのURLに置き換えます。\n・Places APIは使いません(追加料金なし)\n・すでに消えているURLは外します(全部消えた行は「写真なし」)\n・6分を超えるぶんは、自動で続きを実行します\n\n始めますか？`,
      ui.ButtonSet.YES_NO);
    if (ans !== ui.Button.YES) return;
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    if (!background) Browser.msgBox('別の処理が実行中です。少し待ってからやり直してください。');
    return;
  }
  try {
    const props = PropertiesService.getScriptProperties();
    clearMigrateTriggers_();
    if (!background) props.deleteProperty('MIGRATE_COUNT');

    const r = migrateSheets_(Date.now());
    const left = countGooglePhotoRows_();
    const t = r.total;

    let again = false;
    if (r.timedOut && !r.aborted && left > 0) {
      const cont = Number(props.getProperty('MIGRATE_COUNT') || 0);
      if (cont < MAX_MIGRATE_CONTINUES) {
        props.setProperty('MIGRATE_COUNT', String(cont + 1));
        ScriptApp.newTrigger(MIGRATE_HANDLER).timeBased().after(60 * 1000).create();
        again = true;
      }
    }
    if (!again) props.deleteProperty('MIGRATE_COUNT');

    const msg = `移行 ${t.saved}枚 / 消えていた ${t.dead}枚 / 一時エラー ${t.transient}枚 / 残り ${left}行` +
      (again ? '(1分後に続きを自動実行します)' : '') +
      (r.fatal === 'storage' ? '\n※ ドライブの容量が足りません。空きを作ってから、もう一度実行してください(移行済みの行は飛ばして続きから)。'
        : r.fatal === 'daily' ? '\n※ Googleのドライブ操作の1日の上限に達しました。翌日にもう一度実行してください(続きから再開します)。'
        : r.aborted ? '\n※ 一時エラーが続いたため中断しました。しばらくしてからもう一度実行してください。' : '');
    console.log(msg);
    try {
      if (background) SpreadsheetApp.getActiveSpreadsheet().toast(msg, '写真の移行', 10);
      else Browser.msgBox(msg);
    } catch (e) { /* 画面が使えない実行では無視 */ }
  } finally {
    lock.releaseLock();
  }
}


function buildPhotoPickerHtml_(ctx) {
  const json = JSON.stringify(ctx).replace(/</g, '\\u003c');
  return '<!DOCTYPE html><html><head><base target="_top"><meta charset="utf-8">' +
    '<style>' +
    'body{font-family:sans-serif;margin:14px;font-size:14px;color:#222}' +
    'h3{margin:0 0 10px;font-size:16px}' +
    '#url{width:calc(100% - 110px);padding:6px;box-sizing:border-box}' +
    'button{padding:6px 12px;cursor:pointer}' +
    '#msg{margin:8px 0;min-height:20px}#msg.err{color:#c62828}' +
    '#grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:10px;margin-top:8px}' +
    '.card{border:4px solid #ddd;border-radius:8px;overflow:hidden;cursor:pointer;position:relative;background:#fafafa}' +
    '.card.sel{border-color:#1a73e8}' +
    '.card img{width:100%;height:110px;object-fit:cover;display:block;background:#eee}' +
    '.card .no{display:none;padding:8px;font-size:12px;color:#888}.card.broken img{display:none}.card.broken .no{display:block}' +
    '.card .n{position:absolute;top:4px;right:6px;background:#1a73e8;color:#fff;border-radius:10px;padding:0 7px;font-size:12px;display:none}.card.sel .n{display:block}' +
    '.card .src{font-size:11px;color:#666;padding:2px 6px}' +
    '#opts{margin-top:14px;padding-top:10px;border-top:1px solid #ddd}' +
    '</style></head><body>' +
    '<h3 id="t"></h3>' +
    '<div>ページのURL(店の公式サイト・ブログ・画像そのもののURL・ドライブの共有リンク):</div>' +
    '<input id="url" type="text"> <button id="find">候補を探す</button>' +
    '<div id="msg"></div><div id="grid"></div>' +
    '<div id="opts" style="display:none">' +
    '<label><input type="radio" name="mode" value="replace" checked> 今の写真を置き換える</label> ' +
    '<label><input type="radio" name="mode" value="append" id="mAppend"> 今の写真に追加する</label><br><br>' +
    '<button id="save">選んだ写真を保存</button> <span id="cnt"></span></div>' +
    '<script>' +
    'var CTX=' + json + ';var MAX=' + PHOTO_MAX_SAVE + ';var selected=[];' +
    'function $(id){return document.getElementById(id);}' +
    'function msg(t,err){var m=$("msg");m.textContent=t||"";m.className=err?"err":"";}' +
    '$("t").textContent=CTX.title+"("+CTX.sheet+" "+CTX.row+"行目)";' +
    '$("url").value=CTX.website||"";' +
    'if(CTX.hasPhoto){$("mAppend").checked=true;}' +
    'function refresh(){var cards=document.querySelectorAll(".card");for(var i=0;i<cards.length;i++){var u=cards[i].getAttribute("data-url");var k=selected.indexOf(u);cards[i].className="card"+(k>=0?" sel":"")+(cards[i].getAttribute("data-broken")?" broken":"");cards[i].querySelector(".n").textContent=k>=0?(k+1):"";}$("cnt").textContent=selected.length+"枚選択中(最大"+MAX+"枚)";$("save").disabled=selected.length===0;}' +
    'function render(list){selected=[];var g=$("grid");g.innerHTML="";for(var i=0;i<list.length;i++){(function(item){var c=document.createElement("div");c.className="card";c.setAttribute("data-url",item.url);' +
    'var img=document.createElement("img");img.setAttribute("referrerpolicy","no-referrer");img.src=item.url;img.onerror=function(){c.setAttribute("data-broken","1");refresh();};' +
    'var no=document.createElement("div");no.className="no";no.textContent="プレビューできません(選べば保存を試します)";' +
    'var n=document.createElement("span");n.className="n";var s=document.createElement("div");s.className="src";s.textContent=item.source;' +
    'c.appendChild(img);c.appendChild(no);c.appendChild(n);c.appendChild(s);' +
    'c.onclick=function(){var k=selected.indexOf(item.url);if(k>=0){selected.splice(k,1);}else if(selected.length<MAX){selected.push(item.url);}else{msg("選べるのは"+MAX+"枚までです。",true);return;}msg("");refresh();};' +
    'g.appendChild(c);})(list[i]);}' +
    '$("opts").style.display="block";msg(list.length+"件の候補が見つかりました。保存したい写真をクリックして選んでください(選んだ順に並びます)。");refresh();}' +
    '$("find").onclick=function(){var u=$("url").value.trim();if(!u){msg("URLを入力してください。",true);return;}msg("候補を探しています…");$("grid").innerHTML="";$("opts").style.display="none";' +
    'google.script.run.withSuccessHandler(render).withFailureHandler(function(e){msg(e&&e.message?e.message:String(e),true);}).findImageCandidates(u);};' +
    '$("save").onclick=function(){if(!selected.length)return;$("save").disabled=true;msg("保存しています…");var mode=document.querySelector("input[name=mode]:checked").value;' +
    'google.script.run.withSuccessHandler(function(r){var t=r.saved+"枚を保存しました。";if(r.failed&&r.failed.length){t+=" 保存できなかった写真: "+r.failed.join(" / ");}msg(t,r.saved===0);if(r.saved>0){setTimeout(function(){google.script.host.close();},1800);}else{$("save").disabled=false;}})' +
    '.withFailureHandler(function(e){msg(e&&e.message?e.message:String(e),true);$("save").disabled=false;}).saveSelectedImages(CTX.sheet,CTX.row,selected,mode);};' +
    'refresh();if(CTX.website){$("find").click();}' +
    '</script></body></html>';
}

// メニュー「🖼 写真を探す」: 選択中の行について、候補を選ぶ画面を開く
function openPhotoPicker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getActiveSheet();
  const name = sheet.getName();
  if ([SHEET_MAIN, SHEET_SORT].indexOf(name) < 0) {
    ui.alert('「シート1」か「仕分け」で、写真を付けたい行を選んでから実行してください。');
    return;
  }
  const row = sheet.getActiveRange().getRow();
  if (row < 2) { ui.alert('見出し行ではなく、写真を付けたい行を選んでから実行してください。'); return; }

  const vals = sheet.getRange(row, 1, 1, NUM_COLS).getValues()[0];
  const ctx = {
    sheet: name,
    row: row,
    title: String(vals[COL.TITLE - 1] || ''),
    website: String(vals[COL.WEBSITE - 1] || ''),
    hasPhoto: !noPhoto_(vals[COL.IMG - 1])
  };
  const output = HtmlService.createHtmlOutput(buildPhotoPickerHtml_(ctx)).setWidth(780).setHeight(680);
  ui.showModalDialog(output, '写真を探す');
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

    // 列構成 A〜N に合わせる（H=hideは空、I=map_url）
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
      '',             // L status
      '',             // M lat
      ''              // N lng
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
