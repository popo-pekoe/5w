// GAS Version 1.31 (map_urlの書き込み先をI列に修正・H列=hideには書かない)
// メニューに修復ボタンを追加
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🌟プレゼン準備')
    .addItem('空欄のデータを自動取得（全部入り）', 'fetchPlaceData')
    .addItem('🔧 リンク切れ画像を自動修復', 'fixBrokenImages')
    .addToUi();
}

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

function fetchPlaceData(isBackground = false) { 
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('シート1');
  if (!sheet) return;
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  // 列構成: A type / B title / C img / D category / E comment / F website / G address / H hide / I map_url / J age_limit
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 10);
  const data = dataRange.getValues();
  
  const apiKey = PropertiesService.getScriptProperties().getProperty('MAPS_API_KEY');
  const cx = PropertiesService.getScriptProperties().getProperty('SEARCH_ENGINE_ID'); 
  if (!apiKey) return;
  
  let updateCount = 0;

  for (let i = 0; i < data.length; i++) {
    const type = data[i][0]; 
    let title = data[i][1]; 
    const currentImg = data[i][2]; 
    
    // ★ 修正: 地図URLはI列(9列目)のみ。H列はhide(非表示フラグ)なので読まない
    let rawI = data[i][8] ? String(data[i][8]) : '';
    if (rawI.includes('query=') || rawI.includes('maps/search')) rawI = '';
    const mapUrl = rawI; 
    
    if (!currentImg && (title || mapUrl)) {
      
      if (type === 'eat' || type === 'lunch') { 
        if (!cx) continue;
        if (!title) continue;

        const searchUrl = `https://customsearch.googleapis.com/customsearch/v1?cx=${cx}&q=${encodeURIComponent(title)}&searchType=image&key=${apiKey}&num=3`;
        const res = UrlFetchApp.fetch(searchUrl, {muteHttpExceptions: true});
        const json = JSON.parse(res.getContentText());
        
        if (json.items && json.items.length > 0) {
          const photoUrls = json.items.map(item => item.link);
          sheet.getRange(i + 2, 3).setValue(photoUrls.join(','));
          sheet.getRange(i + 2, 4).setValue('ごはん'); 
          updateCount++;
        }
        Utilities.sleep(500); 
        continue; 
      }

      let latLngStr = ''; 
      let urlName = ''; 

      if (mapUrl && mapUrl.includes('goo.gl')) {
        try {
          const options = { "method": "get", "followRedirects": false, "muteHttpExceptions": true };
          const res = UrlFetchApp.fetch(mapUrl, options);
          const locationUrl = res.getHeaders()['Location']; 
          if (locationUrl) {
            let latLngMatch = locationUrl.match(/@([0-9.,\-]+)/) || locationUrl.match(/[?&]ll=([0-9.,\-]+)/);
            if (latLngMatch) latLngStr = latLngMatch[1].split(',').slice(0, 2).join(',');

            let match = locationUrl.match(/\/maps\/place\/([^\/?]+)/) || locationUrl.match(/\/maps\/search\/([^\/?]+)/) || locationUrl.match(/[?&]q=([^&]+)/); 
            if (match && match[1]) {
              let decoded = decodeURIComponent(match[1].replace(/\+/g, ' '));
              if (!decoded.match(/^[0-9.,\-]+$/) && !decoded.includes('°')) urlName = decoded;
            }
          }
        } catch(e) {}
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
          const nsRes = UrlFetchApp.fetch(nsUrl, {muteHttpExceptions: true});
          searchJson = JSON.parse(nsRes.getContentText());
        }

        if (!searchJson.results || searchJson.results.length === 0) {
          let tsUrl = `https://maps.googleapis.com/maps/api/place/textsearch/json?query=${encodeURIComponent(apiSearchQuery)}&language=ja&key=${apiKey}`;
          if (latLngStr) tsUrl += `&location=${latLngStr}&radius=3000`;
          const tsRes = UrlFetchApp.fetch(tsUrl, {muteHttpExceptions: true});
          searchJson = JSON.parse(tsRes.getContentText());
        }
        
        if (searchJson.results && searchJson.results.length > 0) {
          if (latLngStr) {
            let minDistance = Infinity;
            let closestPlaceId = null;
            let closestName = '';
            
            const coords = latLngStr.split(',');
            const urlLat = parseFloat(coords[0]);
            const urlLng = parseFloat(coords[1]);

            for (let r of searchJson.results) {
              if (r.geometry && r.geometry.location) {
                const dist = getDistance(urlLat, urlLng, r.geometry.location.lat, r.geometry.location.lng);
                if (dist < minDistance) {
                  minDistance = dist;
                  closestPlaceId = r.place_id;
                  closestName = r.name;
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

      if (fallbackMode || !placeId) {
        let addressStr = '';
        if (latLngStr) {
          try {
            const geoUrl = `https://maps.googleapis.com/maps/api/geocode/json?latlng=${latLngStr}&language=ja&key=${apiKey}`;
            const geoRes = UrlFetchApp.fetch(geoUrl, {muteHttpExceptions: true});
            const geoJson = JSON.parse(geoRes.getContentText());
            if (geoJson.results && geoJson.results.length > 0) {
              addressStr = geoJson.results[0].formatted_address.replace(/^日本、\s*(〒\d{3}-\d{4}\s*)?/, '');
              const match = addressStr.match(/^(...?[都道府県]).*?[市区町村]/);
              if(match) addressStr = match[0];
            }
          } catch(e) {}
        }

        let categoryStr = 'おでかけスポット';
        if (type === 'eat' || type === 'eat_out') categoryStr = 'ごはん';

        sheet.getRange(i + 2, 2).setValue(cleanName || "名称不明");
        sheet.getRange(i + 2, 3).setValue('写真なし');
        sheet.getRange(i + 2, 4).setValue(categoryStr);
        sheet.getRange(i + 2, 7).setValue(addressStr); 
        sheet.getRange(i + 2, 9).setValue(mapUrl || ''); // ★ 修正: 8→9列目(I列 map_url)
        
        updateCount++;
        Utilities.sleep(500);
        continue; 
      }
      
      const detailsUrl = `https://maps.googleapis.com/maps/api/place/details/json?place_id=${placeId}&fields=name,photos,types,editorial_summary,website,formatted_address,url&language=ja&key=${apiKey}`;
      const detailsRes = UrlFetchApp.fetch(detailsUrl, {muteHttpExceptions: true});
      const detailsJson = JSON.parse(detailsRes.getContentText());
      
      if (detailsJson.result) {
        const result = detailsJson.result;
        
        let photoUrls = [];
        if (result.photos && result.photos.length > 0) {
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
           if(match) addressStr = match[0];
        }

        if (!title || title.includes('http') || title.includes('〒')) sheet.getRange(i + 2, 2).setValue(cleanName);
        sheet.getRange(i + 2, 3).setValue(photoUrls.length > 0 ? photoUrls.join(',') : '写真なし');
        sheet.getRange(i + 2, 4).setValue(categoryStr);
        sheet.getRange(i + 2, 6).setValue(result.website || '');
        sheet.getRange(i + 2, 7).setValue(addressStr);
        sheet.getRange(i + 2, 9).setValue(mapUrl ? mapUrl : (result.url || '')); // ★ 修正: 8→9列目(I列 map_url)
        
        updateCount++;
      } else {
        sheet.getRange(i + 2, 3).setValue('エラー: 詳細情報の取得に失敗しました');
      }
      Utilities.sleep(500); 
    }
  }
}

function onFormSubmit(e) {
  try {
    const mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('シート1'); 
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
        } catch(error) { finalPhotoUrl = 'エラー: 画像の権限変更に失敗しました'; }
      }
    }

    // ★ 修正: 列構成 A〜J に合わせる（H=hideは空、I=map_url）
    mainSheet.appendRow([
      type,      // A type
      title,     // B title
      finalPhotoUrl, // C img
      category,  // D category
      comment,   // E comment
      '',        // F website
      '',        // G address
      '',        // H hide
      mapUrl,    // I map_url
      ageLimit   // J age_limit
    ]);
    SpreadsheetApp.flush();

    if (!finalPhotoUrl && (title || mapUrl)) {
      try { fetchPlaceData(true); } catch(fetchError) {}
    }
  } catch(e) {}
}

// ==========================================
// リンク切れ画像を自動検知して修復するプログラム
// ==========================================
function fixBrokenImages() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('シート1');
  if (!sheet) return;
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  // A列からC列（写真URL）までを取得
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 3);
  const data = dataRange.getValues();
  
  let brokenCount = 0;
  
  // 1. リンク切れの画像をチェックして空欄にする
  for (let i = 0; i < data.length; i++) {
    const imgStr = data[i][2]; // C列（写真）
    
    // Googleの画像URLが入っている場合のみチェック
    if (imgStr && imgStr.includes('googleusercontent.com')) {
      const firstUrl = imgStr.split(',')[0].trim(); // 複数ある場合は最初の1枚目でテスト
      
      try {
        // 画像URLにアクセスを試みる
        const res = UrlFetchApp.fetch(firstUrl, { muteHttpExceptions: true });
        
        // 400エラー等（リンク切れ）が返ってきた場合
        if (res.getResponseCode() >= 400) {
          console.log(`リンク切れ検知 (${i + 2}行目): ${firstUrl}`);
          sheet.getRange(i + 2, 3).setValue(''); // C列を空欄にする
          brokenCount++;
        }
      } catch (e) {
        // 接続エラー等の場合も空欄にする
        console.log(`アクセスエラー (${i + 2}行目)`);
        sheet.getRange(i + 2, 3).setValue('');
        brokenCount++;
      }
    }
  }
  
  // 2. 空欄にしたデータを再取得する
  if (brokenCount > 0) {
    SpreadsheetApp.flush(); // 一旦スプレッドシートに変更を反映
    
    // 既存の「fetchPlaceData」を裏側で呼び出し、空欄になった場所だけ再取得させる！
    fetchPlaceData(true); 
    
    Browser.msgBox(`完了！\n${brokenCount}件のリンク切れ画像を検知し、最新の画像に更新しました。`);
  } else {
    Browser.msgBox('リンク切れの画像はありませんでした。すべて正常です！');
  }
}
