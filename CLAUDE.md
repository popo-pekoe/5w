# CLAUDE.md

家族のお出かけ・食事を決める「えらんで！プレゼンター」。詳細な仕様は @overview.md、設定値・取り込み手順・既知の課題は @setting.md を参照。

## 構成
- `index.html` 1ファイル完結のSPA(HTML/CSS/Vanilla JS)。ビルド・パッケージ管理なし。
- データ: Googleフォーム → スプレッドシート → GAS(API補完) → 「ウェブに公開(CSV)」→ `Papa.parse` で読込。HTMLは読み取り専用。
- GASのコードは `gas/Code.gs`(Apps Script側の控え)。実際に動くのはApps Script側。

## 開発・確認
- `index.html`: ビルドなし。ブラウザで開いて確認(CSVはネット経由で取得)。
- GAS: `node gas/test/run.js`(Google側をモックにしたテスト。Nodeのみ・キー不要)。`gas/Code.gs` を直したら実行する。
- 変更したら +0.01 でバージョンを上げる。表記は `index.html` 先頭のコメントと `.version-display` の2か所(詳細は @AGENTS.md の「バージョン管理ルール」)。
- 反映は `index.html` をホスティング先に上書きするだけ。

## コードの要点(index.html)
- 状態: `data`(全件) / `currentDisplayData`(表示中) / `currentTab` / `activeLocationFilters` / `showOnlyLiked`。
- 年齢計算は `birthYear`/`birthMonth` のハードコード。`age_limit` 未満はロック表示+スロット対象外。
- スロット(`startSlotSpin`): リール位置リセット → `void offsetHeight` で強制リフロー → 二重 `requestAnimationFrame` で要素追加。この順序は描画最適化回避のため意図的。崩さない。
- リール高さ(JSの `itemHeight*`)・モバイル幅850px はCSSと揃える。
- タブ追加は HTML(`.tab-btn`+CSS色) と GAS と シート/フォームの3箇所セット。
- データはテンプレート文字列で `innerHTML` に入れている(信頼できるシート前提)。

## GASの運用(gas/Code.gs)
- 正本は2つある: Apps Script(実行される方)と `gas/Code.gs`(履歴用の控え)。**修正は両方に反映する**。手動コピペで同期する。
- 先頭のバージョンコメント(`// GAS Version x.xx`)を +0.01 して、両者が揃っているか確認できるようにする。
- 修正案を出すときは「変更箇所の指摘 → ファイル全体」の順で出す(部分コピペは認知負荷が高いため)。
- 公開リポジトリなので、キーの値・個人情報を `gas/` に絶対書かない(キーは `PropertiesService` のみ)。秘密のデバッグ用直書きをコミットしない。万一pushしたら、履歴から消すだけでなくキーを再発行する。
- 列構成(シート1・仕分け共通): A type / B title / C img / D category / E comment / F website / G address / H hide / I map_url / J age_limit / K gtypes / L status / M lat / N lng(仕分けのみ O 登録状態)。GASは列番号で参照する(`COL` 定数)。
- 保存リストの取り込みは3シート・2段階: 取り込み(CSV貼付) →①→ 仕分け(人力で整理) →③→ シート1。入口はメニュー「▶ 取り込みを進める」(作業シートにデータが残っていれば途中とみなし、続きの段階を判定する)。詳細は @setting.md。
- 重複判定は座標の距離 `DUP_DISTANCE_M`(50m)。場所IDが同じ/店名が同じ相手を距離で判定し、座標が不明な同名は保留にする。
- 写真の手動追加はメニュー「🖼 写真を探す」(`openPhotoPicker` → ダイアログ → `findImageCandidates` / `saveSelectedImages`)。画像はドライブに保存してサムネイルURLを画像欄に書く。他サイトの画像を扱うため著作権に注意。
- 6分制限: `TIME_LIMIT_MS` で新しい行への着手を止め、未処理が残る間は続き実行トリガー(1つだけ・自動削除)で再開する。`LockService` で二重実行を防ぐ。
- `status` 列は公開CSVに出る。APIのエラー文(URL・キーが混ざりうる)をシートに書かない(例外の詳細は `console.error` のみ)。

## 守ること
- **APIキー等の秘密を、コード・md・コミットに書かない**(キー名 `MAPS_API_KEY`/`SEARCH_ENGINE_ID` のみ可。値はGASのスクリプトプロパティ)。
- `spreadsheetUrl` は公開CSVのURL。シートに個人情報を入れない。
- スプレッドシートの列順はGASが依存するので変えない(列を足すときは末尾に)。
- 日本語UI(子供向けのひらがな中心の文言)を維持する。

## 既知の課題
`setting.md` の「既知の気になる点」を参照。
