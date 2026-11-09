# CLAUDE.md

家族のお出かけ・食事を決める「えらんで！プレゼンター」。詳細な仕様は @◆プレゼンシステム開発の共同制作overview.md、設定値は @setting.md を参照。

## 構成
- `index.html` 1ファイル完結のSPA(HTML/CSS/Vanilla JS)。ビルド・パッケージ管理なし。
- データ: Googleフォーム → スプレッドシート → GAS(API補完) → 「ウェブに公開(CSV)」→ `Papa.parse` で読込。HTMLは読み取り専用。
- GASのコードは `gas/Code.gs`(Apps Script側の控え)。実際に動くのはApps Script側。

## 開発・確認
- ビルド/テストなし。`index.html` をブラウザで開く(CSVはネット経由で取得)。
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
- 先頭のバージョンコメント(`// GAS Version x.xx`)を更新して、両者が揃っているか確認できるようにする。
- 修正案を出すときは「変更箇所の指摘 → ファイル全体」の順で出す(部分コピペは認知負荷が高いため)。
- 公開リポジトリなので、キーの値・シートID以外の秘密・個人情報を `gas/` に絶対書かない(キーは `PropertiesService` のみ)。
- 秘密のデバッグ用直書きをコミットしない。万一pushしたら、履歴から消すだけでなくキーを再発行する。

## 守ること
- **APIキー等の秘密を、コード・md・コミットに書かない**(キー名 `MAPS_API_KEY`/`SEARCH_ENGINE_ID` のみ可。値はGASのスクリプトプロパティ)。
- `spreadsheetUrl` は公開CSVのURL。シートに個人情報を入れない。
- スプレッドシートの列順はGASが依存するので変えない。
- 日本語UI(子供向けのひらがな中心の文言)を維持する。

## 既知の課題
`setting.md` の「既知の気になる点」を参照(年齢列が `years`、`hide`/`map_url` の列ずれ、等)。
