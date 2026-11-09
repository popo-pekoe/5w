# setting.md — 設定項目一覧

`index.html` から読み取れる設定値と、システム全体(GAS・シート)で必要な設定のまとめ。
**このファイルは公開リポジトリに入る。APIキー等の秘密の値は絶対に書かないこと(キー名のみ記載)。**

## 1. index.html 内の設定(コード直書き)

| 項目 | 場所 | 現在値 | 説明 |
| :--- | :--- | :--- | :--- |
| `spreadsheetUrl` | script 冒頭(約220行) | Googleスプレッドシート「ウェブに公開」CSV URL (`.../pub?output=csv`) | データ取得元。公開用URLで読み取り専用。差し替えるとデータ元が変わる |
| `birthYear` / `birthMonth` | 約222-223行 | 2022 / 1 | 子の生年月。満年齢の自動計算に使用(`currentAge`)。年齢ロックの基準 |
| `currentTab` 初期値 | 約236行 | `'go_out'` | 起動時のタブ。HTMLの `.tab-btn.active` と揃える |
| `regionMap` | 約240-248行 | 地方→都道府県の対応表 | 場所フィルタのグループ分け。「ちかくのエリア（市）」= 横浜市/川崎市/藤沢市 |
| `targetCities` | 約286行 | 横浜市, 川崎市, 藤沢市 | 住所から市単位でも抽出する対象。`regionMap` の「ちかくのエリア」と揃える |
| `itemHeightDesktop` / `itemHeightMobile` | 約487-488行 | 200 / 120 (px) | スロットのリール1枠の高さ。CSSの `.reel-item` 高さと一致させること |
| モバイル判定幅 | `startSlotSpin` | `window.innerWidth <= 850` | CSSのメディアクエリ幅と揃える |
| `spinTime` / `delayBetweenReels` | `startSlotSpin` | 3000ms / 800ms | スロット回転時間と、リールごとの停止ずらし |
| `spinCycles` | `startSlotSpin` | `20 + i*10` | リールごとの回転枚数 |
| イージング | `startSlotSpin` | `cubic-bezier(0.1, 0.7, 0.1, 1.15)` | スロット停止時の動き |
| バージョン表記 | `.version-display` (約217行) | `v1.15` | 更新時に手動で上げる |

### 外部CDN(ビルド不要・URL直指定)
- Google Fonts: M PLUS Rounded 1c (500, 800)
- Font Awesome 6.4.0 (cdnjs)
- PapaParse 5.4.1 (cdnjs)

### タブ(カテゴリ)定義
`type` 値とタブの対応。タブ追加時は HTML の `.tab-btn`、CSS の `.tab-btn.active.<type>` 色、GAS側の分岐をセットで変更する。

| type | 表示 | 色 | GASの取得方式 |
| :--- | :--- | :--- | :--- |
| `eat` | 🍔 たべる | 橙 | Custom Search(画像のみ) |
| `lunch` | 🍱 昼食 | 紫 | Custom Search(画像のみ) |
| `go_out` | 🚗 お出かけ | 緑 | Google Maps API |
| `eat_out` | 🍽️ 外食 | ピンク | Google Maps API |

### 表示条件(`getAvailableItems`)
タブの `type` 一致 / 画像が1枚以上 / `hide` が `1` でない / お気に入りフィルタ / 場所フィルタ。
年齢ロックは別処理: `age_limit`(数値) > 現在年齢 の項目は鍵付きグレーアウト、スロット候補からも除外。

## 2. スプレッドシート(データ仕様)

シート名: `シート1`。1行目ヘッダー。GAS は列番号に依存するため **列の順序を変えない**。

| 列 | ヘッダー | 内容 |
| :--- | :--- | :--- |
| A | `type` | カテゴリID |
| B | `title` | 名称(API検索クエリにも使用) |
| C | `img` | 画像URL(カンマ区切りで複数) |
| D | `category` | サブカテゴリ |
| E | `comment` | 概要 |
| F | `website` | 公式サイト |
| G | `address` | 住所(場所フィルタの元) |
| H | `hide` | `1` で非表示(手入力のみ。GASは書かない) |
| I | `map_url` | マップURL |
| J | `age_limit` | 年齢下限(旧ヘッダー名 `years` → `age_limit` に変更) |

HTMLが読むキー名: `type, title, img, category, comment, website, address, hide`、
マップは `マップURL` / `map_url` / `mapUrl`、年齢は `対象年齢` / `age_limit` / `age`。

## 3. GAS(スプレッドシート紐づきスクリプト)

コードは `gas/Code.gs` に控えがある(正本は Apps Script エディタ側。修正は両方に反映する)。

| 設定 | 保存場所 | 備考 |
| :--- | :--- | :--- |
| `MAPS_API_KEY` | GAS「スクリプトプロパティ」 | Google Maps Platform。値はリポジトリ・ドキュメントに書かない |
| `SEARCH_ENGINE_ID` | GAS「スクリプトプロパティ」 | Custom Search の cx |
| フォーム送信トリガー | GASのトリガー設定 | `onFormSubmit(e)` |
| カスタムメニュー「🌟プレゼン準備」 | スプレッドシート | `fixBrokenImages()` でリンク切れ画像を修復 |

## 4. 公開・セキュリティ確認結果(2026-11-05 時点)

- リポジトリ追跡ファイル(`index.html`, `.gitignore`, `5w.code-workspace`)および git 全履歴に、APIキー・トークン・パスワードは**含まれていない**。
- 新形式キー(`AQ.` 始まり)・旧形式(`AIza`)・`key=` パラメータのいずれも、作業ツリー・git全履歴・公開CSVに存在しない。公開CSV(157行)にも含まれていない(画像は `lh3.googleusercontent.com` / `drive.google.com`)。
- `spreadsheetUrl` は公開前提の読み取り専用URL。HTMLから書き込みは不可。
- 注意: 公開CSVは誰でも取得できる。シートに個人情報・非公開メモを入れない。
- 注意: `5w.code-workspace` にローカルの Dropbox パス(`../../Dropbox/data/!作業用`)がある(秘密ではないが公開不要なら除外を検討)。
- 注意: `.gitignore` は現在 `.env` のみ。秘密ファイルを増やす場合は追記する。
- 注意: `overview.md` はまだ git 未追跡。コミットする場合も、キーの値を書かないこと(現状はキー名のみで問題なし)。

## 5. 既知の気になる点(要確認)

- (対応済 v1.15) `play` タブを削除し、初期タブを `go_out` にした。
- 現CSVの年齢列ヘッダーは `years`(`対象年齢`ではない)。HTMLは `対象年齢`/`age_limit`/`age` しか読まないため、年齢ロックが効いていない可能性がある。ヘッダー名を `age_limit` に直すか、HTMLに `years` を追加する。
- (v1.31で修正・Apps Script側へ要適用) GAS が `hide` 列(H)に `map_url` を書いてしまう不具合。原因: `onFormSubmit` の `appendRow` が H・I 両方に `mapUrl` を入れ、`fetchPlaceData` が H(8列目)へ書き込み・H を地図URLとして読む。修正案は引き継ぎ時のやり取り参照(I列=9列目に統一、Hは空)。
- (対応) 年齢列は `age_limit` に統一する(シートJ列ヘッダーを書き換え)。GASは列番号参照なので影響なし。
