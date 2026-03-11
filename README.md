# 引継ぎドキュメントシステム (Hikitsugi Doc XML)

チームごとに蓄積されるExcel引継ぎドキュメントをブラウザで閲覧・検索・編集できるナレッジベースシステムです。  
バックエンドは **Rust / Axum** で実装されています。

## 機能

| 機能 | 説明 |
|------|------|
| 📂 階層ナビゲーション | チーム → 作業コード → ドキュメントのツリー表示 |
| 📊 スプレッドシートビュー | ExcelをXMLに変換してブラウザで表示（書式・結合セル・画像保持）|
| 🗂️ XMLソースエディタ | CodeMirrorによるXML構文ハイライト付きエディタ |
| 🔍 全文検索 | 全チーム・全ドキュメントを横断して検索（ヒット箇所ハイライト）|
| ✏️ セル編集 | 編集モードでセルを直接クリックして値を変更 |
| 💾 保存 | 編集したXMLをキャッシュに保存 |
| ⬇️ Excelエクスポート | XMLからExcel (.xlsx) を再生成してダウンロード |
| 🔄 自動反映 | フォルダが増えてもリロードで即反映 |
| ⚡ キャッシュ | MD5ハッシュで変更検知・未変更ファイルは再変換スキップ |

## フォルダ構造

```
data/watch/
  チーム1/
    AAA-01/
      AAA-01作業引継ぎドキュメント.xlsx
  チーム2/
    BBB-01/
      BBB-01作業引継ぎドキュメント.xlsx
    CBA-A-01/
      CBA-A-01作業引継ぎドキュメント.xlsx
```

## セットアップ

### 必要なもの

- Rust toolchain 1.75 以上 ([rustup](https://rustup.rs/) でインストール)

### ビルドと起動

```bash
# サンプルデータを生成（オプション・Python + openpyxl が必要）
python3 create_sample_data.py

# ビルド
cargo build --release

# サーバーを起動（ポート 5000）
./target/release/hikitsugi-doc-xml

# ポートを変更する場合
PORT=8080 ./target/release/hikitsugi-doc-xml
```

ブラウザで `http://localhost:5000` を開いてください。

### 開発時の起動

```bash
RUST_LOG=debug cargo run
```

## ファイル構成

```
Cargo.toml                 # Rust依存ライブラリ
src/
  main.rs                  # Axum Webサーバー・APIルート
  converter.rs             # Excel → XML 変換ロジック（xlsx ZIP解析）
  exporter.rs              # XML → Excel エクスポートロジック
create_sample_data.py      # デモ用サンプルデータ生成スクリプト（Python）
data/watch/                # Excelファイルを配置するフォルダ
cache/                     # 変換済みXMLのキャッシュ（自動生成）
templates/index.html       # フロントエンドHTML（SPA）
static/css/style.css       # UIスタイル
static/js/app.js           # フロントエンドJavaScript
static/vendor/codemirror/  # CodeMirrorエディタ（ローカル）
```

## API エンドポイント

| メソッド | パス | 説明 |
|--------|------|------|
| GET | `/` | メインページ |
| GET | `/api/structure` | チーム・コード・ドキュメント一覧 |
| GET | `/api/document/<team>/<code>/<file>` | XMLドキュメント取得 |
| POST | `/api/document/<team>/<code>/<file>/save` | 編集済みXML保存 |
| GET | `/api/document/<team>/<code>/<file>/export` | Excelダウンロード |
| GET | `/api/search?q=<query>` | 全文検索 |

## XMLフォーマット

Excelファイルは以下の情報を保持したカスタムXML形式で保存されます：

- セル値・データ型
- フォント（太字・斜体・下線・サイズ・色）
- 背景色・罫線
- 配置（横・縦・折り返し）
- 結合セル
- 行の高さ・列の幅
- 埋め込み画像（Base64エンコード）

## 技術スタック

| 役割 | ライブラリ |
|------|----------|
| Webフレームワーク | [Axum](https://github.com/tokio-rs/axum) 0.7 + Tokio |
| 静的ファイル配信 | tower-http ServeDir |
| xlsx解析 | zip 2 + quick-xml 0.36（ZIPアーカイブ直接解析）|
| xlsxエクスポート | [rust_xlsxwriter](https://github.com/jmcnamara/rust_xlsxwriter) 0.68 |
| JSONシリアライズ | serde + serde_json |
| キャッシュ検証 | md5 |
| 画像エンコード | base64 |

