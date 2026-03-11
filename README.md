# 引継ぎドキュメントシステム (Hikitsugi Doc XML)

チームごとに蓄積されるExcel引継ぎドキュメントをブラウザで閲覧・検索・編集できるナレッジベースシステムです。

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

```bash
# 依存ライブラリをインストール
pip install -r requirements.txt

# サンプルデータを生成（オプション）
python3 create_sample_data.py

# サーバーを起動
python3 app.py
```

ブラウザで `http://localhost:5000` を開いてください。

## ファイル構成

```
app.py                     # Flaskアプリケーション（APIルート）
excel_converter.py         # Excel → XML 変換ロジック
xml_exporter.py            # XML → Excel エクスポートロジック
create_sample_data.py      # デモ用サンプルデータ生成スクリプト
requirements.txt           # Python依存ライブラリ
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

