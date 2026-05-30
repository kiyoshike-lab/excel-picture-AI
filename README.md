# 建設FAX Excel AI 試作

建設業の資材発注書、見積依頼書、納品依頼FAXを読み取り、Excel入力用の項目へ自動整理する試作品です。

## できること

- FAXのOCR済みテキストを貼り付けて抽出
- 文字情報が入ったPDFを読み込み
- 写真・画像ファイルの読み込み口を用意
- 受信FAXフォルダを処理してExcelを自動作成
- 工事名、現場名、納入場所、納期、希望時間、担当者などを抽出
- 品番、資材名、数量、単位、単価、金額を明細行として抽出
- 画面上で修正してからExcel出力

## 使い方

### 画面で試す

```bat
start_fax_excel_ai.bat
```

ブラウザで `http://127.0.0.1:8765` が開きます。

### FAX受信フォルダを自動処理する

複合機の受信FAX保存先を `incoming_fax` フォルダにします。

```bat
start_fax_auto_watch.bat
```

新しいPDF、画像、テキストが入ると、`outputs` フォルダにExcelを作成し、処理済み原本を `processed_fax` に移動します。

## セットアップ

Python 3.11以上を想定しています。

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py 8765
```

## Renderで写真OCRつきデモを公開する

このリポジトリにはRender用の `Dockerfile` と `render.yaml` を入れています。

1. GitHubにこのリポジトリを上げる
2. Renderで「New Web Service」を作成
3. GitHubリポジトリを選択
4. EnvironmentはDockerを選択
5. Deploy

Docker内にTesseract OCR日本語パックを入れるため、Render上では写真・画像アップロードからOCR読み取りまで動く構成です。

## 写真OCRについて

写真・画像の読み取り口は実装済みです。ローカルPCで動かす場合はTesseract OCRを入れると、撮影した写真から文字を読んでExcel化できます。RenderのDocker構成ではTesseract OCRを自動で入れます。

GitHub Pages版は画面デモ用です。写真OCRや本格的なExcel生成は、Renderまたはローカルアプリで使います。

## 注意

この試作は、まず「AIが入力して、人が確認して確定する」業務フローを想定しています。納入場所、数量、金額などは現場影響が大きいため、最初から完全自動確定にしない設計です。

## 今後追加するとよいもの

- スキャンPDF、スマホ写真向けOCRの本格対応
- 会社ごとのFAXレイアウト学習
- 工事台帳、発注台帳、見積管理表など既存Excelへの直接転記
- 低信頼項目だけ確認する承認画面
- 処理履歴と原本ファイルの保管
