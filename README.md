# 建設FAXかんたんExcel入力

建設業のFAX、PDF、スマホ写真を読み取り、社員が画面で確認したものだけExcel台帳へ反映する試作アプリです。

## できること

- FAX複合機が保存したPDFを自動で取り込み
- 写真、画像、PDF、Word、Excel、メール、ZIP、テキストを読み取り
- 工事名、現場名、納入場所、納期、希望時間、担当者を抽出
- 品番、資材名、数量、単位、単価、金額を明細化
- 確認待ち一覧に保存
- 社員が画面で確認、修正
- OKしたものだけExcel台帳へ反映
- 差し戻し、履歴保存
- Excel台帳が無い場合は自動作成

## 本番の流れ

```text
FAX受信
↓
複合機がPDF保存
↓
会社PCの受信フォルダへ入る
↓
アプリが確認待ち一覧へ追加
↓
社員が画面で確認
↓
OKしたものだけExcel台帳へ反映
```

社員はExcelを直接触らず、確認画面だけで運用できます。

## 画面で起動する

```bat
start_fax_excel_ai.bat
```

ブラウザで `http://127.0.0.1:8765` を開きます。

## FAXフォルダを自動監視する

```bat
start_fax_auto_watch.bat
```

標準では `incoming_fax` フォルダを見ます。PDFや画像が入ると、確認待ち一覧に追加します。

## 設定

`config.example.json` を `config.json` にコピーして、会社PCに合わせます。

```json
{
  "incoming_fax_dir": "C:\\FAX受信",
  "ledger_path": "C:\\Users\\Public\\Documents\\発注台帳.xlsx",
  "reviewer_default": "社員確認",
  "auto_process_to_review": true
}
```

## セットアップ

Python 3.11以上を想定しています。

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py 8765
```

## 写真OCR、スキャンPDF OCR

写真やスキャンPDFを読むにはTesseract OCRが必要です。Render用のDocker構成では自動で入ります。会社PCで本番運用する場合は、そのPCにもTesseract OCRを入れてください。

OCRは常に自動最適化します。まず速い `eng` で読み取り、英数字だけで十分読めた場合はそこで止めます。日本語が必要そうな場合だけ `jpn`、最後に `jpn+eng` を試します。画像の白黒化、拡大、コントラスト補正も標準で行います。

`config.json` で変更できます。

```json
{
  "ocr_min_chars": 8,
  "ocr_psm": "6",
  "ocr_timeout_seconds": 8,
  "ocr_preprocess": true,
  "ocr_pdf_dpi": 140,
  "ocr_target_long_side": 1100,
  "ocr_max_pages": 1
}
```

読み取り時間を短くするため、スキャンPDFは標準で先頭1ページをOCRします。複数ページすべてを読みたい場合は `ocr_max_pages` を増やしてください。

## 対応ファイル

```text
PDF
JPG / JPEG / PNG / BMP / TIFF / WEBP
DOCX
XLSX / XLSM
TXT / CSV / TSV / LOG / JSON / XML / MD / RTF
HTML
EML
ZIP
```

HEIC/HEIF、古いWordの`.doc`、古いExcelの`.xls`は環境によって読み取りが難しいため、JPG/PNG/PDF/DOCX/XLSXに変換して使う想定です。

## GitHub PagesとRenderの使い分け

- GitHub Pages: 画面デモ用
- Render: 写真/PDFアップロードのWebデモ用
- 会社PC: FAXフォルダ監視、自社Excel台帳反映の本番用

## GitHubに上げないもの

```text
outputs/
incoming_fax/
processed_fax/
pc_excel/
pending_reviews/
rejected_fax/
history.jsonl
config.json
__pycache__/
```
