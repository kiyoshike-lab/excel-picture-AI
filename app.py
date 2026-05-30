from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse
import cgi
import io
import json
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from pypdf import PdfReader


ROOT = Path(__file__).resolve().parent
PUBLIC = ROOT / "public"
OUTPUTS = ROOT / "outputs"
INCOMING = ROOT / "incoming_fax"
PROCESSED = ROOT / "processed_fax"


FIELD_LABELS = {
    "received_date": "受信日",
    "document_type": "書類種別",
    "sender_company": "送信元会社",
    "recipient_company": "宛先会社",
    "fax_no": "FAX番号",
    "order_no": "注文番号",
    "project_name": "工事名",
    "site_name": "現場名",
    "delivery_place": "納入場所",
    "delivery_date": "納期",
    "desired_time": "希望時間",
    "person_in_charge": "担当者",
    "notes": "備考",
}

ITEM_LABELS = {
    "item_code": "品番",
    "item_name": "商品名",
    "quantity": "数量",
    "unit": "単位",
    "unit_price": "単価",
    "amount": "金額",
}

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"}
TEXT_EXTENSIONS = {".txt", ".csv", ".log"}
PDF_EXTENSIONS = {".pdf"}


def normalize_text(text):
    replacements = {
        "\r\n": "\n",
        "\r": "\n",
        "　": " ",
        "㈱": "株式会社",
        "（株）": "株式会社",
    }
    for before, after in replacements.items():
        text = text.replace(before, after)
    return re.sub(r"[ \t]+", " ", text).strip()


def first_match(patterns, text, default=""):
    for pattern in patterns:
        match = re.search(pattern, text, re.MULTILINE | re.IGNORECASE)
        if match:
            return match.group(1).strip(" :：\t")
    return default


def normalize_date(value):
    if not value:
        return ""
    value = value.strip()
    match = re.search(r"(\d{4})[./年-](\d{1,2})[./月-](\d{1,2})", value)
    if match:
        y, m, d = match.groups()
        return f"{int(y):04d}-{int(m):02d}-{int(d):02d}"
    match = re.search(r"(\d{1,2})[./月-](\d{1,2})", value)
    if match:
        m, d = match.groups()
        return f"{datetime.now().year:04d}-{int(m):02d}-{int(d):02d}"
    return value


def number_or_blank(value):
    if value is None:
        return ""
    cleaned = re.sub(r"[^0-9.-]", "", str(value))
    if cleaned in {"", "-", ".", "-."}:
        return ""
    try:
        number = float(cleaned)
        return int(number) if number.is_integer() else number
    except ValueError:
        return ""


def parse_items(text):
    items = []
    line_patterns = [
        re.compile(
            r"(?P<code>[A-Z0-9][A-Z0-9\-]{2,})\s+"
            r"(?P<name>.+?)\s+"
            r"(?P<qty>\d+(?:\.\d+)?)\s*(?P<unit>個|枚|箱|本|台|kg|t|式|袋|缶|束|セット|m|ｍ|m2|m3|㎡|㎥)?\s+"
            r"(?P<price>[\d,]+)?\s+"
            r"(?P<amount>[\d,]+)?$",
            re.IGNORECASE,
        ),
        re.compile(
            r"品番[:： ]?(?P<code>[A-Z0-9\-]+).*?"
            r"商品名[:： ]?(?P<name>.+?)\s+"
            r"数量[:： ]?(?P<qty>\d+(?:\.\d+)?).*?"
            r"(?:単価[:： ]?(?P<price>[\d,]+))?.*?"
            r"(?:金額[:： ]?(?P<amount>[\d,]+))?",
            re.IGNORECASE,
        ),
    ]

    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if any(word in line for word in ["品番", "商品名", "数量", "単価", "金額"]) and len(line) < 18:
            continue
        for pattern in line_patterns:
            match = pattern.search(line)
            if not match:
                continue
            data = match.groupdict()
            qty = number_or_blank(data.get("qty"))
            price = number_or_blank(data.get("price"))
            amount = number_or_blank(data.get("amount"))
            if amount == "" and qty != "" and price != "":
                amount = qty * price
            items.append(
                {
                    "item_code": data.get("code", "").strip(),
                    "item_name": data.get("name", "").strip(" -:："),
                    "quantity": qty,
                    "unit": (data.get("unit") or "個").strip(),
                    "unit_price": price,
                    "amount": amount,
                }
            )
            break

    if not items:
        items.append({"item_code": "", "item_name": "", "quantity": "", "unit": "", "unit_price": "", "amount": ""})
    return items


def extract_fax(text):
    text = normalize_text(text)
    header = {
        "received_date": normalize_date(first_match([r"(?:受信日|日付|発注日)[:： ]*([0-9./年月日-]+)"], text)),
        "document_type": first_match(
            [r"(見積依頼書|発注書|注文書|納品依頼書|資材依頼書|作業依頼書)", r"(?:書類種別|帳票種別)[:： ]*(.+)"],
            text,
        ),
        "sender_company": first_match(
            [
                r"(?:送信元|差出人|From|発注元)[:： ]*(.+)",
                r"(.+?(?:株式会社|有限会社|合同会社|商事|工業|産業|物流|建設|工務店|土木|設備))\s*(?:御中|様)?$",
            ],
            text,
        ),
        "recipient_company": first_match([r"(?:宛先|To|納入先)[:： ]*(.+)"], text),
        "fax_no": first_match([r"(?:FAX|Fax|ファックス)[:： ]*([0-9\-\(\) ]{8,})"], text),
        "order_no": first_match([r"(?:注文番号|発注番号|注文No\.?|Order No\.?)[:： ]*([A-Z0-9\-]+)"], text),
        "project_name": first_match([r"(?:工事名|案件名|件名)[:： ]*(.+)"], text),
        "site_name": first_match([r"(?:現場名|現場)[:： ]*(.+)"], text),
        "delivery_place": first_match([r"(?:納入場所|搬入場所|納品場所|現場住所|住所)[:： ]*(.+)"], text),
        "delivery_date": normalize_date(first_match([r"(?:納期|希望納期|納入日|搬入日|納品日)[:： ]*([0-9./年月日-]+)"], text)),
        "desired_time": first_match([r"(?:希望時間|搬入時間|納品時間|時間指定)[:： ]*(.+)"], text),
        "person_in_charge": first_match([r"(?:担当者|担当)[:： ]*(.+)"], text),
        "notes": first_match([r"(?:備考|摘要|連絡事項)[:： ]*(.+)"], text),
    }
    confidence = {key: 0.92 if value else 0.35 for key, value in header.items()}
    return {"header": header, "items": parse_items(text), "confidence": confidence, "raw_text": text}


def read_pdf_text(file_bytes):
    reader = PdfReader(io.BytesIO(file_bytes))
    pages = []
    for index, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""
        if text.strip():
            pages.append(f"--- PDF {index}ページ ---\n{text.strip()}")
    return "\n\n".join(pages).strip()


def read_image_text(file_bytes, suffix):
    tesseract = shutil.which("tesseract")
    if not tesseract:
        return "", "写真・画像OCRにはOCRエンジンの追加が必要です。TesseractまたはクラウドOCRを入れると、このまま自動入力できます。"

    with tempfile.TemporaryDirectory() as temp_dir:
        temp = Path(temp_dir)
        image_path = temp / f"input{suffix}"
        output_base = temp / "ocr"
        image_path.write_bytes(file_bytes)
        command = [tesseract, str(image_path), str(output_base), "-l", "jpn+eng", "--psm", "6"]
        result = subprocess.run(command, capture_output=True, text=True, timeout=60)
        if result.returncode != 0:
            return "", f"OCRに失敗しました: {result.stderr.strip() or result.stdout.strip()}"
        text_path = output_base.with_suffix(".txt")
        text = text_path.read_text(encoding="utf-8", errors="ignore") if text_path.exists() else ""
        return text.strip(), "" if text.strip() else "画像から文字を読み取れませんでした。写真の明るさや傾きを確認してください。"


def read_file_text(filename, file_bytes):
    suffix = Path(filename).suffix.lower()
    if suffix in PDF_EXTENSIONS:
        text = read_pdf_text(file_bytes)
        warning = "" if text else "PDF内に文字情報が見つかりませんでした。スキャン画像PDFの場合はOCR追加が必要です。"
        return text, warning
    if suffix in IMAGE_EXTENSIONS:
        return read_image_text(file_bytes, suffix)
    if suffix in TEXT_EXTENSIONS:
        return file_bytes.decode("utf-8-sig", errors="ignore").strip(), ""
    return "", "対応ファイルはPDF、画像、テキストです。"


def create_workbook(payload):
    OUTPUTS.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_source = re.sub(r"[^A-Za-z0-9_-]+", "_", payload.get("source_name", "")).strip("_")
    name = f"fax_excel_{timestamp}{'_' + safe_source[:32] if safe_source else ''}.xlsx"
    output_path = OUTPUTS / name

    wb = Workbook()
    ws = wb.active
    ws.title = "FAX入力"
    raw = wb.create_sheet("原文")

    title_fill = PatternFill("solid", fgColor="1F4E79")
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    thin = Side(style="thin", color="B7C9D6")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws["A1"] = "建設FAX自動入力 結果"
    ws["A1"].font = Font(size=16, bold=True, color="FFFFFF")
    ws["A1"].fill = title_fill
    ws.merge_cells("A1:H1")

    row = 3
    header = payload.get("header", {})
    for key, label in FIELD_LABELS.items():
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=2, value=header.get(key, ""))
        ws.cell(row=row, column=1).fill = header_fill
        row += 1

    row += 1
    start_row = row
    for col, label in enumerate(ITEM_LABELS.values(), start=1):
        cell = ws.cell(row=row, column=col, value=label)
        cell.fill = header_fill
        cell.font = Font(bold=True)
    row += 1

    for item in payload.get("items", []):
        for col, key in enumerate(ITEM_LABELS.keys(), start=1):
            ws.cell(row=row, column=col, value=item.get(key, ""))
        row += 1

    ws.cell(row=row, column=5, value="合計")
    ws.cell(row=row, column=6, value=f"=SUM(F{start_row + 1}:F{row - 1})")
    ws.cell(row=row, column=5).font = Font(bold=True)
    ws.cell(row=row, column=6).font = Font(bold=True)

    for sheet in [ws, raw]:
        for cells in sheet.iter_rows():
            for cell in cells:
                cell.border = border
                cell.alignment = Alignment(vertical="center", wrap_text=True)

    for index, width in enumerate([18, 38, 16, 12, 14, 14, 14, 18], start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    ws.freeze_panes = f"A{start_row + 1}"
    ws.auto_filter.ref = f"A{start_row}:F{max(start_row + 1, row - 1)}"

    raw["A1"] = "FAX原文"
    raw["A2"] = payload.get("raw_text", "")
    raw["A1"].font = Font(bold=True)
    raw.column_dimensions["A"].width = 120
    raw.row_dimensions[2].height = 220
    raw["A2"].alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(output_path)
    return output_path


def process_incoming_once():
    INCOMING.mkdir(exist_ok=True)
    PROCESSED.mkdir(exist_ok=True)
    results = []
    for path in sorted(p for p in INCOMING.iterdir() if p.is_file()):
        text, warning = read_file_text(path.name, path.read_bytes())
        if not text:
            results.append({"source": path.name, "status": "needs_ocr", "warning": warning})
            continue
        data = extract_fax(text)
        data["source_name"] = path.stem
        output = create_workbook(data)
        target = PROCESSED / path.name
        if target.exists():
            target = PROCESSED / f"{path.stem}_{datetime.now().strftime('%H%M%S')}{path.suffix}"
        path.replace(target)
        results.append({"source": path.name, "status": "created", "file": output.name, "download_url": f"/download/{output.name}", "warning": warning})
    return results


class Handler(BaseHTTPRequestHandler):
    def send_json(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path.startswith("/download/"):
            self.send_download(Path(parsed.path).name)
            return
        path = "/index.html" if parsed.path == "/" else parsed.path
        file_path = (PUBLIC / path.lstrip("/")).resolve()
        if not str(file_path).startswith(str(PUBLIC.resolve())) or not file_path.exists():
            self.send_error(404)
            return
        mime = "text/html; charset=utf-8"
        if file_path.suffix == ".css":
            mime = "text/css; charset=utf-8"
        elif file_path.suffix == ".js":
            mime = "application/javascript; charset=utf-8"
        body = file_path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", mime)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def multipart_file(self):
        ctype, _ = cgi.parse_header(self.headers.get("Content-Type", ""))
        if ctype != "multipart/form-data":
            return None
        form = cgi.FieldStorage(fp=self.rfile, headers=self.headers, environ={"REQUEST_METHOD": "POST"})
        return form["file"] if "file" in form else None

    def do_POST(self):
        if self.path in {"/api/read-file", "/api/read-pdf"}:
            file_item = self.multipart_file()
            if file_item is None or not getattr(file_item, "file", None):
                self.send_json({"error": "ファイルを読み取れませんでした。"}, 400)
                return
            text, warning = read_file_text(file_item.filename or "upload", file_item.file.read())
            self.send_json({"text": text, "warning": warning})
            return

        if self.path == "/api/auto-run":
            self.send_json({"folder": str(INCOMING), "results": process_incoming_once()})
            return

        length = int(self.headers.get("Content-Length", "0"))
        try:
            data = json.loads(self.rfile.read(length).decode("utf-8") or "{}")
        except json.JSONDecodeError:
            self.send_json({"error": "JSONを読み取れませんでした。"}, 400)
            return

        if self.path == "/api/extract":
            self.send_json(extract_fax(data.get("text", "")))
            return

        if self.path == "/api/export":
            output_path = create_workbook(data)
            self.send_json({"file": output_path.name, "download_url": f"/download/{output_path.name}"})
            return

        self.send_error(404)

    def send_download(self, filename):
        file_path = (OUTPUTS / filename).resolve()
        if not str(file_path).startswith(str(OUTPUTS.resolve())) or not file_path.exists():
            self.send_error(404)
            return
        body = file_path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, format, *args):
        return


def main():
    INCOMING.mkdir(exist_ok=True)
    OUTPUTS.mkdir(exist_ok=True)
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8765
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    print(f"http://127.0.0.1:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
