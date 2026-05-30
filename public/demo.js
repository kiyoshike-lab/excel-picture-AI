const labels = {
  received_date: "日付",
  document_type: "書類の種類",
  sender_company: "送った会社",
  recipient_company: "あて先",
  fax_no: "FAX番号",
  order_no: "注文番号",
  project_name: "工事名",
  site_name: "現場名",
  delivery_place: "納入場所",
  delivery_date: "納期",
  desired_time: "希望時間",
  person_in_charge: "担当者",
  notes: "備考",
};

const itemKeys = ["item_code", "item_name", "quantity", "unit", "unit_price", "amount"];

const sampleText = `資材発注書
日付: 2026/05/29
送信元: 株式会社青葉建設
宛先: 株式会社東都資材 御中
FAX: 03-1234-5678
注文番号: AOB-20260529-018
工事名: 日本橋第3ビル改修工事
現場名: 日本橋三丁目作業所
納入場所: 東京都中央区日本橋3-1-1 北側搬入口
担当者: 佐藤
納期: 2026/06/05
希望時間: 8:00-10:00

品番 商品名 数量 単位 単価 金額
RC-2525 普通ポルトランドセメント 20 袋 780 15600
RB-0013 異形鉄筋 D13 80 本 510 40800
CN-0045 コンクリート釘 10 箱 920 9200

備考: 搬入前に現場担当へ電話してください。`;

let current = { header: {}, items: [], confidence: {}, raw_text: "" };

const faxText = document.querySelector("#faxText");
const headerFields = document.querySelector("#headerFields");
const itemRows = document.querySelector("#itemRows");
const status = document.querySelector("#status");
const exportBtn = document.querySelector("#exportBtn");
const download = document.querySelector("#download");

function setStatus(text, type = "neutral") {
  status.textContent = text;
  status.dataset.type = type;
}

function first(patterns, text) {
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) return match[1].trim();
  }
  return "";
}

function normalizeDate(value) {
  const match = (value || "").match(/(\d{4})[./年-](\d{1,2})[./月-](\d{1,2})/);
  return match ? `${match[1]}-${String(Number(match[2])).padStart(2, "0")}-${String(Number(match[3])).padStart(2, "0")}` : value;
}

function toNumber(value) {
  const parsed = Number(String(value || "").replace(/[^0-9.-]/g, ""));
  return Number.isFinite(parsed) ? parsed : "";
}

function extractLocal(text) {
  const header = {
    received_date: normalizeDate(first([/(?:受信日|日付|発注日)[:： ]*([0-9./年月日-]+)/], text)),
    document_type: first([/(見積依頼書|発注書|注文書|納品依頼書|資材依頼書|作業依頼書)/], text),
    sender_company: first([/(?:送信元|差出人|発注元)[:： ]*(.+)/], text),
    recipient_company: first([/(?:宛先|納入先)[:： ]*(.+)/], text),
    fax_no: first([/(?:FAX|Fax)[:： ]*([0-9\-() ]{8,})/], text),
    order_no: first([/(?:注文番号|発注番号|注文No\.?)[:： ]*([A-Z0-9-]+)/i], text),
    project_name: first([/(?:工事名|案件名|件名)[:： ]*(.+)/], text),
    site_name: first([/(?:現場名|現場)[:： ]*(.+)/], text),
    delivery_place: first([/(?:納入場所|搬入場所|納品場所|現場住所|住所)[:： ]*(.+)/], text),
    delivery_date: normalizeDate(first([/(?:納期|希望納期|納入日|搬入日|納品日)[:： ]*([0-9./年月日-]+)/], text)),
    desired_time: first([/(?:希望時間|搬入時間|納品時間|時間指定)[:： ]*(.+)/], text),
    person_in_charge: first([/(?:担当者|担当)[:： ]*(.+)/], text),
    notes: first([/(?:備考|摘要|連絡事項)[:： ]*(.+)/], text),
  };

  const itemPattern = /^([A-Z0-9][A-Z0-9-]{2,})\s+(.+?)\s+(\d+(?:\.\d+)?)\s*(個|枚|箱|本|台|kg|t|式|袋|缶|束|セット|m|ｍ|m2|m3|㎡|㎥)?\s+([\d,]+)?\s+([\d,]+)?$/i;
  const items = text
    .split("\n")
    .map((line) => line.trim())
    .map((line) => line.match(itemPattern))
    .filter(Boolean)
    .map((match) => {
      const quantity = toNumber(match[3]);
      const unitPrice = toNumber(match[5]);
      const amount = toNumber(match[6]) || (quantity && unitPrice ? quantity * unitPrice : "");
      return {
        item_code: match[1],
        item_name: match[2],
        quantity,
        unit: match[4] || "個",
        unit_price: unitPrice,
        amount,
      };
    });

  return {
    header,
    items: items.length ? items : [{ item_code: "", item_name: "", quantity: "", unit: "", unit_price: "", amount: "" }],
    confidence: Object.fromEntries(Object.entries(header).map(([key, value]) => [key, value ? 0.92 : 0.35])),
    raw_text: text,
  };
}

function renderHeader() {
  headerFields.innerHTML = "";
  for (const [key, label] of Object.entries(labels)) {
    const wrapper = document.createElement("div");
    wrapper.className = `field ${(current.confidence[key] || 0) < 0.7 ? "low" : ""}`;
    const fieldLabel = document.createElement("label");
    fieldLabel.textContent = label;
    const input = document.createElement("input");
    input.value = current.header[key] || "";
    input.addEventListener("input", () => {
      current.header[key] = input.value;
    });
    wrapper.append(fieldLabel, input);
    headerFields.append(wrapper);
  }
}

function renderItems() {
  itemRows.innerHTML = "";
  current.items.forEach((item, index) => {
    const row = document.createElement("tr");
    itemKeys.forEach((key) => {
      const cell = document.createElement("td");
      const input = document.createElement("input");
      input.value = item[key] ?? "";
      input.addEventListener("input", () => {
        current.items[index][key] = input.value;
      });
      cell.append(input);
      row.append(cell);
    });
    itemRows.append(row);
  });
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[char]));
}

document.querySelector("#sampleBtn").addEventListener("click", () => {
  faxText.value = sampleText;
  setStatus("サンプルを入れました。次に「読み取る」を押してください", "success");
});

document.querySelector("#extractBtn").addEventListener("click", () => {
  if (!faxText.value.trim()) {
    setStatus("先にサンプルを入れるか、文章を貼り付けてください", "warning");
    return;
  }
  current = extractLocal(faxText.value.trim());
  renderHeader();
  renderItems();
  exportBtn.disabled = false;
  setStatus("読み取りました。右側を確認してください", "success");
});

exportBtn.addEventListener("click", () => {
  const headerRows = Object.entries(labels)
    .map(([key, label]) => `<tr><th>${label}</th><td>${escapeHtml(current.header[key])}</td></tr>`)
    .join("");
  const itemRowsHtml = current.items
    .map((item) => `<tr>${itemKeys.map((key) => `<td>${escapeHtml(item[key])}</td>`).join("")}</tr>`)
    .join("");
  const html = `<html><head><meta charset="utf-8"></head><body><table border="1"><tr><th colspan="6">建設FAX入力結果</th></tr>${headerRows}<tr></tr><tr><th>品番</th><th>商品名</th><th>数量</th><th>単位</th><th>単価</th><th>金額</th></tr>${itemRowsHtml}</table></body></html>`;
  const url = URL.createObjectURL(new Blob([html], { type: "application/vnd.ms-excel;charset=utf-8" }));
  const link = document.createElement("a");
  link.href = url;
  link.download = "fax_excel_demo.xls";
  link.click();
  URL.revokeObjectURL(url);
  setStatus("Excelで開けるファイルを保存しました", "success");
});

document.querySelector("#autoRunBtn").addEventListener("click", () => {
  setStatus("受信FAXの自動処理は本番アプリで使えます", "neutral");
});

document.querySelector("#fileInput").addEventListener("change", async (event) => {
  const [file] = event.target.files;
  if (!file) return;
  faxText.value = await file.text();
  setStatus("読み込みました。次に「読み取る」を押してください", "success");
});

renderHeader();
renderItems();
