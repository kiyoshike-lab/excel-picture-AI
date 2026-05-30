const labels = {
  received_date: "受信日",
  document_type: "書類種別",
  sender_company: "送信元会社",
  recipient_company: "宛先会社",
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
const fileExtensions = [".pdf", ".jpg", ".jpeg", ".png", ".tif", ".tiff", ".webp"];

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

let current = {
  header: {},
  items: [],
  confidence: {},
  raw_text: "",
};

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

function renderHeader() {
  headerFields.innerHTML = "";
  for (const [key, label] of Object.entries(labels)) {
    const wrapper = document.createElement("div");
    wrapper.className = `field ${(current.confidence[key] || 0) < 0.7 ? "low" : ""}`;

    const fieldLabel = document.createElement("label");
    fieldLabel.textContent = label;

    const input = document.createElement("input");
    input.value = current.header[key] || "";
    input.dataset.key = key;
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
      input.inputMode = ["quantity", "unit_price", "amount"].includes(key) ? "decimal" : "text";
      input.addEventListener("input", () => {
        current.items[index][key] = input.value;
      });
      cell.append(input);
      row.append(cell);
    });
    itemRows.append(row);
  });
}

async function extract() {
  const text = faxText.value.trim();
  if (!text) {
    setStatus("原本データを入力してください", "warning");
    return;
  }

  setStatus("読み取り結果を整理しています...", "busy");
  download.innerHTML = "";
  const response = await fetch("/api/extract", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ text }),
  });
  current = await response.json();
  renderHeader();
  renderItems();
  exportBtn.disabled = false;
  setStatus("確認して、必要な箇所だけ修正できます", "success");
}

async function exportExcel() {
  setStatus("Excelを作成しています...", "busy");
  const response = await fetch("/api/export", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(current),
  });
  const result = await response.json();
  download.innerHTML = `<a href="${result.download_url}">${result.file}</a> を作成しました`;
  setStatus("Excel出力が完了しました", "success");
}

async function autoRun() {
  setStatus("受信FAXフォルダを確認しています...", "busy");
  download.innerHTML = "";
  const response = await fetch("/api/auto-run", { method: "POST" });
  const result = await response.json();
  const rows = result.results || [];
  if (!rows.length) {
    setStatus("受信FAXフォルダに新しいファイルはありません", "neutral");
    download.textContent = `監視対象: ${result.folder}`;
    return;
  }
  download.innerHTML = rows
    .map((row) => {
      if (row.status === "created") {
        return `<div>${row.source}: <a href="${row.download_url}">${row.file}</a></div>`;
      }
      return `<div>${row.source}: ${row.warning || "確認が必要です"}</div>`;
    })
    .join("");
  setStatus(`${rows.length}件を処理しました`, "success");
}

function shouldUploadToServer(file) {
  const name = file.name.toLowerCase();
  return fileExtensions.some((extension) => name.endsWith(extension));
}

document.querySelector("#sampleBtn").addEventListener("click", () => {
  faxText.value = sampleText;
  setStatus("サンプルを読み込みました", "success");
});

document.querySelector("#extractBtn").addEventListener("click", extract);
exportBtn.addEventListener("click", exportExcel);
document.querySelector("#autoRunBtn").addEventListener("click", autoRun);

document.querySelector("#fileInput").addEventListener("change", async (event) => {
  const [file] = event.target.files;
  if (!file) return;

  if (shouldUploadToServer(file)) {
    setStatus("ファイルを読み取っています...", "busy");
    const formData = new FormData();
    formData.append("file", file);
    const response = await fetch("/api/read-file", { method: "POST", body: formData });
    const result = await response.json();
    if (!response.ok || result.error) {
      setStatus(result.error || "ファイルを読み取れませんでした", "warning");
      return;
    }
    faxText.value = result.text || "";
    setStatus(result.warning || `${file.name} を読み込みました`, result.warning ? "warning" : "success");
    return;
  }

  faxText.value = await file.text();
  setStatus(`${file.name} を読み込みました`, "success");
});

renderHeader();
renderItems();
