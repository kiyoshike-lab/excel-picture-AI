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
const serverReadExtensions = [
  ".pdf",
  ".jpg",
  ".jpeg",
  ".png",
  ".bmp",
  ".tif",
  ".tiff",
  ".webp",
  ".heic",
  ".heif",
  ".docx",
  ".xlsx",
  ".xlsm",
  ".html",
  ".htm",
  ".eml",
  ".zip",
];
const browserCompressExtensions = [".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"];
const browserImageMaxSide = 1200;
const browserImageQuality = 0.78;
const uploadTimeoutMs = 25000;

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

let current = { header: {}, items: [], confidence: {}, raw_text: "", source_name: "manual" };
let currentReviewId = "";
let selectedSourceName = "manual";

const faxText = document.querySelector("#faxText");
const headerFields = document.querySelector("#headerFields");
const itemRows = document.querySelector("#itemRows");
const status = document.querySelector("#status");
const exportBtn = document.querySelector("#exportBtn");
const approveBtn = document.querySelector("#approveBtn");
const rejectBtn = document.querySelector("#rejectBtn");
const saveReviewBtn = document.querySelector("#saveReviewBtn");
const refreshReviewsBtn = document.querySelector("#refreshReviewsBtn");
const autoRunBtn = document.querySelector("#autoRunBtn");
const reviewList = document.querySelector("#reviewList");
const reviewerInput = document.querySelector("#reviewerInput");
const download = document.querySelector("#download");
const progressPanel = document.querySelector("#progressPanel");
const progressTitle = document.querySelector("#progressTitle");
const progressTime = document.querySelector("#progressTime");
const progressFill = document.querySelector("#progressFill");
const progressDetail = document.querySelector("#progressDetail");
let progressTimer = null;
let progressStartedAt = 0;

function setStatus(text, type = "neutral") {
  status.textContent = text;
  status.dataset.type = type;
}

function updateProgress(percent, title, detail) {
  if (!progressPanel) return;
  progressPanel.hidden = false;
  progressFill.style.width = `${Math.max(0, Math.min(100, percent))}%`;
  if (title) progressTitle.textContent = title;
  if (detail) progressDetail.textContent = detail;
}

function startProgress(title, detail) {
  progressStartedAt = Date.now();
  updateProgress(3, title, detail);
  clearInterval(progressTimer);
  progressTimer = setInterval(() => {
    const elapsed = Math.floor((Date.now() - progressStartedAt) / 1000);
    progressTime.textContent = `${elapsed}秒`;
    if (elapsed >= 20) {
      progressDetail.textContent = "20秒を超えています。ファイルが重いか、サーバーが混み合っています。";
    }
  }, 250);
}

function finishProgress(title, detail, percent = 100) {
  updateProgress(percent, title, detail);
  clearInterval(progressTimer);
  progressTimer = null;
  const elapsed = Math.floor((Date.now() - progressStartedAt) / 1000);
  progressTime.textContent = `${elapsed}秒`;
}

function hideProgressSoon() {
  window.setTimeout(() => {
    if (!progressPanel) return;
    progressFill.style.width = "0%";
    progressTitle.textContent = "待機中";
    progressTime.textContent = "0秒";
    progressDetail.textContent = "ファイルを選ぶと進み具合がここに出ます";
  }, 2500);
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

function enableReviewActions(enabled) {
  exportBtn.disabled = !enabled;
  approveBtn.disabled = !enabled;
  rejectBtn.disabled = !currentReviewId;
  saveReviewBtn.disabled = !enabled;
}

function renderCurrent() {
  renderHeader();
  renderItems();
  enableReviewActions(true);
}

async function extract() {
  const text = faxText.value.trim();
  if (!text) {
    setStatus("先に写真・PDFを選ぶか、文章を貼り付けてください", "warning");
    return;
  }

  startProgress("項目を整理中", "読み取った文字から日付・会社名・品名を探しています");
  setStatus("内容を読み取っています...", "busy");
  download.innerHTML = "";
  const response = await fetch("/api/extract", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ text }),
  });
  current = await response.json();
  current.source_name = selectedSourceName || "manual";
  currentReviewId = "";
  renderCurrent();
  finishProgress("整理完了", "社員確認に進めます");
  hideProgressSoon();
  setStatus("読み取りました。確認してからExcelへ反映できます", "success");
}

async function saveReview() {
  setStatus("確認待ちに保存しています...", "busy");
  const response = await fetch("/api/save-review", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(current),
  });
  const review = await response.json();
  currentReviewId = review.id;
  rejectBtn.disabled = false;
  await refreshReviews();
  setStatus("確認待ちに保存しました", "success");
}

async function approveCurrent() {
  setStatus("確認済みとしてExcel台帳へ反映しています...", "busy");
  const body = currentReviewId
    ? { id: currentReviewId, reviewer: reviewerInput.value, payload: current }
    : { ...current, reviewer: reviewerInput.value };

  const response = await fetch(currentReviewId ? "/api/approve-review" : "/api/append-ledger", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  const result = await response.json();
  if (!response.ok || result.error) {
    setStatus(result.error || "Excel台帳へ反映できませんでした", "warning");
    return;
  }
  download.innerHTML = `${result.rows_added}行を確認済みとしてExcel台帳へ反映しました<br>${result.path}`;
  currentReviewId = "";
  rejectBtn.disabled = true;
  await refreshReviews();
  setStatus("確認済みとしてExcel台帳へ反映しました", "success");
}

async function rejectCurrent() {
  if (!currentReviewId) return;
  const reason = window.prompt("差し戻し理由を入力してください", "内容確認が必要");
  if (reason === null) return;
  const response = await fetch("/api/reject-review", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ id: currentReviewId, reason }),
  });
  const result = await response.json();
  if (!response.ok || result.error) {
    setStatus(result.error || "差し戻しできませんでした", "warning");
    return;
  }
  currentReviewId = "";
  enableReviewActions(false);
  await refreshReviews();
  setStatus("差し戻しにしました", "success");
}

async function exportExcel() {
  setStatus("Excelファイルを別保存しています...", "busy");
  const response = await fetch("/api/export", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(current),
  });
  const result = await response.json();
  download.innerHTML = `<a href="${result.download_url}">${result.file}</a> を保存できます`;
  setStatus("Excelファイルができました", "success");
}

async function autoRun() {
  setStatus("受信FAXフォルダを確認待ちへ取り込んでいます...", "busy");
  download.innerHTML = "";
  const response = await fetch("/api/auto-run", { method: "POST" });
  const result = await response.json();
  const rows = result.results || [];
  await refreshReviews();
  if (!rows.length) {
    setStatus("新しいFAXファイルはありません", "neutral");
    download.textContent = `確認したフォルダ: ${result.folder}`;
    return;
  }
  download.innerHTML = rows
    .map((row) => `<div>${row.source}: ${row.status === "pending_review" ? "確認待ちに追加" : row.warning || "確認が必要"}</div>`)
    .join("");
  setStatus(`${rows.length}件を確認待ちへ取り込みました`, "success");
}

async function refreshReviews() {
  const response = await fetch("/api/reviews");
  const result = await response.json();
  const reviews = result.reviews || [];
  if (!reviews.length) {
    reviewList.innerHTML = `<div class="empty-state">確認待ちはありません</div>`;
    return;
  }
  reviewList.innerHTML = reviews
    .map(
      (review) => `
        <button class="review-item" data-id="${review.id}" type="button">
          <strong>${review.project_name || review.order_no || review.source_name || "確認待ち"}</strong>
          <span>${review.created_at || ""} / ${review.items_count || 0}明細</span>
        </button>
      `
    )
    .join("");
  reviewList.querySelectorAll(".review-item").forEach((button) => {
    button.addEventListener("click", () => openReview(button.dataset.id));
  });
}

async function openReview(id) {
  setStatus("確認待ちを開いています...", "busy");
  const response = await fetch(`/api/reviews/${id}`);
  const review = await response.json();
  if (!response.ok || review.error) {
    setStatus(review.error || "確認待ちを開けませんでした", "warning");
    return;
  }
  currentReviewId = review.id;
  current = review.payload;
  faxText.value = current.raw_text || "";
  renderCurrent();
  setStatus("内容を確認して、OKならExcelへ反映してください", "success");
}

function shouldUploadToServer(file) {
  const name = file.name.toLowerCase();
  return serverReadExtensions.some((extension) => name.endsWith(extension));
}

function shouldCompressBeforeUpload(file) {
  const name = file.name.toLowerCase();
  return browserCompressExtensions.some((extension) => name.endsWith(extension));
}

async function fetchWithTimeout(url, options, timeoutMs = uploadTimeoutMs) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetch(url, { ...options, signal: controller.signal });
  } finally {
    clearTimeout(timeout);
  }
}

function uploadReadFile(formData, onUploadProgress) {
  return new Promise((resolve, reject) => {
    const request = new XMLHttpRequest();
    const timeout = setTimeout(() => {
      request.abort();
      reject(new Error("timeout"));
    }, uploadTimeoutMs);

    request.open("POST", "/api/read-file");
    request.upload.onprogress = (event) => {
      if (event.lengthComputable && onUploadProgress) {
        onUploadProgress(Math.round((event.loaded / event.total) * 100));
      }
    };
    request.onload = () => {
      clearTimeout(timeout);
      let result = {};
      try {
        result = JSON.parse(request.responseText || "{}");
      } catch (_) {
        result = { error: "読み取り結果を確認できませんでした。" };
      }
      resolve({ ok: request.status >= 200 && request.status < 300, status: request.status, result });
    };
    request.onerror = () => {
      clearTimeout(timeout);
      reject(new Error("network"));
    };
    request.onabort = () => {
      clearTimeout(timeout);
      reject(new Error("aborted"));
    };
    request.send(formData);
  });
}

function loadImageForCompression(file) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file);
    const image = new Image();
    image.onload = () => {
      URL.revokeObjectURL(url);
      resolve(image);
    };
    image.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("image load failed"));
    };
    image.src = url;
  });
}

async function compressImageBeforeUpload(file) {
  if (!shouldCompressBeforeUpload(file)) return file;
  try {
    const image = await loadImageForCompression(file);
    const longSide = Math.max(image.naturalWidth, image.naturalHeight);
    if (!longSide || longSide <= browserImageMaxSide && file.size < 900000) return file;

    const scale = Math.min(1, browserImageMaxSide / longSide);
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(image.naturalWidth * scale));
    canvas.height = Math.max(1, Math.round(image.naturalHeight * scale));
    const context = canvas.getContext("2d", { alpha: false });
    context.fillStyle = "#fff";
    context.fillRect(0, 0, canvas.width, canvas.height);
    context.drawImage(image, 0, 0, canvas.width, canvas.height);

    const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/jpeg", browserImageQuality));
    if (!blob || blob.size >= file.size) return file;
    const baseName = file.name.replace(/\.[^.]+$/, "");
    return new File([blob], `${baseName}.jpg`, { type: "image/jpeg", lastModified: file.lastModified });
  } catch (_) {
    return file;
  }
}

document.querySelector("#sampleBtn").addEventListener("click", () => {
  faxText.value = sampleText;
  currentReviewId = "";
  selectedSourceName = "sample";
  setStatus("サンプルを入れました。次に「読み取る」を押してください", "success");
});

document.querySelector("#extractBtn").addEventListener("click", extract);
saveReviewBtn.addEventListener("click", saveReview);
approveBtn.addEventListener("click", approveCurrent);
rejectBtn.addEventListener("click", rejectCurrent);
exportBtn.addEventListener("click", exportExcel);
autoRunBtn.addEventListener("click", autoRun);
refreshReviewsBtn.addEventListener("click", refreshReviews);

document.querySelector("#fileInput").addEventListener("change", async (event) => {
  const files = Array.from(event.target.files || []);
  if (!files.length) return;

  if (files.some(shouldUploadToServer)) {
    startProgress("ファイル準備中", "画像を軽くして、送信の準備をしています");
    setStatus("ファイルを読み込んでいます...", "busy");
    const formData = new FormData();
    const uploadFiles = await Promise.all(files.map(compressImageBeforeUpload));
    uploadFiles.forEach((file) => formData.append("file", file));
    updateProgress(18, "アップロード中", "サーバーへ送っています");
    let uploadResult;
    try {
      uploadResult = await uploadReadFile(formData, (percent) => {
        updateProgress(20 + percent * 0.35, "アップロード中", `${percent}% 送信しました`);
      });
    } catch (error) {
      finishProgress("中断しました", "25秒で止めました。ファイルを小さくして試してください", 100);
      setStatus("25秒で止めました。ファイルが大きすぎるか、古いサイトが動いています。PDFは1ページ、画像は小さめで試してください。", "warning");
      return;
    }
    updateProgress(70, "文字認識中", "OCRで文字を読んでいます");
    const { ok, result } = uploadResult;
    if (!ok || result.error) {
      finishProgress("読み取り停止", result.error || "ファイルを読み込めませんでした", 100);
      setStatus(result.error || "ファイルを読み込めませんでした", "warning");
      return;
    }
    faxText.value = result.text || "";
    selectedSourceName = files.map((file) => file.name).join(", ");
    const versionText = result.version ? ` / ${result.version}` : "";
    const elapsedText = result.elapsed_seconds ? ` / ${result.elapsed_seconds}秒` : "";
    finishProgress("文字読み取り完了", `次に「読み取る」を押してください${elapsedText}${versionText}`, 100);
    hideProgressSoon();
    setStatus(result.warning || `読み込み完了${elapsedText}${versionText}`, result.warning ? "warning" : "success");
    return;
  }

  startProgress("テキスト読み込み中", "選んだファイルの文字を画面に入れています");
  faxText.value = (await Promise.all(files.map(async (file) => `--- ${file.name} ---\n${await file.text()}`))).join("\n\n");
  selectedSourceName = files.map((file) => file.name).join(", ");
  finishProgress("読み込み完了", "次に「読み取る」を押してください");
  hideProgressSoon();
  setStatus("読み込みました。次に「読み取る」を押してください", "success");
});

renderHeader();
renderItems();
refreshReviews();
