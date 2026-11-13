/* PDF 转可编辑 PPTX —— 前端交互 */

const state = {
  files: [],
  maxSizeMb: 100,
};

const el = {
  dropzone: document.getElementById("dropzone"),
  fileInput: document.getElementById("file-input"),
  fileList: document.getElementById("file-list"),
  convertBtn: document.getElementById("convert-btn"),
  status: document.getElementById("status"),
  results: document.getElementById("results"),
  errors: document.getElementById("errors"),
  optOcr: document.getElementById("opt-ocr"),
  optArt: document.getElementById("opt-art"),
  optTables: document.getElementById("opt-tables"),
  optLayering: document.getElementById("opt-layering"),
  maxSize: document.getElementById("max-size"),
  dialog: document.getElementById("detail-dialog"),
  dialogTitle: document.getElementById("detail-title"),
  dialogBody: document.getElementById("detail-body"),
};

function formatSize(bytes) {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
  return (bytes / 1024 / 1024).toFixed(2) + " MB";
}

function setStatus(message, kind = "") {
  el.status.className = "status" + (kind ? " " + kind : "");
  el.status.innerHTML = message;
}

function addFiles(fileList) {
  for (const file of fileList) {
    if (!file.name.toLowerCase().endsWith(".pdf")) {
      showErrors([`${file.name}：仅支持 PDF 格式`]);
      continue;
    }
    if (file.size > state.maxSizeMb * 1024 * 1024) {
      showErrors([`${file.name}：超过 ${state.maxSizeMb} MB 上限`]);
      continue;
    }
    const duplicated = state.files.some((item) => item.name === file.name && item.size === file.size);
    if (!duplicated) state.files.push(file);
  }
  renderFiles();
}

function renderFiles() {
  el.fileList.innerHTML = "";
  state.files.forEach((file, index) => {
    const li = document.createElement("li");
    li.className = "file-item";
    li.innerHTML = `
      <span class="name" title="${escapeHtml(file.name)}">${escapeHtml(file.name)}</span>
      <span class="size">${formatSize(file.size)}</span>
      <button type="button" title="移除" aria-label="移除">✕</button>`;
    li.querySelector("button").addEventListener("click", () => {
      state.files.splice(index, 1);
      renderFiles();
    });
    el.fileList.appendChild(li);
  });
  el.convertBtn.disabled = state.files.length === 0;
}

function escapeHtml(text) {
  return String(text).replace(/[&<>"']/g, (ch) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[ch]));
}

const IMAGE_KIND_LABELS = {
  complex_vector_layer: "复杂矢量路径（独立图片对象）",
  background: "重建背景层",
  photo: "独立照片对象",
  illustration: "独立插画对象",
  logo: "独立标识对象",
  graphic: "独立图形对象",
  pdf: "PDF 内嵌位图",
};

const SHAPE_KIND_LABELS = {
  rect: "矩形",
  line: "直线（连接符）",
  ellipse: "椭圆",
  freeform: "自由曲线（斜切多边形）",
  complex: "复杂矢量",
};

function showErrors(messages) {
  el.errors.innerHTML = messages.map((msg) => `<p>${escapeHtml(msg)}</p>`).join("");
}

/* ---------- 转换 ---------- */
async function convert() {
  if (state.files.length === 0) return;
  el.convertBtn.disabled = true;
  el.results.innerHTML = "";
  el.errors.innerHTML = "";
  setStatus('<span class="spinner"></span>正在转换，请稍候…');

  const form = new FormData();
  state.files.forEach((file) => form.append("files", file, file.name));
  form.append("enable_ocr", el.optOcr.checked ? "1" : "0");
  form.append("keep_art_layer", el.optArt.checked ? "1" : "0");
  form.append("detect_tables", el.optTables.checked ? "1" : "0");
  form.append("enable_layering", el.optLayering.checked ? "1" : "0");

  try {
    const response = await fetch("/api/convert", { method: "POST", body: form });
    const payload = await response.json();
    if (!response.ok) {
      const detail = payload.detail || "转换失败";
      setStatus(escapeHtml(detail), "error");
      showErrors(Array.isArray(detail) ? detail.map((d) => d.msg) : [detail]);
      return;
    }
    renderResults(payload);
    setStatus(
      `转换完成，共 ${payload.files.length} 个文件，用时 ${payload.elapsed_seconds} 秒`,
      "ok"
    );
    if (payload.errors && payload.errors.length) showErrors(payload.errors);
  } catch (error) {
    setStatus("请求失败：" + escapeHtml(error.message), "error");
  } finally {
    el.convertBtn.disabled = state.files.length === 0;
  }
}

function renderResults(payload) {
  el.results.innerHTML = "";
  payload.files.forEach((item) => {
    const stats = item.stats;
    const card = document.createElement("div");
    card.className = "result-card";
    const percent = Math.round((stats.editable_ratio || 0) * 100);
    card.innerHTML = `
      <h4>${escapeHtml(item.output_name)}</h4>
      <div class="chips">
        <span class="chip">${stats.page_count} 页</span>
        <span class="chip">文本 ${stats.text_lines} 行 / ${stats.text_chars} 字</span>
        <span class="chip">图片 ${stats.images}</span>
        <span class="chip">图形 ${stats.shapes}</span>
        <span class="chip">表格 ${stats.tables}</span>
        ${stats.background_layers ? `<span class="chip">背景层 ${stats.background_layers}</span>` : ""}
        ${stats.decomposed_pages ? `<span class="chip">已分解 ${stats.decomposed_pages}/${stats.page_count} 页</span>` : ""}
        ${stats.full_page_image_pages ? `<span class="chip warn">整页图 ${stats.full_page_image_pages} 页</span>` : ""}
        ${stats.art_layers ? `<span class="chip gray">复杂图形 ${stats.art_layers} 层</span>` : ""}
      </div>
      <div class="ratio">
        <div class="bar"><i style="width:${percent}%"></i></div>
        <span>可编辑对象占比 ${percent}%</span>
      </div>
      <div class="actions">
        <a class="btn primary-link" href="${item.download_url}">下载 PPTX</a>
        <button class="btn" type="button" data-inspect="${item.inspect_url}" data-name="${escapeHtml(item.source_name)}">查看元素明细</button>
      </div>`;
    card.querySelector("[data-inspect]").addEventListener("click", (event) =>
      showDetail(event.currentTarget.dataset.inspect, event.currentTarget.dataset.name)
    );
    el.results.appendChild(card);
  });
}

/* ---------- 元素明细 ---------- */
async function showDetail(url, name) {
  el.dialogTitle.textContent = name + " —— 元素明细";
  el.dialogBody.innerHTML = '<p class="placeholder">加载中…</p>';
  el.dialog.showModal();
  try {
    const response = await fetch(url);
    if (!response.ok) throw new Error("报告已过期");
    const report = await response.json();
    el.dialogBody.innerHTML = renderReport(report);
  } catch (error) {
    el.dialogBody.innerHTML = `<p class="placeholder">${escapeHtml(error.message)}</p>`;
  }
}

function renderReport(report) {
  const summary = report.summary || {};
  let html = `<div class="chips">
      <span class="chip">页数 ${report.page_count}</span>
      <span class="chip">文本 ${summary.text_lines} 行 / ${summary.text_chars} 字</span>
      <span class="chip">图片 ${summary.images}</span>
      <span class="chip">图形 ${summary.shapes}</span>
      <span class="chip">表格 ${summary.tables}</span>
      ${summary.background_layers ? `<span class="chip">背景层 ${summary.background_layers}</span>` : ""}
      ${summary.decomposed_pages ? `<span class="chip">已分解 ${summary.decomposed_pages}/${report.page_count} 页</span>` : ""}
      ${summary.full_page_image_pages ? `<span class="chip warn">整页图 ${summary.full_page_image_pages} 页</span>` : ""}
      <span class="chip gray">复杂图形 ${summary.art_layers} 层</span>
    </div>`;

  (report.pages || []).forEach((page) => {
    html += `<h4>第 ${page.index} 页（${page.width} × ${page.height} pt）</h4>`;
    html += `<table><thead><tr>
        <th>类型</th><th>位置 (x0,y0,x1,y1)</th><th>详情</th>
      </tr></thead><tbody>`;
    (page.elements || []).forEach((item) => {
      let detail = "";
      if (item.type === "text") {
        detail = `文本「${escapeHtml(item.text)}」 · ${escapeHtml(item.font)} ${item.size}pt${item.bold ? " · 加粗" : ""}`;
      } else if (item.type === "image") {
        detail = IMAGE_KIND_LABELS[item.kind] || `位图 ${(item.pixel || []).join("×")} · ${(item.format || "").toUpperCase()}`;
      } else if (item.type === "shape") {
        detail = `${SHAPE_KIND_LABELS[item.kind] || item.kind} · 填充 ${item.fill || "无"} · 描边 ${item.stroke || "无"}`;
      } else if (item.type === "table") {
        detail = `表格 ${item.rows} 行 × ${item.cols} 列`;
      }
      html += `<tr><td>${item.type}</td><td>${(item.bbox || []).join(", ")}</td><td>${detail}</td></tr>`;
    });
    html += "</tbody></table>";
  });
  return html;
}

/* ---------- 事件绑定 ---------- */
el.dropzone.addEventListener("click", () => el.fileInput.click());
el.dropzone.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    el.fileInput.click();
  }
});
el.fileInput.addEventListener("change", () => {
  addFiles(el.fileInput.files);
  el.fileInput.value = "";
});

["dragenter", "dragover"].forEach((name) =>
  el.dropzone.addEventListener(name, (event) => {
    event.preventDefault();
    el.dropzone.classList.add("dragover");
  })
);
["dragleave", "drop"].forEach((name) =>
  el.dropzone.addEventListener(name, (event) => {
    event.preventDefault();
    el.dropzone.classList.remove("dragover");
  })
);
el.dropzone.addEventListener("drop", (event) => addFiles(event.dataTransfer.files));

el.convertBtn.addEventListener("click", convert);

fetch("/api/health")
  .then((response) => response.json())
  .then((data) => {
    if (data && data.version) {
      el.maxSize.textContent = data.max_upload_mb || state.maxSizeMb;
      state.maxSizeMb = data.max_upload_mb || state.maxSizeMb;
    }
  })
  .catch(() => {});