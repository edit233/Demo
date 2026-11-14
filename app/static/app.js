/* PDF 转可编辑 PPTX —— 前端交互 */

const state = {
  files: [],
  maxSizeMb: 100,
  maxFiles: 10,
};

// 单份 10 页 PDF 的实测转换耗时约 50 秒，用于给用户一个批量耗时的预期
const SECONDS_PER_FILE = 50;

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
    if (duplicated) continue;
    if (state.files.length >= state.maxFiles) {
      // 批量是逐个串行累加的，文件数直接决定请求耗时，超出上限就提前拦下
      showErrors([`单次最多转换 ${state.maxFiles} 个文件，请分批提交`]);
      continue;
    }
    state.files.push(file);
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

function showErrors(messages) {
  el.errors.innerHTML = messages.map((msg) => `<p>${escapeHtml(msg)}</p>`).join("");
}

/* ---------- 响应解析 ---------- */

// 常见网关状态码的排查提示，便于直接定位问题
const HTTP_HINTS = {
  404: "反向代理未把 /api/ 转发到后端",
  413: "反向代理的上传体积上限过小（nginx client_max_body_size）",
  502: "后端容器无响应或已被重启（查看 docker compose logs）",
  504: "反向代理读取超时（nginx proxy_read_timeout），转换耗时可能超过代理上限",
};

async function readJsonResponse(response) {
  // 先按文本读取：代理返回 HTML 错误页时 JSON.parse 会直接抛异常并吞掉状态码
  const raw = await response.text();
  try {
    return { payload: JSON.parse(raw), raw };
  } catch {
    return { payload: null, raw };
  }
}

function describeHttpError(response, raw) {
  const head = String(raw || "").replace(/\s+/g, " ").trim().slice(0, 180);
  const body = head ? `返回内容：${head}` : "返回内容为空";
  const hint = HTTP_HINTS[response.status];
  return (
    `HTTP ${response.status} ${response.statusText || ""}｜服务端未返回 JSON，${body}` +
    (hint ? `（提示：${hint}）` : "")
  );
}

/* ---------- 转换 ---------- */
async function convert() {
  if (state.files.length === 0) return;
  el.convertBtn.disabled = true;
  el.results.innerHTML = "";
  el.errors.innerHTML = "";

  const total = state.files.length;
  const estimate = total * SECONDS_PER_FILE;
  const startedAt = Date.now();
  const progressText = (extra) =>
    `<span class="spinner"></span>正在转换 ${total} 个文件${extra}，预计约 ${estimate} 秒，请勿关闭页面…`;

  setStatus(progressText(""));
  // 批量转换可能持续数分钟，用计时器给出已用时，避免看起来像卡死
  const timer = setInterval(() => {
    const elapsed = Math.round((Date.now() - startedAt) / 1000);
    setStatus(progressText(`，已用时 ${elapsed} 秒`));
  }, 1000);

  const form = new FormData();
  state.files.forEach((file) => form.append("files", file, file.name));
  form.append("enable_ocr", el.optOcr.checked ? "1" : "0");
  form.append("keep_art_layer", el.optArt.checked ? "1" : "0");
  form.append("detect_tables", el.optTables.checked ? "1" : "0");

  try {
    const response = await fetch("/api/convert", { method: "POST", body: form });
    const { payload, raw } = await readJsonResponse(response);
    if (!response.ok) {
      const detail = payload && payload.detail;
      const message = detail
        ? (Array.isArray(detail) ? detail.map((d) => d.msg).join("；") : detail)
        : describeHttpError(response, raw);
      setStatus(escapeHtml(message), "error");
      showErrors(Array.isArray(detail) ? detail.map((d) => d.msg) : [message]);
      return;
    }
    if (!payload) {
      const message = describeHttpError(response, raw);
      setStatus(escapeHtml(message), "error");
      showErrors([message]);
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
    clearInterval(timer);
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
    const { payload, raw } = await readJsonResponse(response);
    if (!response.ok) {
      throw new Error((payload && payload.detail) || describeHttpError(response, raw));
    }
    if (!payload) {
      throw new Error(describeHttpError(response, raw));
    }
    el.dialogBody.innerHTML = renderReport(payload);
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
        detail = item.kind === "complex_vector_layer"
          ? "复杂矢量路径（独立图片对象）"
          : `位图 ${(item.pixel || []).join("×")} · ${(item.format || "").toUpperCase()}`;
      } else if (item.type === "shape") {
        detail = `${item.kind} · 填充 ${item.fill || "无"} · 描边 ${item.stroke || "无"}`;
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
      state.maxFiles = data.max_files || state.maxFiles;
    }
  })
  .catch(() => {});