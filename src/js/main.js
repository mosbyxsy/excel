/**
 * 文件加载、远程配置、页面事件绑定与应用初始化入口。
 * 此文件由浏览器原生 ES Modules 直接加载，不依赖 npm、打包器或构建脚本。
 */

import {
  config,
  dom,
  state,
  fileExtension,
  nameFromUrl,
  startupRequest,
  setStartupLoading,
  finishStartupLoading,
  formatBytes,
  setStatus,
  showLoading,
  hideLoading,
  setCopyEnabled,
  copyRenderedCell,
  nextPaint
} from "./core.js";

import { parseWorkbook } from "./parsers/workbook-parser.js";

import {
  setWorkbook,
  setSourceMessage,
  enterFileSelection,
  beginFileLoad,
  setView,
  renderCurrentSheet,
  renderVirtualWindow,
  syncRawAxesScroll,
  revealRawMatch,
  stepRawMatch,
  // 行号、列号的点击事件由入口模块统一代理，因此需要显式导入固定切换函数。
  toggleRawAxisPin
} from "./viewer/viewer.js";

/* ======================================================================== */
/* 6. 本地文件、远程 URL 和页面事件                                         */
/* ======================================================================== */

async function loadArrayBuffer(buffer, fileName, typeHint, byteLength, sequence, sourcePath, autoFit) {
  showLoading("正在解析工作表", `${fileName} · ${formatBytes(byteLength)}`);
  setStatus("正在读取文件内容…");
  await nextPaint();
  try {
    const workbook = await parseWorkbook(buffer, fileName, typeHint);
    if (sequence !== state.loadSequence) return;
    // 解析完成后先让浏览器绘制“正在打开”阶段，再执行工作表标签和表格 DOM 渲染。
    showLoading("正在打开工作表", fileName);
    await nextPaint();
    if (sequence !== state.loadSequence) return;
    setWorkbook(workbook, byteLength, sourcePath, autoFit);
  } catch (error) {
    if (sequence !== state.loadSequence) return;
    showLoadError(error, fileName);
  } finally {
    if (sequence === state.loadSequence) hideLoading();
  }
}

async function loadLocalFile(file) {
  if (!file) return;
  if (!isSupportedName(file.name)) {
    setSourceMessage("请选择 XLSX、XLSM、XLS 或 CSV 文件。");
    return;
  }
  if (state.fetchController) state.fetchController.abort();
  const sequence = beginFileLoad(file.name);
  showLoading("正在读取本地文件", `${file.name} · ${formatBytes(file.size)}`);
  try {
    const buffer = await file.arrayBuffer();
    if (sequence !== state.loadSequence) return;
    // 浏览器不会暴露本地文件的真实可读取路径，因此本地文件成功打开后清除 URL 参数。
    await loadArrayBuffer(buffer, file.name, "", file.size, sequence, "", false);
  } catch (error) {
    if (sequence !== state.loadSequence) return;
    hideLoading();
    showLoadError(error, file.name);
  } finally {
    // 允许用户连续选择同一个文件时仍触发 change 事件。
    dom.localFile.value = "";
  }
}

async function loadRemoteFile(url, item) {
  const trimmedUrl = String(url || "").trim();
  if (!trimmedUrl) {
    setSourceMessage("请输入文件 URL 或选择一个预置文件。");
    return;
  }
  if (state.fetchController) state.fetchController.abort();
  const controller = new AbortController();
  state.fetchController = controller;
  const fileName = item && item.name ? item.name : nameFromUrl(trimmedUrl);
  const sequence = beginFileLoad(fileName);
  showLoading("正在下载在线文件", fileName);
  setStatus("正在请求远程文件…");
  try {
    const response = await fetch(trimmedUrl, {
      method: "GET",
      mode: "cors",
      credentials: "same-origin",
      signal: controller.signal
    });
    if (!response.ok) throw new Error(`服务器返回 HTTP ${response.status} ${response.statusText || ""}`.trim());
    const buffer = await response.arrayBuffer();
    if (sequence !== state.loadSequence) return;
    const responseName = fileNameFromDisposition(response.headers.get("content-disposition")) || fileName;
    // 保存用户实际输入或配置中的路径；相对路径不会被强制改写成绝对地址。
    await loadArrayBuffer(
      buffer,
      responseName,
      item && item.type,
      buffer.byteLength,
      sequence,
      trimmedUrl,
      Boolean(item && item.autoFit)
    );
  } catch (error) {
    if (error.name === "AbortError") return;
    if (sequence !== state.loadSequence) return;
    hideLoading();
    if (error instanceof TypeError) {
      showLoadError(
        new Error("无法读取远程文件。请确认 URL 正确，并且文件服务器已允许 CORS 跨域访问。"),
        fileName
      );
    } else {
      showLoadError(error, fileName);
    }
  } finally {
    if (state.fetchController === controller) state.fetchController = null;
  }
}

function fileNameFromDisposition(disposition) {
  if (!disposition) return "";
  const utf8 = disposition.match(/filename\*=UTF-8''([^;]+)/i);
  if (utf8) {
    try { return decodeURIComponent(utf8[1].replace(/["']/g, "")); } catch (_error) { return ""; }
  }
  const plain = disposition.match(/filename="?([^";]+)"?/i);
  return plain ? plain[1].trim() : "";
}

function showLoadError(error, fileName) {
  const message = error && error.message ? error.message : "未知错误";
  const passwordHint = /password|encrypt|CFB|Unsupported/i.test(message)
    ? " 文件可能受密码保护或包含当前查看器不支持的结构。"
    : "";
  enterFileSelection(`无法打开“${fileName}”：${message}${passwordHint}`);
}

function isSupportedName(name) {
  return ["xlsx", "xlsm", "xls", "csv"].includes(fileExtension(name));
}

/**
 * 清理 config.js 或远程 JSON 中的预置文件项。
 * id、name、url 缺一不可；type 和 action 只接受已知值。
 * autoFit 只有严格为 true 时才开启，缺省或其他值一律按 false 处理。
 */
function normalizePresetFiles(files) {
  if (!Array.isArray(files)) return [];
  const supportedTypes = new Set(["xlsx", "xlsm", "xls", "csv"]);
  return files.reduce((result, item) => {
    if (!item || item.id == null || !String(item.name || "").trim() || !String(item.url || "").trim()) {
      return result;
    }
    const type = String(item.type || "").trim().toLowerCase();
    result.push({
      id: String(item.id),
      name: String(item.name).trim(),
      url: String(item.url).trim(),
      type: supportedTypes.has(type) ? type : "",
      action: String(item.action || "").trim().toLowerCase(),
      autoFit: item.autoFit === true
    });
    return result;
  }, []);
}

/**
 * URL 中的 config 清单优先排列；若 id 重复，也优先采用远程清单中的定义。
 * 这样“多个 action: open 时取第一个”的顺序稳定且符合显式 URL 配置优先原则。
 */
function mergePresetFiles(remoteFiles, configuredFiles) {
  const merged = [];
  const usedIds = new Set();
  for (const file of [...remoteFiles, ...configuredFiles]) {
    if (usedIds.has(file.id)) continue;
    usedIds.add(file.id);
    merged.push(file);
  }
  return merged;
}

/**
 * 根据触发按钮在视口中的实时位置布置顶层菜单。
 *
 * 菜单优先向下打开；下方空间不足且上方更宽裕时改为向上打开。最大高度始终
 * 限制在可视区域内，选项过多时只滚动菜单自身，不再让整个文件选择卡片滚动。
 */
function positionPresetMenu() {
  const triggerRect = dom.presetTrigger.getBoundingClientRect();
  const viewportWidth = document.documentElement.clientWidth;
  const viewportHeight = document.documentElement.clientHeight;
  const viewportGap = 8;
  const menuGap = 6;
  const preferredHeight = 220;
  const minimumUsefulHeight = 96;
  const availableBelow = viewportHeight - triggerRect.bottom - menuGap - viewportGap;
  const availableAbove = triggerRect.top - menuGap - viewportGap;
  const openAbove = availableBelow < minimumUsefulHeight && availableAbove > availableBelow;
  const availableHeight = Math.max(
    64,
    Math.min(preferredHeight, openAbove ? availableAbove : availableBelow)
  );
  const menuWidth = Math.max(
    0,
    Math.min(triggerRect.width, viewportWidth - viewportGap * 2)
  );
  const menuLeft = Math.max(
    viewportGap,
    Math.min(triggerRect.left, viewportWidth - viewportGap - menuWidth)
  );

  dom.presetOptions.style.width = `${menuWidth}px`;
  dom.presetOptions.style.left = `${menuLeft}px`;
  dom.presetOptions.style.maxHeight = `${availableHeight}px`;
  if (openAbove) {
    dom.presetOptions.style.top = "auto";
    dom.presetOptions.style.bottom = `${viewportHeight - triggerRect.top + menuGap}px`;
  } else {
    dom.presetOptions.style.top = `${triggerRect.bottom + menuGap}px`;
    dom.presetOptions.style.bottom = "auto";
  }
}

/** 展开或收起自定义下拉，同时同步无障碍状态。 */
function setPresetMenuOpen(open, focusSelected) {
  const canOpen = Boolean(open && state.presetFiles.length && !dom.presetTrigger.disabled);
  dom.presetFile.classList.toggle("is-open", canOpen);
  dom.presetTrigger.setAttribute("aria-expanded", String(canOpen));
  if (canOpen) positionPresetMenu();
  dom.presetOptions.hidden = !canOpen;
  if (canOpen && focusSelected) {
    const selected = dom.presetOptions.querySelector('[aria-selected="true"]');
    const first = dom.presetOptions.querySelector(".preset-option");
    (selected || first)?.focus();
  }
}

/** 选择一个预置文件，仅改变控件状态；真正读取由“打开”按钮触发。 */
function choosePreset(id, closeMenu) {
  const item = state.presetFiles.find((file) => file.id === String(id)) || null;
  state.selectedPresetId = item ? item.id : "";
  dom.presetValue.textContent = item ? item.name : state.presetFiles.length ? "选择一个预置文件" : "暂无预置文件";
  dom.presetValue.title = item ? item.name : "";
  dom.loadPreset.disabled = !item;
  for (const option of dom.presetOptions.querySelectorAll(".preset-option")) {
    option.setAttribute("aria-selected", String(option.dataset.id === state.selectedPresetId));
  }
  if (closeMenu) {
    setPresetMenuOpen(false, false);
    dom.presetTrigger.focus();
  }
}

/** 根据合并后的预置数据重新绘制自定义列表框。 */
function populatePresetFiles() {
  dom.presetOptions.replaceChildren();
  for (const file of state.presetFiles) {
    const option = document.createElement("button");
    option.type = "button";
    option.className = "preset-option";
    option.role = "option";
    option.dataset.id = file.id;
    option.setAttribute("aria-selected", "false");
    option.title = file.name;
    const label = document.createElement("span");
    label.textContent = file.name;
    option.appendChild(label);
    option.addEventListener("click", () => choosePreset(file.id, true));
    dom.presetOptions.appendChild(option);
  }

  dom.presetTrigger.disabled = state.presetFiles.length === 0;
  const selectedStillExists = state.presetFiles.some((file) => file.id === state.selectedPresetId);
  choosePreset(selectedStillExists ? state.selectedPresetId : "", false);
}

function selectedPreset() {
  return state.presetFiles.find((file) => file.id === state.selectedPresetId) || null;
}

/**
 * 读取 ?config= 指定的 JSON。格式必须是 { files: [] }；文件内容仅作为数据解析，
 * 所有名称仍通过 textContent 写入页面，不会执行清单中的 HTML 或脚本。
 */
async function loadRemoteConfig(configPath) {
  setStartupLoading("正在读取文件清单", configPath);
  let response;
  try {
    response = await fetch(configPath, { method: "GET", mode: "cors", credentials: "same-origin" });
  } catch (_error) {
    throw new Error("无法读取 config JSON，请确认路径正确且服务器允许 CORS 跨域访问。");
  }
  if (!response.ok) {
    throw new Error(`config JSON 返回 HTTP ${response.status} ${response.statusText || ""}`.trim());
  }

  let payload;
  try {
    payload = await response.json();
  } catch (_error) {
    throw new Error("config 文件不是有效的 JSON。");
  }
  if (!payload || Array.isArray(payload) || !Array.isArray(payload.files)) {
    throw new Error('config JSON 格式错误，应为 { "files": [] }。');
  }

  const remoteFiles = normalizePresetFiles(payload.files);
  state.presetFiles = mergePresetFiles(remoteFiles, normalizePresetFiles(config.files));
  populatePresetFiles();
}

dom.localFile.addEventListener("change", () => loadLocalFile(dom.localFile.files[0]));

// 拖拽事件必须阻止浏览器默认行为，否则浏览器会直接导航到被拖入的文件。
for (const eventName of ["dragenter", "dragover"]) {
  dom.dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dom.dropZone.classList.add("is-dragging");
  });
}
for (const eventName of ["dragleave", "drop"]) {
  dom.dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dom.dropZone.classList.remove("is-dragging");
  });
}
dom.dropZone.addEventListener("drop", (event) => {
  const file = event.dataTransfer && event.dataTransfer.files[0];
  loadLocalFile(file);
});

// 自定义下拉的鼠标、触摸和键盘交互；不依赖浏览器原生 select 弹层。
dom.presetTrigger.addEventListener("click", () => {
  const opening = dom.presetOptions.hidden;
  setPresetMenuOpen(opening, false);
});

dom.presetTrigger.addEventListener("keydown", (event) => {
  if (!["ArrowDown", "ArrowUp", "Enter", " "].includes(event.key)) return;
  event.preventDefault();
  setPresetMenuOpen(true, true);
});

dom.presetOptions.addEventListener("keydown", (event) => {
  const options = [...dom.presetOptions.querySelectorAll(".preset-option")];
  const currentIndex = options.indexOf(document.activeElement);
  if (event.key === "Escape") {
    event.preventDefault();
    setPresetMenuOpen(false, false);
    dom.presetTrigger.focus();
    return;
  }
  if (!["ArrowDown", "ArrowUp", "Home", "End"].includes(event.key) || !options.length) return;
  event.preventDefault();
  let nextIndex = currentIndex;
  if (event.key === "Home") nextIndex = 0;
  else if (event.key === "End") nextIndex = options.length - 1;
  else if (event.key === "ArrowDown") nextIndex = (Math.max(-1, currentIndex) + 1) % options.length;
  else nextIndex = (currentIndex - 1 + options.length) % options.length;
  options[nextIndex].focus();
});

document.addEventListener("click", (event) => {
  // 菜单已提升到 body，点击列表本身仍属于下拉内部操作，不能被判定为外部点击。
  if (!dom.presetFile.contains(event.target) && !dom.presetOptions.contains(event.target)) {
    setPresetMenuOpen(false, false);
  }
});

// 菜单使用视口坐标；窗口缩放或任意滚动容器移动时同步其锚点位置。
window.addEventListener("resize", () => {
  if (!dom.presetOptions.hidden) positionPresetMenu();
});
document.addEventListener("scroll", () => {
  if (!dom.presetOptions.hidden) positionPresetMenu();
}, true);

dom.loadPreset.addEventListener("click", () => {
  const item = selectedPreset();
  if (!item) {
    setSourceMessage("请先选择一个预置文件。");
    return;
  }
  loadRemoteFile(item.url, item);
});

dom.loadUrl.addEventListener("click", () => loadRemoteFile(dom.remoteUrl.value, null));
dom.remoteUrl.addEventListener("keydown", (event) => {
  if (event.key === "Enter") loadRemoteFile(dom.remoteUrl.value, null);
});

dom.viewSwitch.addEventListener("click", (event) => {
  const button = event.target.closest("button[data-view]");
  if (button) setView(button.dataset.view);
});

dom.copyToggle.addEventListener("click", () => {
  setCopyEnabled(!state.copyEnabled);
});

// “更换文件”不是打开叠加层，而是明确结束当前查看并回到初始选择状态。
dom.sourceToggle.addEventListener("click", () => enterFileSelection());

dom.search.addEventListener("input", () => {
  clearTimeout(state.searchTimer);
  state.searchTimer = window.setTimeout(() => {
    state.searchText = dom.search.value;
    state.rawMatchIndex = -1;
    dom.viewport.scrollTop = 0;
    dom.viewport.scrollLeft = 0;
    renderCurrentSheet();
    if (state.view === "raw" && state.rawMatches.length) {
      // 等当前表格或首个虚拟窗口挂载后，再定位第一个匹配单元格。
      requestAnimationFrame(revealRawMatch);
    }
  }, 160);
});

dom.search.addEventListener("keydown", (event) => {
  if (event.key !== "Enter" || state.view !== "raw") return;
  event.preventDefault();
  stepRawMatch(event.shiftKey ? -1 : 1);
});

dom.searchPrev.addEventListener("click", () => stepRawMatch(-1));
dom.searchNext.addEventListener("click", () => stepRawMatch(1));

// 使用事件委托，虚拟滚动中新创建的行列标记和数据单元格无需逐个绑定事件。
dom.gridFrame.addEventListener("click", (event) => {
  const pinTarget = event.target.closest("[data-pin-axis][data-pin-index]");
  if (pinTarget && dom.gridFrame.contains(pinTarget)) {
    event.preventDefault();
    toggleRawAxisPin(pinTarget.dataset.pinAxis, Number(pinTarget.dataset.pinIndex));
    return;
  }
  if (!state.copyEnabled) return;
  const cell = event.target.closest(".grid-cell:not(.is-row-number), .raw-table td");
  if (!cell || !dom.viewport.contains(cell)) return;
  copyRenderedCell(cell);
});

/** 非原生 button 的 th/div 轴标记支持 Enter 和空格键切换固定状态。 */
dom.gridFrame.addEventListener("keydown", (event) => {
  if (event.key !== "Enter" && event.key !== " ") return;
  const pinTarget = event.target.closest("[data-pin-axis][data-pin-index]");
  if (!pinTarget || !dom.gridFrame.contains(pinTarget)) return;
  event.preventDefault();
  toggleRawAxisPin(pinTarget.dataset.pinAxis, Number(pinTarget.dataset.pinIndex));
});

dom.viewport.addEventListener("scroll", () => {
  if (state.renderFrame) return;
  state.renderFrame = requestAnimationFrame(() => {
    state.renderFrame = 0;
    // 原始视图坐标轴位于滚动区外，只在这一帧同步对应方向的位移。
    syncRawAxesScroll();
    if (state.renderer) renderVirtualWindow(false);
  });
}, { passive: true });

window.addEventListener("resize", () => {
  if (!state.workbook || state.renderFrame) return;
  state.renderFrame = requestAnimationFrame(() => {
    state.renderFrame = 0;
    if (state.view === "raw" && state.rawFitEnabled) {
      // 自适应依赖视口实时宽度，横竖屏切换或窗口缩放后重新计算所有列宽。
      const scrollTop = dom.viewport.scrollTop;
      renderCurrentSheet();
      dom.viewport.scrollTop = scrollTop;
    } else if (state.renderer) {
      renderVirtualWindow(true);
    }
  });
});

/**
 * 页面启动顺序：建立空状态 -> 读取 config JSON -> 合并预置文件 -> 选择自动打开项。
 * 显式 file/path/url 始终优先，其次是第一个 action: "open"，最后才是 defaultFileId。
 */
async function initializePage() {
  const startup = startupRequest();
  state.configPath = startup.configPath;
  state.presetFiles = normalizePresetFiles(config.files);
  populatePresetFiles();
  // 首次建立空状态时保留查询参数，启动遮罩会阻止选择界面提前闪现。
  enterFileSelection("", true);

  if (startup.hasQuery) setStartupLoading("正在初始化查看器", "正在解析地址栏参数…");

  if (startup.hasConfig) {
    if (!startup.configPath) {
      enterFileSelection("config 参数不能为空。", true);
      finishStartupLoading();
      return;
    }
    try {
      await loadRemoteConfig(startup.configPath);
    } catch (error) {
      enterFileSelection(error.message || "无法读取 config JSON。", true);
      finishStartupLoading();
      return;
    }
  }

  const actionFile = state.presetFiles.find((file) => file.action === "open") || null;
  const defaultFile = config.defaultFileId == null
    ? null
    : state.presetFiles.find((file) => file.id === String(config.defaultFileId)) || null;
  const selectedFile = actionFile || defaultFile;

  if (startup.filePath) {
    dom.remoteUrl.value = startup.filePath;
    setStartupLoading("正在打开工作表", startup.filePath);
    await loadRemoteFile(startup.filePath, null);
  } else if (selectedFile) {
    choosePreset(selectedFile.id, false);
    dom.remoteUrl.value = selectedFile.url;
    setStartupLoading("正在打开预置文件", selectedFile.name);
    await loadRemoteFile(selectedFile.url, selectedFile);
  }

  // 清单和可选工作簿均已处理完成，此时才允许用户看到最终界面。
  finishStartupLoading();
}

// 将列表框提升到 body，脱离 source-panel 的 overflow 裁切和滚动范围。
document.body.appendChild(dom.presetOptions);
initializePage();
