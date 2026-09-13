/* global ExcelJS, XLSX */

/**
 * 轻表格：零构建的 Excel / CSV 只读查看器。
 *
 * 文件被划分为以下部分：
 * 1. 常量、DOM 引用和全局状态；
 * 2. 通用工具函数；
 * 3. ExcelJS、SheetJS 与 CSV 解析；
 * 4. 工作簿和工作表状态切换；
 * 5. 原始视图、数据视图与虚拟滚动；
 * 6. 本地文件、远程 URL 和页面事件。
 *
 * 整个实现只读取文件并生成安全的文本节点，不会执行工作簿中的公式、宏或 HTML。
 */

(function () {
  "use strict";

  /* ======================================================================== */
  /* 1. 常量、DOM 引用和全局状态                                             */
  /* ======================================================================== */

  const DEFAULT_CONFIG = Object.freeze({
    files: [],
    defaultFileId: null,
    virtualizationThreshold: 500,
    overscanRows: 10
  });

  /** 防止异常的工作表范围导致浏览器分配失控。该限制远高于目标的一万行场景。 */
  const SAFETY_LIMITS = Object.freeze({ rows: 100000, columns: 1000 });
  const DATA_ROW_HEIGHT = 38;
  const DEFAULT_RAW_ROW_HEIGHT = 28;
  const ROW_NUMBER_WIDTH = 56;
  const MIN_COLUMN_WIDTH = 64;
  const MAX_COLUMN_WIDTH = 420;

  /**
   * 将部署配置与默认值合并，并将数字参数限制在合理区间。
   * 这样即便 config.js 中误填了负数或字符串，查看器也能继续工作。
   */
  const suppliedConfig = window.EXCEL_VIEWER_CONFIG || {};
  const config = {
    files: Array.isArray(suppliedConfig.files) ? suppliedConfig.files : DEFAULT_CONFIG.files,
    defaultFileId: suppliedConfig.defaultFileId || DEFAULT_CONFIG.defaultFileId,
    virtualizationThreshold: clampInteger(
      suppliedConfig.virtualizationThreshold,
      50,
      10000,
      DEFAULT_CONFIG.virtualizationThreshold
    ),
    overscanRows: clampInteger(
      suppliedConfig.overscanRows,
      2,
      50,
      DEFAULT_CONFIG.overscanRows
    )
  };

  const dom = {
    localFile: document.getElementById("local-file"),
    dropZone: document.getElementById("drop-zone"),
    presetFile: document.getElementById("preset-file"),
    presetTrigger: document.getElementById("preset-trigger"),
    presetValue: document.getElementById("preset-value"),
    presetOptions: document.getElementById("preset-options"),
    loadPreset: document.getElementById("load-preset"),
    remoteUrl: document.getElementById("remote-url"),
    loadUrl: document.getElementById("load-url"),
    sourcePanel: document.getElementById("source-panel"),
    sourceMessage: document.getElementById("source-message"),
    sourceToggle: document.getElementById("source-toggle"),
    viewerCard: document.getElementById("viewer-card"),
    fileName: document.getElementById("file-name"),
    fileMeta: document.getElementById("file-meta"),
    copyToggle: document.getElementById("copy-toggle"),
    copyToast: document.getElementById("copy-toast"),
    viewSwitch: document.getElementById("view-switch"),
    search: document.getElementById("sheet-search"),
    searchCount: document.getElementById("search-count"),
    searchPrev: document.getElementById("search-prev"),
    searchNext: document.getElementById("search-next"),
    sheetBar: document.getElementById("sheet-bar"),
    sheetTabs: document.getElementById("sheet-tabs"),
    statusLine: document.getElementById("status-line"),
    statusText: document.getElementById("status-text"),
    empty: document.getElementById("grid-empty"),
    viewport: document.getElementById("grid-viewport"),
    header: document.getElementById("grid-header"),
    body: document.getElementById("grid-body"),
    loading: document.getElementById("loading-overlay"),
    loadingTitle: document.getElementById("loading-title"),
    loadingDetail: document.getElementById("loading-detail"),
    startupLoader: document.getElementById("startup-loader"),
    startupLoadingTitle: document.getElementById("startup-loading-title"),
    startupLoadingDetail: document.getElementById("startup-loading-detail")
  };

  /**
   * 页面运行时状态只保存在这一处。
   * renderer 保存当前渲染模型，用于滚动事件中快速确定应该挂载哪些行。
   */
  const state = {
    workbook: null,
    sheetIndex: 0,
    view: "raw",
    searchText: "",
    rawMatches: [],
    rawMatchLookup: new Map(),
    rawMatchIndex: -1,
    sort: { column: -1, direction: null },
    copyEnabled: false,
    copyToastTimer: 0,
    renderer: null,
    renderFrame: 0,
    searchTimer: 0,
    fetchController: null,
    loadSequence: 0,
    presetFiles: [],
    selectedPresetId: "",
    configPath: ""
  };

  /* ======================================================================== */
  /* 2. 通用工具函数                                                         */
  /* ======================================================================== */

  function clampInteger(value, min, max, fallback) {
    const number = Number.parseInt(value, 10);
    if (!Number.isFinite(number)) return fallback;
    return Math.min(max, Math.max(min, number));
  }

  function clamp(value, min, max) {
    return Math.min(max, Math.max(min, value));
  }

  /** 将从零开始的列序号转换为 Excel 列名，例如 0 -> A、27 -> AB。 */
  function columnLetter(index) {
    let number = index + 1;
    let result = "";
    while (number > 0) {
      const remainder = (number - 1) % 26;
      result = String.fromCharCode(65 + remainder) + result;
      number = Math.floor((number - 1) / 26);
    }
    return result;
  }

  function fileExtension(name) {
    const match = String(name || "").toLowerCase().match(/\.([a-z0-9]+)(?:[?#].*)?$/);
    return match ? match[1] : "";
  }

  function nameFromUrl(url) {
    try {
      const parsed = new URL(url, window.location.href);
      const segment = parsed.pathname.split("/").filter(Boolean).pop();
      return segment ? decodeURIComponent(segment) : "在线工作表";
    } catch (_error) {
      return "在线工作表";
    }
  }

  /**
   * 解析页面启动参数。
   * - ?config=... 指向包含 { files: [] } 的 JSON 清单；
   * - ?file=、?path=、?url= 指定需要直接打开的文件；
   * - 没有键名的查询字符串继续按文件路径处理，兼容既有分享链接。
   */
  function startupRequest() {
    const rawQuery = window.location.search.slice(1).trim();
    if (!rawQuery) return { hasQuery: false, hasConfig: false, configPath: "", filePath: "" };

    const params = new URLSearchParams(rawQuery);
    const hasConfig = params.has("config");
    const configPath = hasConfig ? String(params.get("config") || "").trim() : "";
    const keyedName = ["file", "path", "url"].find((name) => params.has(name));
    if (hasConfig || keyedName) {
      return {
        hasQuery: true,
        hasConfig,
        configPath,
        filePath: keyedName ? String(params.get(keyedName) || "").trim() : ""
      };
    }

    try {
      return {
        hasQuery: true,
        hasConfig: false,
        configPath: "",
        filePath: decodeURIComponent(rawQuery).trim()
      };
    } catch (_error) {
      // 非法百分号编码不应阻断页面初始化，交给 fetch 输出更具体的路径错误。
      return { hasQuery: true, hasConfig: false, configPath: "", filePath: rawQuery };
    }
  }

  /**
   * 将已成功打开的在线文件路径同步到当前页面地址，但不触发页面跳转。
   * 整段路径使用 encodeURIComponent 编码，因此远程地址自身带有 ?、& 或 # 时，
   * 也不会与查看器的查询参数混淆；刷新后 startupRequest 会将其完整还原。
   * 如果页面由 ?config= 启动，则保留清单路径并将当前文件写入 file 参数。
   * 传入空字符串表示当前文件来自本地选择器，此时移除旧的在线文件参数。
   */
  function replaceStartupFilePath(filePath) {
    const nextUrl = new URL(window.location.href);
    const normalizedPath = String(filePath || "").trim();
    if (state.configPath) {
      nextUrl.search = "";
      nextUrl.searchParams.set("config", state.configPath);
      if (normalizedPath) nextUrl.searchParams.set("file", normalizedPath);
    } else {
      nextUrl.search = normalizedPath ? `?${encodeURIComponent(normalizedPath)}` : "";
    }
    window.history.replaceState(window.history.state, "", nextUrl.href);
  }

  /** 更新全屏启动动画中的阶段说明。 */
  function setStartupLoading(title, detail) {
    dom.startupLoadingTitle.textContent = title || "正在准备工作表";
    dom.startupLoadingDetail.textContent = detail || "请稍候…";
  }

  /** 完成查询参数启动流程，并确保遮罩不再拦截页面操作。 */
  function finishStartupLoading() {
    document.documentElement.classList.remove("is-startup-loading");
    dom.startupLoader.setAttribute("aria-hidden", "true");
  }

  function formatBytes(bytes) {
    if (!Number.isFinite(bytes) || bytes < 0) return "未知大小";
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
  }

  function setStatus(message, type) {
    dom.statusText.textContent = message;
    dom.statusLine.classList.toggle("is-warning", type === "warning");
    dom.statusLine.classList.toggle("is-error", type === "error");
    const icon = dom.statusLine.querySelector(".status-icon");
    if (icon) icon.textContent = type === "error" ? "×" : type === "warning" ? "!" : "i";
  }

  function showLoading(title, detail) {
    dom.loadingTitle.textContent = title || "正在读取文件";
    dom.loadingDetail.textContent = detail || "请稍候…";
    dom.loading.hidden = false;
  }

  function hideLoading() {
    dom.loading.hidden = true;
  }

  /** 在查看器底部短暂显示复制反馈，不改写右侧的工作表状态信息。 */
  function showCopyToast(message, isError) {
    clearTimeout(state.copyToastTimer);
    dom.copyToast.textContent = message;
    dom.copyToast.classList.toggle("is-error", Boolean(isError));
    dom.copyToast.hidden = false;
    state.copyToastTimer = window.setTimeout(hideCopyToast, 1600);
  }

  function hideCopyToast() {
    clearTimeout(state.copyToastTimer);
    state.copyToastTimer = 0;
    dom.copyToast.hidden = true;
  }

  /**
   * 优先使用现代 Clipboard API；在非安全来源中不可用时，回退到临时 textarea。
   * 临时节点只承载纯文本，复制完成后立即移除。
   */
  async function writeClipboardText(text) {
    if (navigator.clipboard && typeof navigator.clipboard.writeText === "function") {
      try {
        await navigator.clipboard.writeText(text);
        return true;
      } catch (_error) {
        // 权限或安全上下文限制会进入下方兼容路径。
      }
    }

    const textarea = document.createElement("textarea");
    textarea.value = text;
    textarea.readOnly = true;
    textarea.style.position = "fixed";
    textarea.style.left = "-9999px";
    document.body.appendChild(textarea);
    textarea.select();
    let copied = false;
    try {
      copied = Boolean(document.execCommand && document.execCommand("copy"));
    } catch (_error) {
      copied = false;
    }
    textarea.remove();
    return copied;
  }

  /** 切换点击复制模式，并同步按钮的视觉状态、无障碍状态与单元格鼠标样式。 */
  function setCopyEnabled(enabled) {
    state.copyEnabled = Boolean(enabled);
    dom.copyToggle.setAttribute("aria-pressed", String(state.copyEnabled));
    dom.copyToggle.classList.toggle("is-active", state.copyEnabled);
    dom.copyToggle.title = state.copyEnabled ? "关闭点击单元格复制" : "开启点击单元格复制";
    dom.viewerCard.classList.toggle("is-copy-mode", state.copyEnabled);
    showCopyToast(state.copyEnabled ? "已开启点击复制" : "已关闭点击复制", false);
  }

  /** 复制当前单元格的格式化显示文本，并在反馈中标明 Excel 地址。 */
  async function copyRenderedCell(element) {
    const text = element.dataset.copyText || "";
    const rowIndex = Number(element.dataset.rowIndex);
    const columnIndex = Number(element.dataset.columnIndex);
    const hasAddress = Number.isInteger(rowIndex) && Number.isInteger(columnIndex);
    const address = hasAddress ? `${columnLetter(columnIndex)}${rowIndex + 1}` : "当前单元格";
    const copied = await writeClipboardText(text);
    showCopyToast(copied ? `已复制 ${address}` : `无法复制 ${address}，请检查浏览器权限`, !copied);
  }

  /**
   * 让浏览器有机会绘制加载状态。
   * 大文件解析本身在主线程运行，先让出一帧可避免用户看不到任何反馈。
   */
  function nextPaint() {
    return new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
  }

  function getCurrentSheet() {
    return state.workbook && state.workbook.sheets[state.sheetIndex];
  }

  function isRowEmpty(row) {
    return !row || !row.cells.some((cell) => cell && String(cell.text).trim() !== "");
  }

  /**
   * Excel 使用磅表示行高；CSS 使用像素。按 96 DPI 将 1pt 换算为 4/3px。
   * 限制最大值可以避免损坏文件声明数万像素行高时破坏滚动体验。
   */
  function pointsToPixels(points) {
    const value = Number(points);
    if (!Number.isFinite(value) || value <= 0) return DEFAULT_RAW_ROW_HEIGHT;
    return clamp(Math.round(value * 96 / 72), 20, 240);
  }

  /** Excel 列宽并非像素，常用近似式 width * 7 + 5 对浏览器展示足够稳定。 */
  function excelWidthToPixels(width) {
    const value = Number(width);
    if (!Number.isFinite(value) || value <= 0) return 120;
    return clamp(Math.round(value * 7 + 5), MIN_COLUMN_WIDTH, MAX_COLUMN_WIDTH);
  }

  function normalizeColumnWidths(widths, count) {
    const result = [];
    for (let index = 0; index < count; index += 1) {
      const column = widths[index] || {};
      result.push({
        width: column.hidden ? 0 : clamp(column.width || 120, 0, MAX_COLUMN_WIDTH),
        hidden: Boolean(column.hidden)
      });
    }
    return result;
  }

  /**
   * 解析类似 A1:C8 的 Excel 区域地址。ExcelJS 与 SheetJS 都使用这种表示法，
   * 因而统一转换为从零开始的行列下标，后续渲染无需关心解析库差异。
   */
  function parseRangeAddress(address) {
    const match = String(address || "").replace(/\$/g, "").match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
    if (!match) return null;
    return {
      startCol: lettersToColumn(match[1]),
      startRow: Number(match[2]) - 1,
      endCol: lettersToColumn(match[3]),
      endRow: Number(match[4]) - 1
    };
  }

  function lettersToColumn(letters) {
    let result = 0;
    for (const character of String(letters).toUpperCase()) {
      result = result * 26 + character.charCodeAt(0) - 64;
    }
    return result - 1;
  }

  function colorToCss(color) {
    if (!color || typeof color !== "object") return "";
    if (typeof color.argb === "string") {
      const argb = color.argb.replace(/[^0-9a-f]/gi, "");
      if (argb.length === 8) return `#${argb.slice(2)}`;
      if (argb.length === 6) return `#${argb}`;
    }
    if (typeof color.rgb === "string") {
      const rgb = color.rgb.replace(/[^0-9a-f]/gi, "");
      if (rgb.length >= 6) return `#${rgb.slice(-6)}`;
    }
    // 主题色需要读取主题 XML 才能精确换算；无法解析时交给默认样式处理。
    return "";
  }

  function normalizeBorderSide(side) {
    if (!side || !side.style) return null;
    const widths = {
      hair: 1,
      thin: 1,
      dotted: 1,
      dashed: 1,
      medium: 2,
      double: 3,
      thick: 3
    };
    const styles = {
      dotted: "dotted",
      dashed: "dashed",
      dashDot: "dashed",
      dashDotDot: "dashed",
      double: "double"
    };
    return {
      width: widths[side.style] || 1,
      style: styles[side.style] || "solid",
      color: colorToCss(side.color) || "#9eaaa6"
    };
  }

  /** 将两个解析库的样式对象统一为页面内部使用的轻量格式。 */
  function normalizeCellStyle(source) {
    if (!source || typeof source !== "object") return null;
    const font = source.font || {};
    const fill = source.fill || {};
    const alignment = source.alignment || {};
    const border = source.border || {};
    const fillColor = colorToCss(fill.fgColor || fill.bgColor);
    const fontColor = colorToCss(font.color);
    return {
      fontName: typeof font.name === "string" ? font.name : "",
      fontSize: Number.isFinite(font.size) ? clamp(font.size, 7, 48) : null,
      bold: Boolean(font.bold),
      italic: Boolean(font.italic),
      underline: Boolean(font.underline),
      strike: Boolean(font.strike),
      fontColor,
      fillColor: fill.type === "pattern" || fill.patternType || fillColor ? fillColor : "",
      horizontal: alignment.horizontal || "",
      vertical: alignment.vertical || "",
      wrapText: Boolean(alignment.wrapText),
      rotation: Number.isFinite(alignment.textRotation) ? alignment.textRotation : 0,
      borders: {
        top: normalizeBorderSide(border.top),
        right: normalizeBorderSide(border.right),
        bottom: normalizeBorderSide(border.bottom),
        left: normalizeBorderSide(border.left)
      }
    };
  }

  /**
   * 只通过 element.style 设置经过白名单筛选的属性。
   * 单元格值始终使用 textContent，不能借样式或内容注入 HTML。
   */
  function applyCellStyle(element, style) {
    if (!style) return;
    if (style.fontName) element.style.fontFamily = style.fontName;
    // Excel 字号单位为磅，按 96 DPI 换算为 CSS 像素。
    if (style.fontSize) element.style.fontSize = `${Math.round(style.fontSize * 96 / 72 * 10) / 10}px`;
    if (style.bold) element.style.fontWeight = "700";
    if (style.italic) element.style.fontStyle = "italic";
    if (style.underline || style.strike) {
      element.style.textDecoration = [style.underline && "underline", style.strike && "line-through"]
        .filter(Boolean)
        .join(" ");
    }
    if (style.fontColor) element.style.color = style.fontColor;
    if (style.fillColor) element.style.backgroundColor = style.fillColor;
    if (style.horizontal) {
      const alignmentMap = { center: "center", right: "flex-end", left: "flex-start", justify: "space-between" };
      element.style.justifyContent = alignmentMap[style.horizontal] || "flex-start";
      element.style.textAlign = style.horizontal;
    }
    if (style.vertical) {
      const verticalMap = { top: "flex-start", middle: "center", bottom: "flex-end" };
      element.style.alignItems = verticalMap[style.vertical] || "center";
      // 标准 table-cell 不响应 flex 的 align-items，需要额外设置 vertical-align。
      element.style.verticalAlign = style.vertical === "middle" ? "middle" : style.vertical;
    }
    if (style.wrapText) {
      element.style.whiteSpace = "normal";
      element.style.overflowWrap = "anywhere";
    }
    if (style.rotation && style.rotation !== 255) {
      element.style.transform = `rotate(${clamp(style.rotation, -90, 90)}deg)`;
    }
    for (const side of ["top", "right", "bottom", "left"]) {
      const border = style.borders && style.borders[side];
      if (border) {
        element.style[`border${side[0].toUpperCase()}${side.slice(1)}`] =
          `${border.width}px ${border.style} ${border.color}`;
      }
    }
  }

  function buildCell(text, raw, style, options) {
    return {
      text: text == null ? "" : String(text),
      raw: raw == null ? "" : raw,
      style: style || null,
      formulaMissing: Boolean(options && options.formulaMissing)
    };
  }

  /* ======================================================================== */
  /* 3. ExcelJS、SheetJS 与 CSV 解析                                          */
  /* ======================================================================== */

  /**
   * 根据扩展名、显式提示和文件头判断格式。
   * ZIP 文件头 PK\x03\x04 对应 xlsx/xlsm，D0 CF 11 E0 对应旧版 OLE xls。
   */
  function inferFileType(name, typeHint, buffer) {
    const hint = String(typeHint || "").toLowerCase();
    if (["xlsx", "xlsm", "xls", "csv"].includes(hint)) return hint;

    const extension = fileExtension(name);
    if (["xlsx", "xlsm", "xls", "csv"].includes(extension)) return extension;

    const bytes = new Uint8Array(buffer, 0, Math.min(buffer.byteLength, 8));
    if (bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x03 && bytes[3] === 0x04) return "xlsx";
    if (bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0) return "xls";
    return "csv";
  }

  /**
   * ExcelJS 不会执行公式。若工作簿保存了缓存结果则显示结果；否则显示公式本身，
   * 并设置 formulaMissing，渲染时会追加提示标记。
   */
  function excelJsCellValue(cell) {
    const value = cell.value;
    if (value && typeof value === "object" && Object.prototype.hasOwnProperty.call(value, "formula")) {
      if (value.result == null) {
        return { text: `=${value.formula}`, raw: "", formulaMissing: true };
      }
      return {
        text: formatDisplayValue(value.result, cell.numFmt),
        raw: value.result,
        formulaMissing: false
      };
    }
    if (value && typeof value === "object" && Array.isArray(value.richText)) {
      const text = value.richText.map((part) => part.text || "").join("");
      return { text, raw: text, formulaMissing: false };
    }
    if (value && typeof value === "object" && Object.prototype.hasOwnProperty.call(value, "hyperlink")) {
      return { text: value.text || value.hyperlink || "", raw: value.text || "", formulaMissing: false };
    }
    return {
      text: formatDisplayValue(value, cell.numFmt),
      raw: value,
      formulaMissing: false
    };
  }

  function formatDisplayValue(value, numberFormat) {
    if (value == null) return "";
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      const includesTime = /[hHsS]/.test(numberFormat || "");
      return new Intl.DateTimeFormat("zh-CN", {
        year: "numeric",
        month: "2-digit",
        day: "2-digit",
        ...(includesTime ? { hour: "2-digit", minute: "2-digit", second: "2-digit" } : {})
      }).format(value);
    }
    if (typeof value === "number") return formatNumber(value, numberFormat);
    if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
    if (typeof value === "object") {
      if (value.error) return String(value.error);
      if (value.text) return String(value.text);
      return String(value);
    }
    return String(value);
  }

  /**
   * 覆盖查看场景中最常见的百分比、千分位和小数位格式。
   * 复杂的会计格式仍保留数值，而不会冒充 Excel 的完整格式引擎。
   */
  function formatNumber(value, numberFormat) {
    const format = String(numberFormat || "");
    const decimalMatch = format.match(/0\.([0#]+)/);
    const decimals = decimalMatch ? decimalMatch[1].length : 0;
    const options = {
      minimumFractionDigits: decimals,
      maximumFractionDigits: Math.min(decimals, 12),
      useGrouping: format.includes(",")
    };
    if (format.includes("%")) {
      return new Intl.NumberFormat("zh-CN", { ...options, style: "percent" }).format(value);
    }
    const formatted = new Intl.NumberFormat("zh-CN", options).format(value);
    const currency = format.match(/[¥￥$€£]/);
    return currency ? `${currency[0]}${formatted}` : formatted;
  }

  async function parseWithExcelJs(buffer, fileName) {
    if (typeof window.ExcelJS === "undefined") {
      throw new Error("ExcelJS 未加载，无法读取 XLSX/XLSM 文件。请检查网络或 CDN 配置。");
    }

    const workbook = new window.ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const sheets = [];

    workbook.eachSheet((worksheet) => {
      const rowCount = Math.max(worksheet.actualRowCount || 0, worksheet.rowCount || 0);
      const mergeRanges = Array.isArray(worksheet.model && worksheet.model.merges)
        ? worksheet.model.merges.map(parseRangeAddress).filter(Boolean)
        : [];
      const mergeMaxRow = mergeRanges.reduce((max, range) => Math.max(max, range.endRow + 1), 0);
      const mergeMaxCol = mergeRanges.reduce((max, range) => Math.max(max, range.endCol + 1), 0);
      const maxRows = Math.max(rowCount, mergeMaxRow);
      const maxCols = Math.max(worksheet.actualColumnCount || 0, worksheet.columnCount || 0, mergeMaxCol, 1);
      assertSafeDimensions(maxRows, maxCols);

      const rows = [];
      const rowHeights = [];
      for (let rowIndex = 0; rowIndex < maxRows; rowIndex += 1) {
        const worksheetRow = worksheet.getRow(rowIndex + 1);
        const cells = new Array(maxCols).fill(null);
        for (let columnIndex = 0; columnIndex < maxCols; columnIndex += 1) {
          const cell = worksheetRow.getCell(columnIndex + 1);
          if (cell.value == null && !cell.hasStyle) continue;
          const value = excelJsCellValue(cell);
          cells[columnIndex] = buildCell(
            value.text,
            value.raw,
            normalizeCellStyle(cell.style),
            { formulaMissing: value.formulaMissing }
          );
        }
        rows.push({ cells, hidden: Boolean(worksheetRow.hidden), sourceIndex: rowIndex });
        rowHeights.push(pointsToPixels(worksheetRow.height));
      }

      const widths = [];
      for (let columnIndex = 0; columnIndex < maxCols; columnIndex += 1) {
        const column = worksheet.getColumn(columnIndex + 1);
        widths.push({
          width: excelWidthToPixels(column.width),
          hidden: Boolean(column.hidden)
        });
      }

      sheets.push(finalizeSheet({
        name: worksheet.name || `Sheet ${sheets.length + 1}`,
        rows,
        maxCols,
        rowHeights,
        colWidths: normalizeColumnWidths(widths, maxCols),
        merges: mergeRanges,
        source: "exceljs"
      }));
    });

    if (!sheets.length) throw new Error("工作簿中没有可显示的工作表。");
    return { name: fileName, type: fileExtension(fileName) || "xlsx", sheets, warnings: [] };
  }

  function assertSafeDimensions(rows, columns) {
    if (rows > SAFETY_LIMITS.rows || columns > SAFETY_LIMITS.columns) {
      throw new Error(
        `工作表范围过大（${rows.toLocaleString()} 行 × ${columns.toLocaleString()} 列），已超过安全显示限制。`
      );
    }
  }

  function parseWithSheetJs(buffer, fileName) {
    if (typeof window.XLSX === "undefined") {
      throw new Error("SheetJS 未加载，无法读取 XLS 文件。请检查网络或 CDN 配置。");
    }
    const workbook = window.XLSX.read(buffer, {
      type: "array",
      cellDates: true,
      cellNF: true,
      cellStyles: true,
      raw: false
    });
    const sheets = workbook.SheetNames.map((name) => sheetJsWorksheetToModel(workbook.Sheets[name], name));
    if (!sheets.length) throw new Error("工作簿中没有可显示的工作表。");
    return { name: fileName, type: fileExtension(fileName) || "xls", sheets, warnings: [] };
  }

  function sheetJsWorksheetToModel(worksheet, name) {
    const decodedRange = worksheet["!ref"] ? window.XLSX.utils.decode_range(worksheet["!ref"]) : null;
    const merges = Array.isArray(worksheet["!merges"])
      ? worksheet["!merges"].map((range) => ({
        startRow: range.s.r,
        startCol: range.s.c,
        endRow: range.e.r,
        endCol: range.e.c
      }))
      : [];
    const mergeMaxRow = merges.reduce((max, range) => Math.max(max, range.endRow + 1), 0);
    const mergeMaxCol = merges.reduce((max, range) => Math.max(max, range.endCol + 1), 0);
    const maxRows = Math.max(decodedRange ? decodedRange.e.r + 1 : 0, mergeMaxRow, 1);
    const maxCols = Math.max(decodedRange ? decodedRange.e.c + 1 : 0, mergeMaxCol, 1);
    assertSafeDimensions(maxRows, maxCols);

    const rows = [];
    const rowMetadata = worksheet["!rows"] || [];
    for (let rowIndex = 0; rowIndex < maxRows; rowIndex += 1) {
      const cells = new Array(maxCols).fill(null);
      for (let columnIndex = 0; columnIndex < maxCols; columnIndex += 1) {
        const address = window.XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
        const sourceCell = worksheet[address];
        if (!sourceCell) continue;
        const formulaMissing = Boolean(sourceCell.f && sourceCell.v == null);
        const text = formulaMissing
          ? `=${sourceCell.f}`
          : sourceCell.w != null
            ? sourceCell.w
            : formatDisplayValue(sourceCell.v, sourceCell.z);
        cells[columnIndex] = buildCell(
          text,
          sourceCell.v,
          normalizeCellStyle(sourceCell.s),
          { formulaMissing }
        );
      }
      const metadata = rowMetadata[rowIndex] || {};
      rows.push({ cells, hidden: Boolean(metadata.hidden), sourceIndex: rowIndex });
    }

    const columnMetadata = worksheet["!cols"] || [];
    const widths = [];
    for (let columnIndex = 0; columnIndex < maxCols; columnIndex += 1) {
      const column = columnMetadata[columnIndex] || {};
      const width = column.wpx || excelWidthToPixels(column.wch || column.width);
      widths.push({ width, hidden: Boolean(column.hidden) });
    }
    const rowHeights = rows.map((_row, index) => {
      const metadata = rowMetadata[index] || {};
      return metadata.hpx || pointsToPixels(metadata.hpt);
    });

    return finalizeSheet({
      name,
      rows,
      maxCols,
      rowHeights,
      colWidths: normalizeColumnWidths(widths, maxCols),
      merges,
      source: "sheetjs"
    });
  }

  /**
   * CSV 编码检测顺序：UTF-8 BOM -> 严格 UTF-8 -> GB18030。
   * TextDecoder 的 fatal 选项会让非法字节直接抛错，避免静默产生大量替换字符。
   */
  function decodeCsvBuffer(buffer) {
    const bytes = new Uint8Array(buffer);
    const hasUtf8Bom = bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf;
    if (hasUtf8Bom) {
      return {
        text: new TextDecoder("utf-8").decode(bytes.subarray(3)),
        encoding: "UTF-8 BOM"
      };
    }
    try {
      return { text: new TextDecoder("utf-8", { fatal: true }).decode(bytes), encoding: "UTF-8" };
    } catch (_utf8Error) {
      try {
        return { text: new TextDecoder("gb18030", { fatal: true }).decode(bytes), encoding: "GB18030" };
      } catch (_gbError) {
        throw new Error("无法按 UTF-8 或 GB18030 解码 CSV 文件。");
      }
    }
  }

  /**
   * 读取前 24 条逻辑记录，对四种候选分隔符评分。
   * 只统计引号外的字符，因此带逗号的带引号字段不会干扰判断。
   */
  function detectCsvDelimiter(text) {
    const candidates = [",", "\t", ";", "|"];
    const counts = new Map(candidates.map((candidate) => [candidate, []]));
    const current = new Map(candidates.map((candidate) => [candidate, 0]));
    let inQuotes = false;
    let recordCount = 0;

    for (let index = 0; index < text.length && recordCount < 24; index += 1) {
      const character = text[index];
      if (character === '"') {
        if (inQuotes && text[index + 1] === '"') {
          index += 1;
        } else {
          inQuotes = !inQuotes;
        }
      } else if (!inQuotes && candidates.includes(character)) {
        current.set(character, current.get(character) + 1);
      } else if (!inQuotes && (character === "\n" || character === "\r")) {
        if (character === "\r" && text[index + 1] === "\n") index += 1;
        for (const candidate of candidates) {
          counts.get(candidate).push(current.get(candidate));
          current.set(candidate, 0);
        }
        recordCount += 1;
      }
    }

    if (recordCount === 0 || Array.from(current.values()).some((value) => value > 0)) {
      for (const candidate of candidates) counts.get(candidate).push(current.get(candidate));
    }

    let best = { delimiter: ",", score: -Infinity, reliable: false };
    for (const candidate of candidates) {
      const values = counts.get(candidate).filter((value) => value > 0);
      if (!values.length) continue;
      const average = values.reduce((sum, value) => sum + value, 0) / values.length;
      const variance = values.reduce((sum, value) => sum + Math.abs(value - average), 0) / values.length;
      const coverage = values.length / Math.max(counts.get(candidate).length, 1);
      const score = average * 3 + coverage * 4 - variance * 2;
      if (score > best.score) {
        best = { delimiter: candidate, score, reliable: coverage >= 0.6 && average >= 1 };
      }
    }
    return best.reliable ? best : { delimiter: ",", score: best.score, reliable: false };
  }

  /**
   * 按 RFC 4180 的核心规则解析 CSV：
   * - 双引号包裹字段；
   * - 两个连续双引号表示一个字面双引号；
   * - 引号内部允许换行；
   * - 保留空字段和每行末尾的空字段。
   */
  function parseCsvRecords(text, delimiter) {
    const records = [];
    let record = [];
    let field = "";
    let inQuotes = false;
    let endedWithRecordBreak = false;

    for (let index = 0; index < text.length; index += 1) {
      const character = text[index];
      endedWithRecordBreak = false;
      if (inQuotes) {
        if (character === '"' && text[index + 1] === '"') {
          field += '"';
          index += 1;
        } else if (character === '"') {
          inQuotes = false;
        } else {
          field += character;
        }
      } else if (character === '"' && field.length === 0) {
        inQuotes = true;
      } else if (character === delimiter) {
        record.push(field);
        field = "";
      } else if (character === "\n" || character === "\r") {
        if (character === "\r" && text[index + 1] === "\n") index += 1;
        record.push(field);
        records.push(record);
        record = [];
        field = "";
        endedWithRecordBreak = true;
      } else {
        field += character;
      }
    }

    if (inQuotes) throw new Error("CSV 文件包含未闭合的双引号字段。");
    if (!endedWithRecordBreak || field !== "" || record.length > 0) {
      record.push(field);
      records.push(record);
    }
    return records;
  }

  function parseCsv(buffer, fileName) {
    const decoded = decodeCsvBuffer(buffer);
    const delimiterResult = detectCsvDelimiter(decoded.text);
    const records = parseCsvRecords(decoded.text, delimiterResult.delimiter);
    const maxCols = Math.max(1, ...records.map((record) => record.length));
    assertSafeDimensions(records.length, maxCols);

    const rows = records.map((record, rowIndex) => ({
      cells: Array.from({ length: maxCols }, (_unused, columnIndex) => {
        const value = record[columnIndex] == null ? "" : record[columnIndex];
        return buildCell(value, value, null);
      }),
      hidden: false,
      sourceIndex: rowIndex
    }));
    const sheet = finalizeSheet({
      name: "CSV",
      rows,
      maxCols,
      rowHeights: rows.map(() => DEFAULT_RAW_ROW_HEIGHT),
      colWidths: inferCsvColumnWidths(rows, maxCols),
      merges: [],
      source: "csv"
    });
    const delimiterNames = { ",": "逗号", "\t": "制表符", ";": "分号", "|": "竖线" };
    const warnings = [];
    if (!delimiterResult.reliable) warnings.push("未能可靠识别 CSV 分隔符，已默认使用逗号");
    return {
      name: fileName,
      type: "csv",
      sheets: [sheet],
      warnings,
      csvInfo: `${decoded.encoding} · ${delimiterNames[delimiterResult.delimiter]}分隔`
    };
  }

  /** 根据 CSV 前 100 行文本长度估算列宽，避免短列占用过多空间。 */
  function inferCsvColumnWidths(rows, maxCols) {
    const widths = [];
    for (let columnIndex = 0; columnIndex < maxCols; columnIndex += 1) {
      let maxLength = columnLetter(columnIndex).length;
      for (let rowIndex = 0; rowIndex < Math.min(rows.length, 100); rowIndex += 1) {
        const cell = rows[rowIndex].cells[columnIndex];
        maxLength = Math.max(maxLength, Array.from(cell ? cell.text : "").length);
      }
      widths.push({ width: clamp(maxLength * 8 + 24, 84, 280), hidden: false });
    }
    return widths;
  }

  /** 清理尾部完全空白行，并缓存工作表后续渲染所需的统计信息。 */
  function finalizeSheet(sheet) {
    let lastMeaningfulRow = sheet.rows.length - 1;
    while (lastMeaningfulRow > 0 && isRowEmpty(sheet.rows[lastMeaningfulRow])) lastMeaningfulRow -= 1;
    const mergeLastRow = sheet.merges.reduce((max, range) => Math.max(max, range.endRow), -1);
    const keepRows = Math.max(lastMeaningfulRow + 1, mergeLastRow + 1, 1);
    sheet.rows = sheet.rows.slice(0, keepRows);
    sheet.rowHeights = sheet.rowHeights.slice(0, keepRows);
    sheet.mergeLookup = null;
    sheet.hasVerticalMerges = sheet.merges.some((range) => range.endRow > range.startRow);
    sheet.nonHiddenRowCount = sheet.rows.filter((row) => !row.hidden).length;
    return sheet;
  }

  /**
   * 执行首选解析器；现代格式失败后用 SheetJS 尝试数据级降级。
   * 降级只影响格式保真度，不改变文件只读和安全渲染原则。
   */
  async function parseWorkbook(buffer, fileName, typeHint) {
    const type = inferFileType(fileName, typeHint, buffer);
    if (type === "csv") return parseCsv(buffer, fileName);
    if (type === "xls") return parseWithSheetJs(buffer, fileName);

    try {
      return await parseWithExcelJs(buffer, fileName);
    } catch (excelError) {
      try {
        const fallback = parseWithSheetJs(buffer, fileName);
        fallback.warnings.push(`ExcelJS 解析失败，已使用兼容模式：${excelError.message}`);
        return fallback;
      } catch (_sheetError) {
        throw excelError;
      }
    }
  }

  /* ======================================================================== */
  /* 4. 工作簿和工作表状态切换                                               */
  /* ======================================================================== */

  function resetViewState() {
    state.searchText = "";
    state.rawMatches = [];
    state.rawMatchLookup = new Map();
    state.rawMatchIndex = -1;
    state.sort = { column: -1, direction: null };
    state.renderer = null;
    dom.search.value = "";
    updateRawSearchControls();
    dom.viewport.scrollTop = 0;
    dom.viewport.scrollLeft = 0;
  }

  function setWorkbook(workbook, byteLength, sourcePath) {
    state.workbook = workbook;
    state.sheetIndex = 0;
    state.view = "raw";
    resetViewState();

    dom.viewerCard.classList.remove("is-empty");
    dom.empty.hidden = true;
    dom.viewport.hidden = false;
    dom.sheetBar.hidden = false;
    dom.fileName.textContent = workbook.name;
    dom.fileMeta.textContent = [
      formatBytes(byteLength),
      `${workbook.sheets.length} 个工作表`,
      workbook.csvInfo || workbook.type.toUpperCase()
    ].join(" · ");

    for (const button of dom.viewSwitch.querySelectorAll("button")) {
      button.disabled = false;
      button.classList.toggle("is-active", button.dataset.view === "raw");
    }
    renderSheetTabs();
    renderCurrentSheet();
    // 文件成功打开后隐藏初始选择区，把整个可用空间交还给表格。
    dom.sourcePanel.hidden = true;
    setSourceMessage("");
    // 只有解析成功后才更新地址，避免损坏文件或 404 地址污染可刷新的页面链接。
    replaceStartupFilePath(sourcePath);
  }

  /**
   * 更新初始文件选择区中的提示信息。
   * 选择区出现时状态栏被隐藏，因此输入校验和加载错误需要在这里就近反馈。
   */
  function setSourceMessage(message) {
    const text = String(message || "").trim();
    dom.sourceMessage.textContent = text;
    dom.sourceMessage.hidden = !text;
  }

  /**
   * 回到“尚未打开文件”的完整空状态。
   * 这里会主动取消下载、作废尚未完成的解析任务，并清除旧工作簿的 DOM 与筛选状态，
   * 确保用户点击“更换文件”后不会在后台又被旧文件切回查看界面。
   */
  function enterFileSelection(message, preserveUrl) {
    state.loadSequence += 1;
    if (state.fetchController) {
      state.fetchController.abort();
      state.fetchController = null;
    }
    if (state.renderFrame) {
      cancelAnimationFrame(state.renderFrame);
      state.renderFrame = 0;
    }
    clearTimeout(state.searchTimer);

    state.workbook = null;
    state.sheetIndex = 0;
    state.view = "raw";
    state.searchText = "";
    state.rawMatches = [];
    state.rawMatchLookup = new Map();
    state.rawMatchIndex = -1;
    state.sort = { column: -1, direction: null };
    state.renderer = null;

    dom.search.value = "";
    dom.search.disabled = true;
    updateRawSearchControls();
    dom.viewport.scrollTop = 0;
    dom.viewport.scrollLeft = 0;
    dom.viewport.hidden = true;
    dom.sheetBar.hidden = true;
    dom.empty.hidden = true;
    dom.header.replaceChildren();
    dom.body.replaceChildren();
    dom.sheetTabs.replaceChildren();
    hideLoading();
    hideCopyToast();

    for (const button of dom.viewSwitch.querySelectorAll("button")) {
      button.disabled = true;
      button.classList.toggle("is-active", button.dataset.view === "raw");
    }

    dom.fileName.textContent = "尚未打开文件";
    dom.fileMeta.textContent = "请选择本地文件或加载在线文件";
    dom.viewerCard.classList.add("is-empty");
    dom.sourcePanel.hidden = false;
    dom.sourcePanel.scrollTop = 0;
    setSourceMessage(message);
    setStatus("准备就绪");
    // 用户主动更换文件或加载失败时清除旧路径；首次初始化需要暂时保留启动参数。
    if (!preserveUrl) replaceStartupFilePath("");
  }

  /**
   * 从初始选择区进入加载状态，并返回本次任务的序号。
   * 后续异步步骤必须核对该序号，以免较慢的旧任务覆盖用户的新选择。
   */
  function beginFileLoad(fileName) {
    const sequence = ++state.loadSequence;
    dom.sourcePanel.hidden = true;
    setSourceMessage("");
    dom.viewerCard.classList.remove("is-empty");
    dom.empty.hidden = true;
    dom.viewport.hidden = true;
    dom.sheetBar.hidden = true;
    dom.fileName.textContent = fileName || "正在打开文件";
    dom.fileMeta.textContent = "正在读取文件内容…";
    return sequence;
  }

  function renderSheetTabs() {
    dom.sheetTabs.replaceChildren();
    state.workbook.sheets.forEach((sheet, index) => {
      const button = document.createElement("button");
      button.type = "button";
      button.className = `sheet-tab${index === state.sheetIndex ? " is-active" : ""}`;
      button.role = "tab";
      button.setAttribute("aria-selected", String(index === state.sheetIndex));
      button.textContent = sheet.name;
      button.title = sheet.name;
      button.addEventListener("click", () => {
        if (state.sheetIndex === index) return;
        state.sheetIndex = index;
        resetViewState();
        renderSheetTabs();
        renderCurrentSheet();
      });
      dom.sheetTabs.appendChild(button);
    });
  }

  function setView(view) {
    if (!state.workbook || !["raw", "data"].includes(view) || state.view === view) return;
    state.view = view;
    state.searchText = "";
    state.rawMatches = [];
    state.rawMatchLookup = new Map();
    state.rawMatchIndex = -1;
    state.sort = { column: -1, direction: null };
    dom.search.value = "";
    updateRawSearchControls();
    dom.viewport.scrollTop = 0;
    for (const button of dom.viewSwitch.querySelectorAll("button")) {
      button.classList.toggle("is-active", button.dataset.view === view);
    }
    renderCurrentSheet();
  }

  function renderCurrentSheet() {
    const sheet = getCurrentSheet();
    if (!sheet) return;
    cancelAnimationFrame(state.renderFrame);
    state.renderFrame = 0;
    state.renderer = null;
    dom.viewport.querySelectorAll(".raw-table").forEach((element) => element.remove());
    dom.header.replaceChildren();
    dom.body.replaceChildren();
    dom.header.hidden = false;
    dom.body.hidden = false;
    // 两种视图都支持搜索：数据视图过滤行，原始视图只高亮匹配单元格。
    dom.search.disabled = false;

    if (state.view === "raw") {
      rebuildRawSearchMatches(sheet);
      renderRawView(sheet);
    } else {
      state.rawMatches = [];
      state.rawMatchLookup = new Map();
      state.rawMatchIndex = -1;
      updateRawSearchControls();
      renderDataView(sheet);
    }
  }

  /* ======================================================================== */
  /* 5. 原始视图、数据视图与虚拟滚动                                         */
  /* ======================================================================== */

  function gridTemplate(sheet, columns) {
    return [
      `${ROW_NUMBER_WIDTH}px`,
      ...columns.map((columnIndex) => `${sheet.colWidths[columnIndex].width}px`)
    ].join(" ");
  }

  function visibleColumnIndices(sheet) {
    const result = [];
    for (let index = 0; index < sheet.maxCols; index += 1) {
      if (!sheet.colWidths[index].hidden) result.push(index);
    }
    return result.length ? result : [0];
  }

  function buildMergeLookup(sheet) {
    if (sheet.mergeLookup) return sheet.mergeLookup;
    const lookup = new Map();
    for (const range of sheet.merges) {
      for (let row = range.startRow; row <= range.endRow; row += 1) {
        for (let column = range.startCol; column <= range.endCol; column += 1) {
          lookup.set(`${row}:${column}`, {
            range,
            master: row === range.startRow && column === range.startCol
          });
        }
      }
    }
    sheet.mergeLookup = lookup;
    return lookup;
  }

  /** 使用源行列序号生成稳定键，虚拟行反复销毁和创建后仍能识别同一匹配项。 */
  function rawMatchKey(rowIndex, columnIndex) {
    return `${rowIndex}:${columnIndex}`;
  }

  /**
   * 扫描原始视图中实际可见的单元格并建立匹配索引。
   * 隐藏行列和合并区域的非主单元格不会显示，因此也不计入匹配数量。
   */
  function rebuildRawSearchMatches(sheet) {
    const query = state.searchText.trim().toLocaleLowerCase("zh-CN");
    const previous = state.rawMatches[state.rawMatchIndex];
    const previousKey = previous ? rawMatchKey(previous.rowIndex, previous.columnIndex) : "";
    const matches = [];
    const lookup = new Map();

    if (query) {
      const columns = visibleColumnIndices(sheet);
      const mergeLookup = buildMergeLookup(sheet);
      for (const row of sheet.rows) {
        if (row.hidden) continue;
        for (const columnIndex of columns) {
          const merge = mergeLookup.get(rawMatchKey(row.sourceIndex, columnIndex));
          if (merge && !merge.master) continue;
          const cell = row.cells[columnIndex];
          if (!cell || !cell.text.toLocaleLowerCase("zh-CN").includes(query)) continue;
          const match = { rowIndex: row.sourceIndex, columnIndex };
          lookup.set(rawMatchKey(match.rowIndex, match.columnIndex), matches.length);
          matches.push(match);
        }
      }
    }

    state.rawMatches = matches;
    state.rawMatchLookup = lookup;
    const preservedIndex = previousKey ? lookup.get(previousKey) : undefined;
    state.rawMatchIndex = Number.isInteger(preservedIndex) ? preservedIndex : matches.length ? 0 : -1;
    updateRawSearchControls();
  }

  /** 根据当前原始视图匹配状态更新计数器与前后导航按钮。 */
  function updateRawSearchControls() {
    const searchingRaw = Boolean(state.workbook && state.view === "raw" && state.searchText.trim());
    const total = state.rawMatches.length;
    dom.searchCount.hidden = !searchingRaw;
    dom.searchCount.textContent = total && state.rawMatchIndex >= 0
      ? `${state.rawMatchIndex + 1}/${total}`
      : "0/0";
    dom.searchPrev.disabled = !searchingRaw || total === 0;
    dom.searchNext.disabled = !searchingRaw || total === 0;
  }

  /** 为已渲染的原始视图单元格附加匹配和当前匹配样式。 */
  function applyRawSearchState(element, rowIndex, columnIndex) {
    const matchIndex = state.rawMatchLookup.get(rawMatchKey(rowIndex, columnIndex));
    if (!Number.isInteger(matchIndex)) return;
    element.classList.add("is-search-match");
    if (matchIndex === state.rawMatchIndex) element.classList.add("is-search-current");
  }

  function renderRawView(sheet) {
    const rows = sheet.rows.filter((row) => !row.hidden);
    const virtual = rows.length > config.virtualizationThreshold;
    if (!virtual) {
      renderRawTable(sheet);
      setSheetStatus(sheet, false, rows.length);
      return;
    }

    const columns = visibleColumnIndices(sheet);
    renderRawHeader(sheet, columns);
    const renderRows = rows.map((row) => ({
      sourceRow: row,
      sourceIndex: row.sourceIndex,
      height: sheet.rowHeights[row.sourceIndex] || DEFAULT_RAW_ROW_HEIGHT
    }));
    setupVirtualRenderer({
      kind: "raw",
      sheet,
      columns,
      rows: renderRows,
      heights: renderRows.map((row) => row.height),
      template: gridTemplate(sheet, columns)
    });
    setSheetStatus(sheet, true, rows.length);
  }

  /** 小表使用标准 table，浏览器可以原生、可靠地处理 rowspan 与 colspan。 */
  function renderRawTable(sheet) {
    dom.header.hidden = true;
    dom.body.hidden = true;
    const table = document.createElement("table");
    table.className = "raw-table";
    const colgroup = document.createElement("colgroup");
    const numberColumn = document.createElement("col");
    numberColumn.style.width = `${ROW_NUMBER_WIDTH}px`;
    colgroup.appendChild(numberColumn);
    for (let columnIndex = 0; columnIndex < sheet.maxCols; columnIndex += 1) {
      const column = document.createElement("col");
      const metadata = sheet.colWidths[columnIndex];
      column.style.width = `${metadata.width}px`;
      if (metadata.hidden) column.style.display = "none";
      colgroup.appendChild(column);
    }
    table.appendChild(colgroup);

    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");
    headerRow.appendChild(document.createElement("th"));
    for (let columnIndex = 0; columnIndex < sheet.maxCols; columnIndex += 1) {
      const header = document.createElement("th");
      header.textContent = columnLetter(columnIndex);
      if (sheet.colWidths[columnIndex].hidden) header.style.display = "none";
      headerRow.appendChild(header);
    }
    thead.appendChild(headerRow);
    table.appendChild(thead);

    const mergeLookup = buildMergeLookup(sheet);
    const tbody = document.createElement("tbody");
    for (const row of sheet.rows) {
      const tr = document.createElement("tr");
      if (row.hidden) tr.hidden = true;
      tr.style.height = `${sheet.rowHeights[row.sourceIndex] || DEFAULT_RAW_ROW_HEIGHT}px`;
      const rowHeader = document.createElement("th");
      rowHeader.scope = "row";
      rowHeader.textContent = String(row.sourceIndex + 1);
      tr.appendChild(rowHeader);

      for (let columnIndex = 0; columnIndex < sheet.maxCols; columnIndex += 1) {
        const merge = mergeLookup.get(`${row.sourceIndex}:${columnIndex}`);
        if (merge && !merge.master) continue;
        const td = document.createElement("td");
        const cell = row.cells[columnIndex];
        setCellContent(td, cell);
        applyCellStyle(td, cell && cell.style);
        prepareInteractiveCell(td, cell, row.sourceIndex, columnIndex, true);
        if (sheet.colWidths[columnIndex].hidden) td.style.display = "none";
        if (merge) {
          td.rowSpan = merge.range.endRow - merge.range.startRow + 1;
          td.colSpan = merge.range.endCol - merge.range.startCol + 1;
        }
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }
    table.appendChild(tbody);
    dom.viewport.appendChild(table);
  }

  function renderRawHeader(sheet, columns) {
    dom.header.style.gridTemplateColumns = gridTemplate(sheet, columns);
    dom.header.appendChild(createHeaderCell("", -1, false, true));
    for (const columnIndex of columns) {
      dom.header.appendChild(createHeaderCell(columnLetter(columnIndex), columnIndex, false, false));
    }
  }

  function prepareDataModel(sheet) {
    const headerIndex = sheet.rows.findIndex((row) => !row.hidden && !isRowEmpty(row));
    const resolvedHeaderIndex = headerIndex >= 0 ? headerIndex : 0;
    const sourceHeader = sheet.rows[resolvedHeaderIndex] || { cells: [] };
    const columns = visibleColumnIndices(sheet);
    const usedNames = new Map();
    const headers = columns.map((columnIndex) => {
      const sourceText = sourceHeader.cells[columnIndex] ? sourceHeader.cells[columnIndex].text.trim() : "";
      const base = sourceText || columnLetter(columnIndex);
      const count = (usedNames.get(base) || 0) + 1;
      usedNames.set(base, count);
      return { text: count === 1 ? base : `${base} (${count})`, columnIndex };
    });

    let rows = sheet.rows
      .slice(resolvedHeaderIndex + 1)
      .filter((row) => !row.hidden && !isRowEmpty(row))
      .map((row, stableIndex) => ({ sourceRow: row, sourceIndex: row.sourceIndex, stableIndex }));

    const query = state.searchText.trim().toLocaleLowerCase("zh-CN");
    if (query) {
      rows = rows.filter((entry) => columns.some((columnIndex) => {
        const cell = entry.sourceRow.cells[columnIndex];
        return cell && cell.text.toLocaleLowerCase("zh-CN").includes(query);
      }));
    }

    if (state.sort.column >= 0 && state.sort.direction) {
      const sourceColumn = columns[state.sort.column];
      const direction = state.sort.direction === "asc" ? 1 : -1;
      rows.sort((left, right) => {
        const comparison = compareCells(
          left.sourceRow.cells[sourceColumn],
          right.sourceRow.cells[sourceColumn]
        );
        return comparison === 0
          ? left.stableIndex - right.stableIndex
          : comparison * direction;
      });
    }
    return { headerIndex: resolvedHeaderIndex, headers, columns, rows };
  }

  /** 空值固定排在末尾；其余按数字/日期优先、文本次之进行稳定比较。 */
  function compareCells(leftCell, rightCell) {
    const left = leftCell ? leftCell.raw : "";
    const right = rightCell ? rightCell.raw : "";
    const leftEmpty = left == null || left === "";
    const rightEmpty = right == null || right === "";
    if (leftEmpty && rightEmpty) return 0;
    if (leftEmpty) return 1;
    if (rightEmpty) return -1;
    if (left instanceof Date && right instanceof Date) return left.getTime() - right.getTime();
    if (typeof left === "number" && typeof right === "number") return left - right;
    return String(left).localeCompare(String(right), "zh-CN", { numeric: true, sensitivity: "base" });
  }

  function renderDataView(sheet) {
    const model = prepareDataModel(sheet);
    dom.header.style.gridTemplateColumns = gridTemplate(sheet, model.columns);
    dom.header.appendChild(createHeaderCell("#", -1, false, true));
    model.headers.forEach((header, displayIndex) => {
      const element = createHeaderCell(header.text, displayIndex, true, false);
      element.title = `${header.text}（点击排序）`;
      element.addEventListener("click", () => changeSort(displayIndex));
      dom.header.appendChild(element);
    });

    if (!model.rows.length) {
      dom.body.style.height = "220px";
      const empty = document.createElement("div");
      empty.className = "grid-empty-result";
      empty.textContent = state.searchText ? "没有符合搜索条件的数据" : "该工作表没有数据行";
      dom.body.appendChild(empty);
      setDataStatus(sheet, 0, 0, false);
      return;
    }

    const virtual = model.rows.length > config.virtualizationThreshold;
    if (virtual) {
      setupVirtualRenderer({
        kind: "data",
        sheet,
        columns: model.columns,
        rows: model.rows,
        heights: model.rows.map(() => DATA_ROW_HEIGHT),
        template: gridTemplate(sheet, model.columns)
      });
    } else {
      dom.body.style.height = "auto";
      dom.body.style.width = `${templateWidth(sheet, model.columns)}px`;
      const fragment = document.createDocumentFragment();
      model.rows.forEach((entry) => {
        fragment.appendChild(createGridRow({
          kind: "data",
          sheet,
          columns: model.columns,
          entry,
          template: gridTemplate(sheet, model.columns),
          virtual: false,
          top: 0,
          height: DATA_ROW_HEIGHT
        }));
      });
      dom.body.appendChild(fragment);
    }
    const total = sheet.rows.slice(model.headerIndex + 1).filter((row) => !row.hidden && !isRowEmpty(row)).length;
    setDataStatus(sheet, model.rows.length, total, virtual);
  }

  function createHeaderCell(text, index, sortable, rowNumber) {
    const element = document.createElement(sortable ? "button" : "div");
    if (sortable) element.type = "button";
    element.className = `grid-header-cell${rowNumber ? " is-row-number" : ""}`;
    const label = document.createElement("span");
    label.textContent = text;
    element.appendChild(label);
    if (sortable && state.sort.column === index && state.sort.direction) {
      const mark = document.createElement("span");
      mark.className = "sort-mark";
      mark.textContent = state.sort.direction === "asc" ? "▲" : "▼";
      element.appendChild(mark);
    }
    return element;
  }

  function changeSort(column) {
    if (state.sort.column !== column) state.sort = { column, direction: "asc" };
    else if (state.sort.direction === "asc") state.sort.direction = "desc";
    else if (state.sort.direction === "desc") state.sort = { column: -1, direction: null };
    else state.sort = { column, direction: "asc" };
    dom.viewport.scrollTop = 0;
    renderCurrentSheet();
  }

  function templateWidth(sheet, columns) {
    return ROW_NUMBER_WIDTH + columns.reduce((sum, index) => sum + sheet.colWidths[index].width, 0);
  }

  /**
   * 构建前缀高度数组：prefix[i] 表示第 i 行顶部相对表体的像素位置。
   * 原始视图允许不同行高，因此不能简单地用 scrollTop / 固定行高计算索引。
   */
  function buildHeightPrefix(heights) {
    const prefix = new Array(heights.length + 1);
    prefix[0] = 0;
    for (let index = 0; index < heights.length; index += 1) {
      prefix[index + 1] = prefix[index] + heights[index];
    }
    return prefix;
  }

  /** 在有序前缀数组中二分查找包含给定像素位置的行。 */
  function findRowAtOffset(prefix, offset) {
    let low = 0;
    let high = prefix.length - 1;
    while (low < high) {
      const middle = Math.floor((low + high + 1) / 2);
      if (prefix[middle] <= offset) low = middle;
      else high = middle - 1;
    }
    return Math.min(low, prefix.length - 2);
  }

  function setupVirtualRenderer(model) {
    model.prefix = buildHeightPrefix(model.heights);
    model.totalHeight = model.prefix[model.prefix.length - 1];
    model.lastRange = "";
    state.renderer = model;
    dom.body.style.height = `${model.totalHeight}px`;
    dom.body.style.width = `${templateWidth(model.sheet, model.columns)}px`;
    renderVirtualWindow(true);
  }

  /**
   * 只创建视口附近的行。滚动事件会被 requestAnimationFrame 合并，
   * 同一帧内无论触发多少次 scroll，都只进行一次范围计算和 DOM 替换。
   */
  function renderVirtualWindow(force) {
    const renderer = state.renderer;
    if (!renderer || !renderer.rows.length) return;
    const bodyTop = dom.header.offsetHeight;
    const visibleTop = Math.max(0, dom.viewport.scrollTop - bodyTop);
    const visibleBottom = visibleTop + dom.viewport.clientHeight;
    const firstVisible = findRowAtOffset(renderer.prefix, visibleTop);
    const lastVisible = findRowAtOffset(renderer.prefix, Math.min(visibleBottom, renderer.totalHeight));
    const start = Math.max(0, firstVisible - config.overscanRows);
    const end = Math.min(renderer.rows.length - 1, lastVisible + config.overscanRows);
    const rangeKey = `${start}:${end}`;
    if (!force && rangeKey === renderer.lastRange) return;
    renderer.lastRange = rangeKey;

    const fragment = document.createDocumentFragment();
    for (let index = start; index <= end; index += 1) {
      fragment.appendChild(createGridRow({
        kind: renderer.kind,
        sheet: renderer.sheet,
        columns: renderer.columns,
        entry: renderer.rows[index],
        template: renderer.template,
        virtual: true,
        top: renderer.prefix[index],
        height: renderer.heights[index]
      }));
    }
    dom.body.replaceChildren(fragment);
  }

  function createGridRow(options) {
    const rowElement = document.createElement("div");
    rowElement.className = `grid-row${options.virtual ? " is-virtual" : ""}`;
    rowElement.style.gridTemplateColumns = options.template;
    rowElement.style.height = `${options.height}px`;
    if (options.virtual) rowElement.style.transform = `translateY(${options.top}px)`;

    const rowNumber = document.createElement("div");
    rowNumber.className = "grid-cell is-row-number";
    rowNumber.textContent = String(options.entry.sourceIndex + 1);
    rowElement.appendChild(rowNumber);

    if (options.kind === "raw") appendRawCells(rowElement, options);
    else appendDataCells(rowElement, options);
    return rowElement;
  }

  function appendDataCells(rowElement, options) {
    for (const columnIndex of options.columns) {
      const cellElement = document.createElement("div");
      cellElement.className = "grid-cell";
      const cell = options.entry.sourceRow.cells[columnIndex];
      setCellContent(cellElement, cell);
      prepareInteractiveCell(cellElement, cell, options.entry.sourceIndex, columnIndex, false);
      rowElement.appendChild(cellElement);
    }
  }

  /**
   * 大表的原始视图只保留横向合并。跨行合并区域中的非左上角单元格显示为空，
   * 从而使每个虚拟行互相独立，不会因某行卸载而破坏其他行的布局。
   */
  function appendRawCells(rowElement, options) {
    const lookup = buildMergeLookup(options.sheet);
    const rowIndex = options.entry.sourceIndex;
    for (const columnIndex of options.columns) {
      const merge = lookup.get(`${rowIndex}:${columnIndex}`);
      if (merge && merge.range.endRow === merge.range.startRow && !merge.master) continue;

      const cellElement = document.createElement("div");
      cellElement.className = "grid-cell";
      let cell = options.entry.sourceRow.cells[columnIndex];
      if (merge && merge.range.endRow > merge.range.startRow && !merge.master) cell = null;
      setCellContent(cellElement, cell);
      applyCellStyle(cellElement, cell && cell.style);
      prepareInteractiveCell(cellElement, cell, rowIndex, columnIndex, true);

      if (merge && merge.master && merge.range.endRow === merge.range.startRow) {
        const visibleSpan = options.columns.filter(
          (index) => index >= merge.range.startCol && index <= merge.range.endCol
        ).length;
        if (visibleSpan > 1) cellElement.style.gridColumn = `span ${visibleSpan}`;
      }
      rowElement.appendChild(cellElement);
    }
  }

  function setCellContent(element, cell) {
    const text = cell ? cell.text : "";
    element.textContent = text;
    element.classList.toggle("is-empty", text === "");
    if (cell && cell.formulaMissing) {
      element.classList.add("is-formula-missing");
      element.title = `${text}（工作簿未保存公式计算结果）`;
    } else if (text.length > 40 || text.includes("\n")) {
      element.title = text;
    }
  }

  /**
   * 给数据单元格添加复制所需的信息；原始视图额外应用搜索高亮。
   * 所有内容都写入 dataset 和 textContent，不把工作簿文本解释为 HTML。
   */
  function prepareInteractiveCell(element, cell, rowIndex, columnIndex, searchable) {
    element.dataset.rowIndex = String(rowIndex);
    element.dataset.columnIndex = String(columnIndex);
    element.dataset.copyText = cell ? cell.text : "";
    if (searchable) applyRawSearchState(element, rowIndex, columnIndex);
  }

  /**
   * 将当前匹配项滚动到视口中央。
   * 虚拟滚动场景先根据前缀高度定位源行，再强制挂载目标附近 DOM；
   * 普通表格则依据元素与滚动容器的矩形差值调整滚动位置。
   */
  function revealRawMatch() {
    const match = state.rawMatches[state.rawMatchIndex];
    if (!match || state.view !== "raw") return;

    if (state.renderer && state.renderer.kind === "raw") {
      const renderer = state.renderer;
      const rowPosition = renderer.rows.findIndex((entry) => entry.sourceIndex === match.rowIndex);
      if (rowPosition >= 0) {
        const rowTop = dom.header.offsetHeight + renderer.prefix[rowPosition];
        const rowHeight = renderer.heights[rowPosition];
        dom.viewport.scrollTop = Math.max(0, rowTop - (dom.viewport.clientHeight - rowHeight) / 2);

        const columnPosition = renderer.columns.indexOf(match.columnIndex);
        if (columnPosition >= 0) {
          const cellLeft = ROW_NUMBER_WIDTH + renderer.columns
            .slice(0, columnPosition)
            .reduce((sum, index) => sum + renderer.sheet.colWidths[index].width, 0);
          const cellWidth = renderer.sheet.colWidths[match.columnIndex].width;
          dom.viewport.scrollLeft = Math.max(0, cellLeft - (dom.viewport.clientWidth - cellWidth) / 2);
        }
        renderer.lastRange = "";
        renderVirtualWindow(true);
      }
    }

    dom.viewport.querySelectorAll(".is-search-current").forEach((cell) => {
      cell.classList.remove("is-search-current");
    });
    const selector = `[data-row-index="${match.rowIndex}"][data-column-index="${match.columnIndex}"]`;
    const target = dom.viewport.querySelector(selector);
    if (!target) return;
    target.classList.add("is-search-current");

    if (!state.renderer) {
      const viewportRect = dom.viewport.getBoundingClientRect();
      const targetRect = target.getBoundingClientRect();
      dom.viewport.scrollTop += targetRect.top - viewportRect.top
        - (dom.viewport.clientHeight - targetRect.height) / 2;
      dom.viewport.scrollLeft += targetRect.left - viewportRect.left
        - (dom.viewport.clientWidth - targetRect.width) / 2;
    }
  }

  /** 上一个和下一个采用循环导航，到达末尾后继续回到第一个匹配项。 */
  function stepRawMatch(offset) {
    const total = state.rawMatches.length;
    if (!total || state.view !== "raw") return;
    state.rawMatchIndex = (state.rawMatchIndex + offset + total) % total;
    updateRawSearchControls();
    revealRawMatch();
    setSheetStatus(getCurrentSheet(), Boolean(state.renderer), getVisibleRawRowCount(getCurrentSheet()));
  }

  function getVisibleRawRowCount(sheet) {
    return sheet ? sheet.rows.filter((row) => !row.hidden).length : 0;
  }

  function setSheetStatus(sheet, virtual, rowCount) {
    const parts = [
      `原始视图`,
      `${rowCount.toLocaleString()} 行 × ${sheet.maxCols.toLocaleString()} 列`,
      virtual ? "已启用虚拟滚动" : "完整渲染"
    ];
    if (state.searchText.trim()) {
      const total = state.rawMatches.length;
      const current = total && state.rawMatchIndex >= 0 ? state.rawMatchIndex + 1 : 0;
      parts.push(`匹配 ${total.toLocaleString()} 个单元格（${current}/${total}）`);
    }
    let warning = false;
    if (virtual && sheet.hasVerticalMerges) {
      parts.push("大表中的跨行合并已降级显示");
      warning = true;
    }
    if (state.workbook.warnings.length) {
      parts.push(state.workbook.warnings.join("；"));
      warning = true;
    }
    setStatus(parts.join(" · "), warning ? "warning" : "info");
  }

  function setDataStatus(sheet, visible, total, virtual) {
    const parts = [
      "数据视图",
      state.searchText ? `筛选后 ${visible.toLocaleString()} / ${total.toLocaleString()} 行` : `${visible.toLocaleString()} 行数据`,
      `${sheet.maxCols.toLocaleString()} 列`,
      virtual ? "已启用虚拟滚动" : "完整渲染"
    ];
    if (state.sort.direction) parts.push(`已按第 ${state.sort.column + 1} 列${state.sort.direction === "asc" ? "升序" : "降序"}`);
    setStatus(parts.join(" · "), state.workbook.warnings.length ? "warning" : "info");
  }

  /* ======================================================================== */
  /* 6. 本地文件、远程 URL 和页面事件                                         */
  /* ======================================================================== */

  async function loadArrayBuffer(buffer, fileName, typeHint, byteLength, sequence, sourcePath) {
    showLoading("正在解析工作表", `${fileName} · ${formatBytes(byteLength)}`);
    setStatus("正在读取文件内容…");
    await nextPaint();
    try {
      const workbook = await parseWorkbook(buffer, fileName, typeHint);
      if (sequence !== state.loadSequence) return;
      setWorkbook(workbook, byteLength, sourcePath);
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
      await loadArrayBuffer(buffer, file.name, "", file.size, sequence, "");
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
      await loadArrayBuffer(buffer, responseName, item && item.type, buffer.byteLength, sequence, trimmedUrl);
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
   * id、name、url 缺一不可；type 和 action 只接受已知值，避免异常清单污染状态。
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
        action: String(item.action || "").trim().toLowerCase()
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

  /** 展开或收起自定义下拉，同时同步无障碍状态。 */
  function setPresetMenuOpen(open, focusSelected) {
    const canOpen = Boolean(open && state.presetFiles.length && !dom.presetTrigger.disabled);
    dom.presetFile.classList.toggle("is-open", canOpen);
    dom.presetTrigger.setAttribute("aria-expanded", String(canOpen));
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
    if (!dom.presetFile.contains(event.target)) setPresetMenuOpen(false, false);
  });

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

  // 使用事件委托，虚拟滚动中新创建的单元格不需要逐个重新绑定复制事件。
  dom.viewport.addEventListener("click", (event) => {
    if (!state.copyEnabled) return;
    const cell = event.target.closest(".grid-cell:not(.is-row-number), .raw-table td");
    if (!cell || !dom.viewport.contains(cell)) return;
    copyRenderedCell(cell);
  });

  dom.viewport.addEventListener("scroll", () => {
    if (!state.renderer || state.renderFrame) return;
    state.renderFrame = requestAnimationFrame(() => {
      state.renderFrame = 0;
      renderVirtualWindow(false);
    });
  }, { passive: true });

  window.addEventListener("resize", () => {
    if (!state.renderer || state.renderFrame) return;
    state.renderFrame = requestAnimationFrame(() => {
      state.renderFrame = 0;
      renderVirtualWindow(true);
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

  initializePage();
}());
