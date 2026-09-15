/**
 * 共享常量、DOM 引用、状态、通用工具和 Excel 主题色解析。
 * 此文件由浏览器原生 ES Modules 直接加载，不依赖 npm、打包器或构建脚本。
 */

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
const ROW_NUMBER_WIDTH = 42;
const RAW_HEADER_HEIGHT = 38;
// 预留一个物理像素，规避 table 边框与小数列宽舍入造成的 1px 横向溢出。
const RAW_FIT_WIDTH_GUARD = 1;
// Excel 允许的列宽上限约为 255 个字符；换算后通常不足 1800px。
// 这里只对损坏文件设置宽松安全线，不参与正常工作簿的列宽调整。
const MAX_RAW_COLUMN_WIDTH = 4096;
// 数据视图以检索和排序为主，继续使用可读性范围；原始视图不使用这两个限制。
const MIN_COLUMN_WIDTH = 64;
const MAX_COLUMN_WIDTH = 420;
const DEFAULT_PAGE_TITLE = "轻表格 · Excel / CSV 查看器";

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
  gridFrame: document.getElementById("grid-frame"),
  viewport: document.getElementById("grid-viewport"),
  header: document.getElementById("grid-header"),
  body: document.getElementById("grid-body"),
  rawAxisLayer: document.getElementById("raw-axis-layer"),
  rawAxisCorner: document.getElementById("raw-axis-corner"),
  rawColumnAxis: document.getElementById("raw-column-axis"),
  rawColumnAxisTrack: document.getElementById("raw-column-axis-track"),
  rawColumnPinnedAxis: document.getElementById("raw-column-pinned-axis"),
  rawRowAxis: document.getElementById("raw-row-axis"),
  rawRowAxisTrack: document.getElementById("raw-row-axis-track"),
  rawRowPinnedAxis: document.getElementById("raw-row-pinned-axis"),
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
  // 原始视图是否按当前视口宽度等比缩放所有可见列；每次打开文件时由文件配置初始化。
  rawFitEnabled: false,
  // 点击行号/列号可固定多个轴；Set 便于切换状态并避免重复项。
  pinnedRows: new Set(),
  pinnedColumns: new Set(),
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

/**
 * 显示统一的全屏加载动画并更新阶段说明。
 * 页面准备、清单下载、工作簿下载、解析和打开都调用同一入口，避免不同阶段在
 * “全屏轨道动画”和“表格内小转圈”之间跳变。
 */
function setStartupLoading(title, detail) {
  dom.startupLoadingTitle.textContent = title || "正在准备工作表";
  dom.startupLoadingDetail.textContent = detail || "请稍候…";
  dom.startupLoader.removeAttribute("aria-hidden");
  document.documentElement.classList.add("is-startup-loading");
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
  setStartupLoading(title || "正在读取文件", detail || "请稍候…");
}

function hideLoading() {
  finishStartupLoading();
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

/**
 * Excel 列宽以默认字体中“0”字符的宽度为单位，并额外包含约 5px 边距。
 * 这里采用 Calibri/Arial 常见的 7px 最大数字宽度进行 OOXML 近似换算。
 *
 * 原始视图必须保留窄列和宽列的比例，因此不再使用数据视图的 64~420px
 * 可读性限制；仅以 4096px 防御损坏文件中的异常值。
 */
function excelWidthToPixels(width) {
  const value = Number(width);
  if (!Number.isFinite(value) || value <= 0) return 120;
  const maximumDigitWidth = 7;
  const contentPixels = Math.floor(
    ((256 * value + Math.floor(128 / maximumDigitWidth)) / 256) * maximumDigitWidth
  );
  return clamp(contentPixels + 5, 1, MAX_RAW_COLUMN_WIDTH);
}

/**
 * 同时缓存两套宽度：
 * - width：Excel 原始宽度换算结果，供原始视图和自适应比例计算；
 * - dataWidth：限制后的交互宽度，供数据视图使用。
 *
 * 这样原始视图不会因可读性下限放大窄列，自适应模式也能基于真实比例缩放。
 */
function normalizeColumnWidths(widths, count) {
  const result = [];
  for (let index = 0; index < count; index += 1) {
    const column = widths[index] || {};
    const rawWidth = clamp(Number(column.width) || 120, 1, MAX_RAW_COLUMN_WIDTH);
    result.push({
      width: column.hidden ? 0 : rawWidth,
      dataWidth: column.hidden ? 0 : clamp(rawWidth, MIN_COLUMN_WIDTH, MAX_COLUMN_WIDTH),
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

/**
 * Excel 颜色并不总是直接保存为 RGB。实际文件经常保存“主题色编号 + 色调”，
 * 例如示例文件的隔行底色就是 theme=0/theme=3 再叠加 tint。旧实现只识别
 * ARGB/RGB，因此这些单元格会错误地显示成白色。
 *
 * 这里保留一份 Office 默认主题作为容错值。读取 XLSX 时会再从文件自身的
 * theme1.xml 中覆盖这些颜色，所以使用自定义主题的工作簿也能正确显示。
 */
const DEFAULT_THEME_COLORS = [
  "FFFFFF", "000000", "E7E6E6", "44546A",
  "4472C4", "ED7D31", "A5A5A5", "FFC000",
  "5B9BD5", "70AD47", "0563C1", "954F72"
];

/**
 * Excel 97-2003 以及少数兼容软件会使用 0~63 的索引色。
 * 64 表示“自动颜色”，不能当作真实颜色渲染。
 */
const EXCEL_INDEXED_COLORS = [
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
  "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
  "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
  "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
  "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
  "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333"
];

/** 主题色编号并不等于 XML 节点顺序，必须按 OOXML 规定的固定名称映射。 */
const THEME_COLOR_NAMES = [
  "lt1", "dk1", "lt2", "dk2", "accent1", "accent2",
  "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"
];

function normalizeHexColor(value) {
  const hex = String(value || "").replace(/[^0-9a-f]/gi, "").toUpperCase();
  if (hex.length === 3) return hex.split("").map((character) => character + character).join("");
  if (hex.length === 6 || hex.length === 8) return hex;
  return "";
}

/**
 * 将 Excel 的 ARGB 转为 CSS。绝大多数颜色的透明度为 FF；若确实带透明度，
 * 使用 rgba 保留效果，而不是像旧实现一样直接丢弃 alpha。
 */
function hexToCss(value) {
  const hex = normalizeHexColor(value);
  if (hex.length === 6) return `#${hex}`;
  if (hex.length !== 8) return "";
  const alpha = parseInt(hex.slice(0, 2), 16) / 255;
  if (alpha >= 0.999) return `#${hex.slice(2)}`;
  const red = parseInt(hex.slice(2, 4), 16);
  const green = parseInt(hex.slice(4, 6), 16);
  const blue = parseInt(hex.slice(6, 8), 16);
  return `rgba(${red}, ${green}, ${blue}, ${Math.round(alpha * 1000) / 1000})`;
}

/**
 * Excel 沿用了 Win32 RGBToHLS/HLSToRGB 的 240 阶整数色彩空间。
 * 不能直接调用浏览器的浮点 HSL：它会让某些主题色相差 1~2 个 RGB 色阶，
 * 例如 5B9BD5 + 80% tint 应得到 DDEBF7，而不是 DEEBF7。
 */
const EXCEL_HLS_MAX = 240;
const EXCEL_RGB_MAX = 255;

/** JavaScript 的 % 对负数保留负号，这里转换成数学意义上的正模。 */
function positiveModulo(value, divisor) {
  return ((value % divisor) + divisor) % divisor;
}

/** 按 Win32 的整数运算顺序把 RGB 转成 0~240 的 HLS。 */
function rgbToExcelHls(red, green, blue) {
  const brightest = Math.max(red, green, blue);
  const darkest = Math.min(red, green, blue);
  const sum = brightest + darkest;
  const span = brightest - darkest;
  const luminance = Math.trunc(
    (sum * EXCEL_HLS_MAX + EXCEL_RGB_MAX) / (2 * EXCEL_RGB_MAX)
  );
  if (span === 0) return [0, luminance, 0];

  const saturation = luminance <= EXCEL_HLS_MAX / 2
    ? Math.trunc((span * EXCEL_HLS_MAX + Math.trunc(sum / 2)) / sum)
    : Math.trunc(
      (span * EXCEL_HLS_MAX + Math.trunc((2 * EXCEL_RGB_MAX - sum) / 2))
      / (2 * EXCEL_RGB_MAX - sum)
    );
  const distance = (channel) => Math.trunc(
    ((brightest - channel) * (EXCEL_HLS_MAX / 6) + Math.trunc(span / 2)) / span
  );
  let hue;
  if (red === brightest) {
    hue = distance(blue) - distance(green);
  } else if (green === brightest) {
    hue = EXCEL_HLS_MAX / 3 + distance(red) - distance(blue);
  } else {
    hue = 2 * EXCEL_HLS_MAX / 3 + distance(green) - distance(red);
  }
  return [positiveModulo(hue, EXCEL_HLS_MAX), luminance, saturation];
}

/** Win32 HueToRGB 的整数插值，所有除法都保留原算法的截断时机。 */
function excelHueLevel(lower, upper, hueValue) {
  const hue = positiveModulo(hueValue, EXCEL_HLS_MAX);
  const sixth = EXCEL_HLS_MAX / 6;
  if (hue < sixth) {
    return lower + Math.trunc(
      ((upper - lower) * hue + EXCEL_HLS_MAX / 12) / sixth
    );
  }
  if (hue < EXCEL_HLS_MAX / 2) return upper;
  if (hue < 2 * EXCEL_HLS_MAX / 3) {
    return lower + Math.trunc(
      ((upper - lower) * (2 * EXCEL_HLS_MAX / 3 - hue) + EXCEL_HLS_MAX / 12) / sixth
    );
  }
  return lower;
}

/** 按 Win32 的整数 HLS 规则还原 RGB；灰色也走通用分支以保持 Excel 舍入。 */
function excelHlsToRgb(hue, luminance, saturation) {
  const upper = luminance <= EXCEL_HLS_MAX / 2
    ? Math.trunc(
      (luminance * (EXCEL_HLS_MAX + saturation) + EXCEL_HLS_MAX / 2) / EXCEL_HLS_MAX
    )
    : luminance + saturation - Math.trunc(
      (luminance * saturation + EXCEL_HLS_MAX / 2) / EXCEL_HLS_MAX
    );
  const lower = 2 * luminance - upper;
  const channel = (hueOffset) => clamp(Math.trunc(
    (excelHueLevel(lower, upper, hue + hueOffset) * EXCEL_RGB_MAX + EXCEL_HLS_MAX / 2)
    / EXCEL_HLS_MAX
  ), 0, EXCEL_RGB_MAX);
  return [
    channel(EXCEL_HLS_MAX / 3),
    channel(0),
    channel(-EXCEL_HLS_MAX / 3)
  ];
}

/**
 * Excel 的 tint 范围为 -1~1：负值降低亮度，正值向白色提高亮度。
 * 调整后的亮度必须截断而非四舍五入，这是复现 Excel 实际颜色的关键。
 */
function applyExcelTint(hexValue, tintValue) {
  const hex = normalizeHexColor(hexValue).slice(-6);
  const tint = Number(tintValue);
  if (hex.length !== 6 || !Number.isFinite(tint) || tint === 0) return hex;
  const [hue, originalLuminance, saturation] = rgbToExcelHls(
    parseInt(hex.slice(0, 2), 16),
    parseInt(hex.slice(2, 4), 16),
    parseInt(hex.slice(4, 6), 16)
  );
  const normalizedTint = clamp(tint, -1, 1);
  const movedLuminance = normalizedTint < 0
    ? originalLuminance * (1 + normalizedTint)
    : originalLuminance * (1 - normalizedTint) + EXCEL_HLS_MAX * normalizedTint;
  return excelHlsToRgb(hue, Math.trunc(movedLuminance), saturation)
    .map((channel) => channel.toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase();
}

/**
 * 这些值仅用于启动时校验算法，不参与渲染配色。若未来重构破坏 Excel 的
 * 整数舍入规则，页面会在开发阶段立即给出明确错误，而不是静默偏色。
 */
function verifyExcelTintImplementation() {
  const referenceCases = [
    ["4F81BD", 0.8, "DCE6F1"],
    ["4F81BD", 0.4, "95B3D7"],
    ["C0504D", 0.8, "F2DCDB"],
    ["C0504D", 0.4, "DA9694"],
    ["5B9BD5", 0.8, "DDEBF7"],
    ["5B9BD5", 0.4, "9BC2E6"],
    // 非默认主题示例可防止实现只对 Office 默认强调色“碰巧正确”。
    ["DAB6BA", 0.7999, "F8EFF0"],
    ["FFFFFF", -0.15, "D9D9D9"]
  ];
  for (const [source, tint, expected] of referenceCases) {
    if (applyExcelTint(source, tint) !== expected) {
      throw new Error(`Excel 主题色 tint 校验失败：${source} / ${tint}`);
    }
  }
}

verifyExcelTintImplementation();

/**
 * 从 ExcelJS 保留的 theme1.xml 中提取主题颜色。解析失败时返回默认 Office
 * 主题，避免某个非标准主题文件导致整个工作簿无法显示。
 */
function extractThemeColors(workbook) {
  const palette = DEFAULT_THEME_COLORS.slice();
  try {
    const themes = workbook && workbook.model && workbook.model.themes;
    if (!themes || typeof themes !== "object") return palette;
    const themeXml = typeof themes.theme1 === "string"
      ? themes.theme1
      : Object.values(themes).find((value) => typeof value === "string");
    if (!themeXml) return palette;

    const documentNode = new DOMParser().parseFromString(themeXml, "application/xml");
    if (documentNode.getElementsByTagName("parsererror").length) return palette;
    const allNodes = Array.from(documentNode.getElementsByTagName("*"));
    const colorScheme = allNodes.find((node) => node.localName === "clrScheme");
    if (!colorScheme) return palette;

    for (let index = 0; index < THEME_COLOR_NAMES.length; index += 1) {
      const name = THEME_COLOR_NAMES[index];
      const themeNode = Array.from(colorScheme.children).find((node) => node.localName === name);
      if (!themeNode) continue;
      const colorNode = Array.from(themeNode.getElementsByTagName("*"))
        .find((node) => node.localName === "srgbClr" || node.localName === "sysClr");
      if (!colorNode) continue;
      const rawValue = colorNode.getAttribute("val");
      const value = normalizeHexColor(
        rawValue === "window" || rawValue === "windowText"
          ? colorNode.getAttribute("lastClr")
          : rawValue || colorNode.getAttribute("lastClr")
      );
      if (value.length >= 6) palette[index] = value.slice(-6);
    }
  } catch (_themeError) {
    // 主题只影响视觉保真度，解析失败不应阻断工作簿中的数据读取。
  }
  return palette;
}

/** 将 ExcelJS/SheetJS 的 ARGB、RGB、主题色和索引色统一为安全的 CSS 颜色。 */
function colorToCss(color, themeColors) {
  if (!color || typeof color !== "object") return "";
  if (typeof color.argb === "string") return hexToCss(color.argb);
  if (typeof color.rgb === "string") return hexToCss(color.rgb);

  let baseHex = "";
  const themeIndex = Number(color.theme);
  const indexed = Number(color.indexed);
  if (Number.isInteger(themeIndex) && themeIndex >= 0) {
    baseHex = (themeColors || DEFAULT_THEME_COLORS)[themeIndex] || "";
  } else if (Number.isInteger(indexed) && indexed >= 0 && indexed < EXCEL_INDEXED_COLORS.length) {
    baseHex = EXCEL_INDEXED_COLORS[indexed];
  }
  if (!baseHex) return "";
  return `#${applyExcelTint(baseHex, color.tint)}`;
}

function normalizeBorderSide(side, themeColors) {
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
    color: colorToCss(side.color, themeColors) || "#9eaaa6"
  };
}

/** 将两个解析库的样式对象统一为页面内部使用的轻量格式。 */
function normalizeCellStyle(source, themeColors) {
  if (!source || typeof source !== "object") return null;
  const font = source.font || {};
  // ExcelJS 把填充放在 fill 中，SheetJS CE 则可能把 patternType/fgColor
  // 直接放在样式根对象上；两种结构在此合并处理。
  const fill = source.fill || (
    source.patternType || source.pattern || source.fgColor || source.bgColor ? source : {}
  );
  const alignment = source.alignment || {};
  const border = source.border || {};
  const pattern = fill.pattern || fill.patternType || "";
  const foregroundColor = colorToCss(fill.fgColor, themeColors);
  const backgroundColor = colorToCss(fill.bgColor, themeColors);
  // solid 填充以 fgColor 为准；其他图案以 bgColor 为底、fgColor 为纹理色。
  const fillColor = pattern === "solid"
    ? foregroundColor || backgroundColor
    : backgroundColor || foregroundColor;
  const fontColor = colorToCss(font.color, themeColors);
  return {
    fontName: typeof font.name === "string" ? font.name : "",
    fontSize: Number.isFinite(font.size) ? clamp(font.size, 7, 48) : null,
    bold: Boolean(font.bold),
    italic: Boolean(font.italic),
    underline: Boolean(font.underline),
    strike: Boolean(font.strike),
    fontColor,
    fillColor: pattern === "none" || pattern === "gray125" ? "" : fillColor,
    fillPattern: pattern && pattern !== "none" && pattern !== "solid" && pattern !== "gray125" ? pattern : "",
    fillPatternColor: foregroundColor,
    horizontal: alignment.horizontal || "",
    vertical: alignment.vertical || "",
    wrapText: Boolean(alignment.wrapText),
    rotation: Number.isFinite(alignment.textRotation) ? alignment.textRotation : 0,
    borders: {
      top: normalizeBorderSide(border.top, themeColors),
      right: normalizeBorderSide(border.right, themeColors),
      bottom: normalizeBorderSide(border.bottom, themeColors),
      left: normalizeBorderSide(border.left, themeColors)
    }
  };
}


export {
  SAFETY_LIMITS,
  DATA_ROW_HEIGHT,
  DEFAULT_RAW_ROW_HEIGHT,
  ROW_NUMBER_WIDTH,
  RAW_HEADER_HEIGHT,
  RAW_FIT_WIDTH_GUARD,
  DEFAULT_PAGE_TITLE,
  config,
  dom,
  state,
  clamp,
  columnLetter,
  fileExtension,
  nameFromUrl,
  startupRequest,
  replaceStartupFilePath,
  setStartupLoading,
  finishStartupLoading,
  formatBytes,
  setStatus,
  showLoading,
  hideLoading,
  hideCopyToast,
  setCopyEnabled,
  copyRenderedCell,
  nextPaint,
  getCurrentSheet,
  isRowEmpty,
  pointsToPixels,
  excelWidthToPixels,
  normalizeColumnWidths,
  parseRangeAddress,
  DEFAULT_THEME_COLORS,
  normalizeHexColor,
  applyExcelTint,
  extractThemeColors,
  colorToCss,
  normalizeCellStyle
};
