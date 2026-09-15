/**
 * 工作簿状态切换、原始/数据视图、搜索、固定行列与虚拟滚动。
 * 此文件由浏览器原生 ES Modules 直接加载，不依赖 npm、打包器或构建脚本。
 */

import {
  DATA_ROW_HEIGHT,
  DEFAULT_RAW_ROW_HEIGHT,
  ROW_NUMBER_WIDTH,
  RAW_FIT_WIDTH_GUARD,
  DEFAULT_PAGE_TITLE,
  config,
  dom,
  state,
  columnLetter,
  replaceStartupFilePath,
  formatBytes,
  setStatus,
  hideLoading,
  hideCopyToast,
  getCurrentSheet,
  isRowEmpty
} from "../core.js";

import {
  applyCellStyle,
  applyRawGridlineState
} from "../table-styles/table-styles.js";

/* ======================================================================== */
/* 4. 工作簿和工作表状态切换                                               */
/* ======================================================================== */

function resetViewState() {
  state.searchText = "";
  state.rawMatches = [];
  state.rawMatchLookup = new Map();
  state.rawMatchIndex = -1;
  state.sort = { column: -1, direction: null };
  state.pinnedRows.clear();
  state.pinnedColumns.clear();
  state.renderer = null;
  dom.search.value = "";
  updateRawSearchControls();
  dom.viewport.scrollTop = 0;
  dom.viewport.scrollLeft = 0;
  dom.viewport.classList.remove("is-raw-fit");
}

function setWorkbook(workbook, byteLength, sourcePath, autoFit) {
  state.workbook = workbook;
  state.sheetIndex = 0;
  state.view = "raw";
  // 本地文件、直接 URL 或未配置 autoFit 的文件默认关闭自适应。
  state.rawFitEnabled = autoFit === true;
  resetViewState();

  dom.viewerCard.classList.remove("is-empty");
  dom.empty.hidden = true;
  dom.gridFrame.hidden = false;
  dom.viewport.hidden = false;
  dom.sheetBar.hidden = false;
  dom.fileName.textContent = workbook.name;
  // 标签页标题直接使用当前打开文件的显示名称，方便同时打开多个文件时快速区分。
  document.title = workbook.name;
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
  state.rawFitEnabled = false;
  state.pinnedRows.clear();
  state.pinnedColumns.clear();
  state.renderer = null;
  dom.viewport.classList.remove("is-raw-fit");

  dom.search.value = "";
  dom.search.disabled = true;
  updateRawSearchControls();
  dom.viewport.scrollTop = 0;
  dom.viewport.scrollLeft = 0;
  dom.gridFrame.hidden = true;
  dom.viewport.hidden = true;
  dom.sheetBar.hidden = true;
  dom.empty.hidden = true;
  dom.header.replaceChildren();
  dom.body.replaceChildren();
  dom.rawAxisCorner.replaceChildren();
  dom.rawColumnAxisTrack.replaceChildren();
  dom.rawColumnPinnedAxis.replaceChildren();
  dom.rawRowAxisTrack.replaceChildren();
  dom.rawRowPinnedAxis.replaceChildren();
  dom.rawAxisLayer.hidden = true;
  dom.sheetTabs.replaceChildren();
  // 初始化流程使用 preserveUrl=true，此时全屏启动层必须持续到配置和默认文件处理完毕。
  if (!preserveUrl) hideLoading();
  hideCopyToast();

  for (const button of dom.viewSwitch.querySelectorAll("button")) {
    button.disabled = true;
    button.classList.toggle("is-active", button.dataset.view === "raw");
  }

  dom.fileName.textContent = "尚未打开文件";
  dom.fileMeta.textContent = "请选择本地文件或加载在线文件";
  document.title = DEFAULT_PAGE_TITLE;
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
  dom.gridFrame.hidden = true;
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
  dom.gridFrame.classList.toggle("is-raw", state.view === "raw");
  dom.rawAxisLayer.hidden = state.view !== "raw";
  // 自适应时禁止横向滚动；切换到数据视图或恢复原宽后立即恢复正常滚动。
  dom.viewport.classList.toggle("is-raw-fit", state.view === "raw" && state.rawFitEnabled);
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

/**
 * 计算原始视图的显示列宽。
 * 开启自适应后只使用一个统一缩放系数，所以各列仍严格保持 Excel 原宽比例；
 * 行号轴已经位于数据滚动区外，当前 viewport 宽度可全部交给数据列。
 */
function rawDisplayWidths(sheet, columns) {
  const originalWidths = columns.map((columnIndex) => sheet.colWidths[columnIndex].width);
  if (!state.rawFitEnabled || state.view !== "raw") return originalWidths;

  const originalTotal = originalWidths.reduce((sum, width) => sum + width, 0);
  const availableWidth = Math.max(
    1,
    dom.viewport.clientWidth - RAW_FIT_WIDTH_GUARD
  );
  if (!originalTotal || !Number.isFinite(availableWidth)) return originalWidths;
  const scale = availableWidth / originalTotal;
  return originalWidths.map((width) => Math.max(1, Math.round(width * scale * 100) / 100));
}

/** 返回当前工作表中仍可见、且已被用户固定的行号，按原始顺序排列。 */
function visiblePinnedRows(sheet) {
  return Array.from(state.pinnedRows)
    .filter((rowIndex) => sheet.rows[rowIndex] && !sheet.rows[rowIndex].hidden)
    .sort((left, right) => left - right);
}

/** 返回当前工作表中仍可见、且已被用户固定的列号，按原始顺序排列。 */
function visiblePinnedColumns(sheet) {
  return Array.from(state.pinnedColumns)
    .filter((columnIndex) => sheet.colWidths[columnIndex] && !sheet.colWidths[columnIndex].hidden)
    .sort((left, right) => left - right);
}

/**
 * 多个固定行从数据视口顶部依次堆叠。
 *
 * 这里先使用 Excel 行高生成一个可立即使用的初始偏移；标准 table 插入 DOM 后，
 * 还会由 syncRawTablePinnedRowOffsets() 按浏览器实际排版高度重新校准。因为字号、
 * 自动换行、内边距或合并单元格都可能让真实行高大于 Excel 记录值，只依赖原始
 * 行高会让后面的固定行覆盖前面的固定行，看起来像被“挤压”在一起。
 */
function pinnedRowTop(sheet, rowIndex) {
  return visiblePinnedRows(sheet)
    .filter((candidate) => candidate < rowIndex)
    .reduce((sum, candidate) => sum + (sheet.rowHeights[candidate] || DEFAULT_RAW_ROW_HEIGHT), 0);
}

/**
 * 标准 table 完成排版后，使用各固定 tr 的真实高度重新计算 sticky top。
 * 虚拟滚动表格的行高由 grid-row 明确锁定，不需要进行这一步；小表则必须以 DOM
 * 实测结果为准，才能兼容大字号、换行、边框以及 rowspan 造成的行高扩张。
 */
function syncRawTablePinnedRowOffsets(table, sheet) {
  const pinnedRows = new Set(visiblePinnedRows(sheet));
  const body = table.tBodies[0];
  if (!body) return [];

  let sourceTop = 0;
  let stickyTop = 0;
  const entries = [];
  for (const rowElement of body.rows) {
    const rowIndex = Number(rowElement.dataset.sourceRowIndex);
    if (rowElement.hidden) continue;
    const renderedRowHeight = rowElement.getBoundingClientRect().height
      || sheet.rowHeights[rowIndex]
      || DEFAULT_RAW_ROW_HEIGHT;
    entries.push({ sourceIndex: rowIndex, top: sourceTop, height: renderedRowHeight });
    sourceTop += renderedRowHeight;
    if (!pinnedRows.has(rowIndex)) continue;

    // 只修改当前 tr 中参与固定的单元格，避免影响普通行和跨列合并结构。
    for (const cellElement of rowElement.cells) {
      if (cellElement.classList.contains("is-pinned-row-cell")) {
        cellElement.style.top = `${stickyTop}px`;
      }
    }

    stickyTop += renderedRowHeight;
  }
  return entries;
}

/** 多个固定列从数据视口左侧依次堆叠，自适应开启时使用缩放后的实时列宽。 */
function pinnedColumnLeft(sheet, columnIndex) {
  const visibleColumns = visibleColumnIndices(sheet);
  const displayWidths = rawDisplayWidths(sheet, visibleColumns);
  const widthByColumn = new Map(
    visibleColumns.map((candidate, index) => [candidate, displayWidths[index]])
  );
  return visiblePinnedColumns(sheet)
    .filter((candidate) => candidate < columnIndex)
    .reduce((sum, candidate) => sum + (widthByColumn.get(candidate) || 0), 0);
}

/** 给行号或列号添加可点击、可键盘操作的固定控制语义。 */
function prepareAxisPinControl(element, axis, index) {
  const pinned = axis === "row" ? state.pinnedRows.has(index) : state.pinnedColumns.has(index);
  const label = axis === "row" ? `第 ${index + 1} 行` : `${columnLetter(index)} 列`;
  element.dataset.pinAxis = axis;
  element.dataset.pinIndex = String(index);
  element.setAttribute("role", "button");
  element.setAttribute("tabindex", "0");
  element.setAttribute("aria-pressed", String(pinned));
  element.setAttribute("aria-label", `${pinned ? "取消固定" : "固定"}${label}`);
  element.title = `${pinned ? "取消固定" : "固定"}${label}（可同时固定多行或多列）`;
  element.classList.toggle("is-pinned-axis", pinned);
}

/**
 * 将单元格放到对应的粘性行/列位置。固定行覆盖普通内容，固定列覆盖普通列；
 * 二者交叉处使用更高层级，避免横纵滚动时被邻近单元格遮挡。
 */
function applyPinnedCellPosition(element, sheet, rowIndex, columnIndex, pinnedOverlay) {
  const pinnedRow = Number.isInteger(rowIndex) && state.pinnedRows.has(rowIndex);
  const pinnedColumn = Number.isInteger(columnIndex) && state.pinnedColumns.has(columnIndex);
  if (!pinnedRow && !pinnedColumn) return;

  if (pinnedRow) {
    element.classList.add("is-pinned-row-cell");
    // 虚拟大表的固定行由独立粘性层定位，层内单元格无需再次设置 top。
    if (!pinnedOverlay) element.style.top = `${pinnedRowTop(sheet, rowIndex)}px`;
  }
  if (pinnedColumn) {
    element.classList.add("is-pinned-column-cell");
    element.style.left = `${pinnedColumnLeft(sheet, columnIndex)}px`;
  }
  element.style.zIndex = pinnedRow && pinnedColumn ? "16" : pinnedRow ? "13" : "11";
}

/** 点击轴标记后重新渲染，并恢复操作前的滚动位置。 */
function toggleRawAxisPin(axis, index) {
  if (!state.workbook || state.view !== "raw" || !Number.isInteger(index) || index < 0) return;
  const collection = axis === "row" ? state.pinnedRows : state.pinnedColumns;
  if (collection.has(index)) collection.delete(index);
  else collection.add(index);

  const scrollTop = dom.viewport.scrollTop;
  const scrollLeft = dom.viewport.scrollLeft;
  renderCurrentSheet();
  dom.viewport.scrollTop = scrollTop;
  dom.viewport.scrollLeft = scrollLeft;
  requestAnimationFrame(() => {
    dom.viewport.scrollTop = scrollTop;
    dom.viewport.scrollLeft = scrollLeft;
  });
}

function gridTemplate(sheet, columns, useRawFit, includeRowNumber = true) {
  const widths = useRawFit
    ? rawDisplayWidths(sheet, columns)
    : columns.map((index) => sheet.colWidths[index].dataWidth || sheet.colWidths[index].width);
  return [
    ...(includeRowNumber ? [`${ROW_NUMBER_WIDTH}px`] : []),
    ...widths.map((width) => `${width}px`)
  ].join(" ");
}

/** 创建左上角自适应按钮；SVG 为页面自身图标，不依赖字体或外部图片。 */
function createRawFitToggle() {
  const button = document.createElement("button");
  button.type = "button";
  button.className = `raw-fit-toggle${state.rawFitEnabled ? " is-active" : ""}`;
  button.setAttribute("aria-pressed", String(state.rawFitEnabled));
  button.setAttribute("aria-label", state.rawFitEnabled ? "关闭表格自适应" : "开启表格自适应");
  button.title = state.rawFitEnabled ? "恢复 Excel 原始列宽" : "按原列宽比例适应可视区域";

  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("viewBox", "0 0 24 24");
  svg.setAttribute("aria-hidden", "true");
  const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
  path.setAttribute("d", "M8 7 3 12l5 5M3 12h18M16 7l5 5-5 5");
  svg.appendChild(path);
  button.appendChild(svg);
  button.addEventListener("click", (event) => {
    event.preventDefault();
    event.stopPropagation();
    state.rawFitEnabled = !state.rawFitEnabled;
    dom.viewport.scrollLeft = 0;
    renderCurrentSheet();
  });
  return button;
}

/** 创建原始视图的行号或列号控件，并复用既有的多选固定语义。 */
function createRawAxisCell(axis, index, text) {
  const cell = document.createElement("div");
  cell.className = `raw-${axis}-axis-cell`;
  cell.textContent = text;
  prepareAxisPinControl(cell, axis, index);
  return cell;
}

/**
 * 绘制独立的顶部列号轴。普通列号轨道只做横向 transform，固定列号则绘制在
 * 不移动的覆盖层中；这样触摸板产生纵向分量时，A/B/C 仍始终贴住顶部。
 */
function renderRawColumnAxis(sheet, columns) {
  const widths = rawDisplayWidths(sheet, columns);
  const totalWidth = widths.reduce((sum, width) => sum + width, 0);
  const widthByColumn = new Map(columns.map((columnIndex, index) => [columnIndex, widths[index]]));
  const leftByColumn = new Map();
  let sourceLeft = 0;

  dom.rawAxisCorner.replaceChildren(createRawFitToggle());
  dom.rawColumnAxisTrack.replaceChildren();
  dom.rawColumnPinnedAxis.replaceChildren();
  dom.rawColumnAxisTrack.style.width = `${totalWidth}px`;
  dom.rawColumnAxisTrack.style.gridTemplateColumns = widths.map((width) => `${width}px`).join(" ");

  for (const columnIndex of columns) {
    leftByColumn.set(columnIndex, sourceLeft);
    const cell = createRawAxisCell("column", columnIndex, columnLetter(columnIndex));
    if (state.pinnedColumns.has(columnIndex)) {
      cell.dataset.pinnedSource = "true";
      cell.dataset.sourceOffset = String(sourceLeft);
      cell.dataset.pinnedOffset = String(pinnedColumnLeft(sheet, columnIndex));
    }
    dom.rawColumnAxisTrack.appendChild(cell);
    sourceLeft += widthByColumn.get(columnIndex) || 0;
  }

  for (const columnIndex of visiblePinnedColumns(sheet)) {
    const pinnedCell = createRawAxisCell("column", columnIndex, columnLetter(columnIndex));
    const pinnedLeft = pinnedColumnLeft(sheet, columnIndex);
    pinnedCell.classList.add("is-pinned-axis-copy");
    pinnedCell.dataset.sourceOffset = String(leftByColumn.get(columnIndex) || 0);
    pinnedCell.dataset.pinnedOffset = String(pinnedLeft);
    pinnedCell.style.left = `${pinnedLeft}px`;
    pinnedCell.style.width = `${widthByColumn.get(columnIndex) || 0}px`;
    pinnedCell.hidden = true;
    dom.rawColumnPinnedAxis.appendChild(pinnedCell);
  }
}

/**
 * 绘制独立的左侧行号轴。entries 使用与数据区完全相同的源位置和真实高度；
 * 普通轨道只同步纵向位移，固定行号副本则留在不移动的覆盖层中。
 */
function renderRawRowAxis(
  sheet,
  entries,
  totalHeight,
  visibleStart = 0,
  visibleEnd = entries.length - 1,
  pinnedAlways = false
) {
  const entryByRow = new Map(entries.map((entry) => [entry.sourceIndex, entry]));
  dom.rawRowAxisTrack.replaceChildren();
  dom.rawRowPinnedAxis.replaceChildren();
  dom.rawRowAxisTrack.style.height = `${totalHeight}px`;

  for (let index = visibleStart; index <= visibleEnd; index += 1) {
    const entry = entries[index];
    if (!entry) continue;
    const cell = createRawAxisCell("row", entry.sourceIndex, String(entry.sourceIndex + 1));
    cell.style.top = `${entry.top}px`;
    cell.style.height = `${entry.height}px`;
    if (state.pinnedRows.has(entry.sourceIndex)) {
      cell.dataset.pinnedSource = "true";
    }
    dom.rawRowAxisTrack.appendChild(cell);
  }

  let pinnedTop = 0;
  for (const rowIndex of visiblePinnedRows(sheet)) {
    const entry = entryByRow.get(rowIndex);
    if (!entry) continue;
    const pinnedCell = createRawAxisCell("row", rowIndex, String(rowIndex + 1));
    pinnedCell.classList.add("is-pinned-axis-copy");
    pinnedCell.dataset.sourceOffset = String(pinnedAlways ? 0 : entry.top);
    pinnedCell.dataset.pinnedOffset = String(pinnedTop);
    pinnedCell.style.top = `${pinnedTop}px`;
    pinnedCell.style.height = `${entry.height}px`;
    pinnedCell.hidden = true;
    dom.rawRowPinnedAxis.appendChild(pinnedCell);

    const sourceCell = dom.rawRowAxisTrack.querySelector(`[data-pin-index="${rowIndex}"]`);
    if (sourceCell) {
      sourceCell.dataset.sourceOffset = String(pinnedAlways ? 0 : entry.top);
      sourceCell.dataset.pinnedOffset = String(pinnedTop);
    }
    pinnedTop += entry.height;
  }
}

/**
 * 将数据区滚动量投影到两个独立坐标轴，并切换已经到达冻结边界的固定轴副本。
 * 这里只写 transform/hidden，不读取布局，可安全地在 requestAnimationFrame 中调用。
 */
function syncRawAxesScroll() {
  if (!state.workbook || state.view !== "raw" || dom.rawAxisLayer.hidden) return;
  const scrollLeft = dom.viewport.scrollLeft;
  const scrollTop = dom.viewport.scrollTop;
  dom.rawColumnAxisTrack.style.transform = `translate3d(${-scrollLeft}px, 0, 0)`;
  dom.rawRowAxisTrack.style.transform = `translate3d(0, ${-scrollTop}px, 0)`;

  for (const source of dom.rawColumnAxisTrack.querySelectorAll('[data-pinned-source="true"]')) {
    const active = scrollLeft + Number(source.dataset.pinnedOffset) >= Number(source.dataset.sourceOffset);
    source.style.visibility = active ? "hidden" : "visible";
  }
  for (const copy of dom.rawColumnPinnedAxis.children) {
    copy.hidden = !(scrollLeft + Number(copy.dataset.pinnedOffset) >= Number(copy.dataset.sourceOffset));
  }
  for (const source of dom.rawRowAxisTrack.querySelectorAll('[data-pinned-source="true"]')) {
    const active = scrollTop + Number(source.dataset.pinnedOffset) >= Number(source.dataset.sourceOffset);
    source.style.visibility = active ? "hidden" : "visible";
  }
  for (const copy of dom.rawRowPinnedAxis.children) {
    copy.hidden = !(scrollTop + Number(copy.dataset.pinnedOffset) >= Number(copy.dataset.sourceOffset));
  }
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
  const columns = visibleColumnIndices(sheet);
  dom.header.hidden = true;
  renderRawColumnAxis(sheet, columns);
  const virtual = rows.length > config.virtualizationThreshold;
  if (!virtual) {
    renderRawTable(sheet, columns);
    setSheetStatus(sheet, false, rows.length);
    return;
  }

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
    template: gridTemplate(sheet, columns, true, false)
  });
  setSheetStatus(sheet, true, rows.length);
}

/** 小表使用标准 table，浏览器可以原生、可靠地处理 rowspan 与 colspan。 */
function renderRawTable(sheet, visibleColumns) {
  dom.header.hidden = true;
  dom.body.hidden = true;
  const table = document.createElement("table");
  table.className = "raw-table";
  // 行号与列号已移出滚动区，table 的宽度只包含真实数据列。
  const displayedWidths = rawDisplayWidths(sheet, visibleColumns);
  const displayedWidthByColumn = new Map(
    visibleColumns.map((columnIndex, index) => [columnIndex, displayedWidths[index]])
  );
  table.style.width = `${displayedWidths.reduce((sum, width) => sum + width, 0)}px`;
  const colgroup = document.createElement("colgroup");
  for (let columnIndex = 0; columnIndex < sheet.maxCols; columnIndex += 1) {
    const column = document.createElement("col");
    const metadata = sheet.colWidths[columnIndex];
    column.style.width = `${metadata.hidden ? 0 : displayedWidthByColumn.get(columnIndex) || metadata.width}px`;
    if (metadata.hidden) column.style.display = "none";
    colgroup.appendChild(column);
  }
  table.appendChild(colgroup);

  const mergeLookup = buildMergeLookup(sheet);
  const tbody = document.createElement("tbody");
  for (const row of sheet.rows) {
    const tr = document.createElement("tr");
    // 保存源行号，供完成 DOM 排版后的固定行高度校准使用。
    tr.dataset.sourceRowIndex = String(row.sourceIndex);
    if (row.hidden) tr.hidden = true;
    tr.style.height = `${sheet.rowHeights[row.sourceIndex] || DEFAULT_RAW_ROW_HEIGHT}px`;

    for (let columnIndex = 0; columnIndex < sheet.maxCols; columnIndex += 1) {
      const merge = mergeLookup.get(`${row.sourceIndex}:${columnIndex}`);
      if (merge && !merge.master) continue;
      const td = document.createElement("td");
      const cell = row.cells[columnIndex];
      setCellContent(td, cell);
      applyCellStyle(td, cell && cell.style, { omitSharedLeadingBorders: true });
      applyRawGridlineState(td, cell && cell.style, sheet);
      prepareInteractiveCell(td, cell, row.sourceIndex, columnIndex, true);
      applyPinnedCellPosition(td, sheet, row.sourceIndex, columnIndex, false);
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
  // table 只有进入文档后才能取得包含字体、换行和边框影响的准确行高。
  const rowEntries = syncRawTablePinnedRowOffsets(table, sheet);
  const totalHeight = rowEntries.reduce((sum, entry) => sum + entry.height, 0);
  renderRawRowAxis(sheet, rowEntries, totalHeight);
  syncRawAxesScroll();
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

function templateWidth(sheet, columns, useRawFit) {
  const widths = useRawFit
    ? rawDisplayWidths(sheet, columns)
    : columns.map((index) => sheet.colWidths[index].dataWidth || sheet.colWidths[index].width);
  return (useRawFit ? 0 : ROW_NUMBER_WIDTH) + widths.reduce((sum, width) => sum + width, 0);
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
  dom.body.style.width = `${templateWidth(model.sheet, model.columns, model.kind === "raw")}px`;
  renderVirtualWindow(true);
}

/**
 * 虚拟滚动只保留视口附近的普通行；固定行可能早已离开渲染窗口，因此单独
 * 建立始终存在的粘性层。各行仍复用 createGridRow，样式、合并降级、搜索
 * 和复制行为与普通虚拟行完全一致。
 */
function createPinnedRowsOverlay(renderer) {
  if (renderer.kind !== "raw" || !state.pinnedRows.size) return null;
  const entries = renderer.rows
    .filter((entry) => state.pinnedRows.has(entry.sourceIndex))
    .sort((left, right) => left.sourceIndex - right.sourceIndex);
  if (!entries.length) return null;

  const overlay = document.createElement("div");
  overlay.className = "raw-pinned-rows";
  overlay.style.width = `${templateWidth(renderer.sheet, renderer.columns, true)}px`;
  for (const entry of entries) {
    const row = createGridRow({
      kind: "raw",
      sheet: renderer.sheet,
      columns: renderer.columns,
      entry,
      template: renderer.template,
      virtual: false,
      pinnedOverlay: true,
      top: 0,
      height: renderer.sheet.rowHeights[entry.sourceIndex] || DEFAULT_RAW_ROW_HEIGHT
    });
    row.classList.add("is-pinned-overlay-row");
    overlay.appendChild(row);
  }
  return overlay;
}

/**
 * 只创建视口附近的行。滚动事件会被 requestAnimationFrame 合并，
 * 同一帧内无论触发多少次 scroll，都只进行一次范围计算和 DOM 替换。
 */
function renderVirtualWindow(force) {
  const renderer = state.renderer;
  if (!renderer || !renderer.rows.length) return;
  // 原始视图的列号轴位于滚动区外，因此虚拟数据从 scrollTop=0 开始。
  const bodyTop = renderer.kind === "raw" ? 0 : dom.header.offsetHeight;
  const visibleTop = Math.max(0, dom.viewport.scrollTop - bodyTop);
  const visibleBottom = visibleTop + dom.viewport.clientHeight;
  const firstVisible = findRowAtOffset(renderer.prefix, visibleTop);
  const lastVisible = findRowAtOffset(renderer.prefix, Math.min(visibleBottom, renderer.totalHeight));
  const start = Math.max(0, firstVisible - config.overscanRows);
  const end = Math.min(renderer.rows.length - 1, lastVisible + config.overscanRows);
  const rangeKey = `${start}:${end}`;
  if (!force && rangeKey === renderer.lastRange) return;
  renderer.lastRange = rangeKey;

  if (renderer.kind === "raw") {
    const axisEntries = renderer.rows.map((entry, index) => ({
      sourceIndex: entry.sourceIndex,
      top: renderer.prefix[index],
      height: renderer.heights[index]
    }));
    renderRawRowAxis(renderer.sheet, axisEntries, renderer.totalHeight, start, end, true);
  }

  const fragment = document.createDocumentFragment();
  const pinnedOverlay = createPinnedRowsOverlay(renderer);
  if (pinnedOverlay) fragment.appendChild(pinnedOverlay);
  for (let index = start; index <= end; index += 1) {
    // 固定行由上方独立层显示；原位置保留在总高度中，但不重复创建同一行。
    if (renderer.kind === "raw" && state.pinnedRows.has(renderer.rows[index].sourceIndex)) continue;
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
  if (renderer.kind === "raw") syncRawAxesScroll();
}

function createGridRow(options) {
  const rowElement = document.createElement("div");
  rowElement.className = `grid-row${options.virtual ? " is-virtual" : ""}${options.kind === "raw" ? " is-raw-row" : ""}`;
  rowElement.style.gridTemplateColumns = options.template;
  rowElement.style.height = `${options.height}px`;
  if (options.virtual) rowElement.style.transform = `translateY(${options.top}px)`;

  // 数据视图仍在表内显示“#”坐标列；原始视图的行号由独立左轴负责。
  if (options.kind === "data") {
    const rowNumber = document.createElement("div");
    rowNumber.className = "grid-cell is-row-number";
    rowNumber.textContent = String(options.entry.sourceIndex + 1);
    rowElement.appendChild(rowNumber);
  }

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
    applyCellStyle(cellElement, cell && cell.style, { omitSharedLeadingBorders: true });
    applyRawGridlineState(cellElement, cell && cell.style, options.sheet);
    prepareInteractiveCell(cellElement, cell, rowIndex, columnIndex, true);
    applyPinnedCellPosition(
      cellElement,
      options.sheet,
      rowIndex,
      columnIndex,
      Boolean(options.pinnedOverlay)
    );

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
        const displayWidths = rawDisplayWidths(renderer.sheet, renderer.columns);
        const cellLeft = ROW_NUMBER_WIDTH + displayWidths
          .slice(0, columnPosition)
          .reduce((sum, width) => sum + width, 0);
        const cellWidth = displayWidths[columnPosition];
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
  if (state.rawFitEnabled) parts.push("已按比例适应宽度");
  if (!sheet.showGridLines) parts.push("已按文件隐藏网格线");
  const pinnedRowCount = visiblePinnedRows(sheet).length;
  const pinnedColumnCount = visiblePinnedColumns(sheet).length;
  if (pinnedRowCount || pinnedColumnCount) {
    const pinnedParts = [];
    if (pinnedRowCount) pinnedParts.push(`${pinnedRowCount} 行`);
    if (pinnedColumnCount) pinnedParts.push(`${pinnedColumnCount} 列`);
    parts.push(`已固定 ${pinnedParts.join("、")}`);
  }
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


export {
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
  // 暴露给 main.js 的事件代理使用；固定状态和重新渲染仍由视图模块集中管理。
  toggleRawAxisPin
};
