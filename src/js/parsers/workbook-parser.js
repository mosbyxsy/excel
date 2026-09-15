/**
 * ExcelJS、SheetJS、CSV 解析以及统一工作簿模型转换。
 * 此文件由浏览器原生 ES Modules 直接加载，不依赖 npm、打包器或构建脚本。
 */

import {
  SAFETY_LIMITS,
  DEFAULT_RAW_ROW_HEIGHT,
  clamp,
  columnLetter,
  fileExtension,
  isRowEmpty,
  pointsToPixels,
  excelWidthToPixels,
  normalizeColumnWidths,
  parseRangeAddress,
  extractThemeColors,
  normalizeCellStyle
} from "../core.js";

import {
  extractOoxmlTableMetadata,
  extractWorksheetTableStyles,
  getTableCellStyle,
  mergeTableAndCellStyle,
  buildCell
} from "../table-styles/table-styles.js";

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
  // ExcelJS 会保留 theme1.xml，但不会替我们把 theme+tint 解析成最终 RGB。
  // 在遍历单元格前建立颜色表，供字体、边框和填充共用。
  const themeColors = extractThemeColors(workbook);
  // 额外读取 XLSX 包中的 table/dataDxf 元数据，补足 ExcelJS 未展开的超级表样式。
  const ooxmlTables = await extractOoxmlTableMetadata(buffer, themeColors);
  const sheets = [];
  const warnings = [];

  workbook.eachSheet((worksheet) => {
    const rowCount = Math.max(worksheet.actualRowCount || 0, worksheet.rowCount || 0);
    const mergeRanges = Array.isArray(worksheet.model && worksheet.model.merges)
      ? worksheet.model.merges.map(parseRangeAddress).filter(Boolean)
      : [];
    const mergeMaxRow = mergeRanges.reduce((max, range) => Math.max(max, range.endRow + 1), 0);
    const mergeMaxCol = mergeRanges.reduce((max, range) => Math.max(max, range.endCol + 1), 0);
    // 超级表的样式存储在区域级 table 模型中，必须先提取，再逐格与 cell.style 合并。
    const tableStyles = extractWorksheetTableStyles(worksheet, themeColors, warnings, ooxmlTables);
    const tableMaxRow = tableStyles.reduce((max, table) => Math.max(max, table.range.endRow + 1), 0);
    const tableMaxCol = tableStyles.reduce((max, table) => Math.max(max, table.range.endCol + 1), 0);
    const maxRows = Math.max(rowCount, mergeMaxRow, tableMaxRow);
    const maxCols = Math.max(
      worksheet.actualColumnCount || 0,
      worksheet.columnCount || 0,
      mergeMaxCol,
      tableMaxCol,
      1
    );
    assertSafeDimensions(maxRows, maxCols);

    const rows = [];
    const rowHeights = [];
    for (let rowIndex = 0; rowIndex < maxRows; rowIndex += 1) {
      const worksheetRow = worksheet.getRow(rowIndex + 1);
      const cells = new Array(maxCols).fill(null);
      for (let columnIndex = 0; columnIndex < maxCols; columnIndex += 1) {
        const cell = worksheetRow.getCell(columnIndex + 1);
        const tableStyle = getTableCellStyle(tableStyles, rowIndex, columnIndex);
        // 超级表中的空白格仍可能有表头/条纹背景，不能因没有值而跳过。
        if (cell.value == null && !cell.hasStyle && !tableStyle) continue;
        const value = excelJsCellValue(cell);
        const cellStyle = normalizeCellStyle(cell.style, themeColors);
        cells[columnIndex] = buildCell(
          value.text,
          value.raw,
          mergeTableAndCellStyle(tableStyle, cellStyle),
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
      // ExcelJS 将 <sheetView showGridLines="0"> 暴露在 worksheet.views 中。
      // 属性省略是 OOXML 默认值 true；只在文件明确写入 false 时隐藏。
      showGridLines: !(
        Array.isArray(worksheet.views)
        && worksheet.views.some((view) => view && view.showGridLines === false)
      ),
      source: "exceljs",
      tableStyleCount: tableStyles.length
    }));
  });

  if (!sheets.length) throw new Error("工作簿中没有可显示的工作表。");
  return { name: fileName, type: fileExtension(fileName) || "xlsx", sheets, warnings };
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
    // 旧版 XLS 经 SheetJS 数据级解析时通常没有稳定的视图元数据，
    // 按 Excel 默认行为显示网格线；显式单元格填充仍会遮蔽网格。
    showGridLines: true,
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
    // CSV 不包含工作表视图信息，使用电子表格的默认网格线展示。
    showGridLines: true,
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
  sheet.showGridLines = sheet.showGridLines !== false;
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


export {
  parseWorkbook
};
