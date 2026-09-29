/**
 * Excel 超级表规则、OOXML 差异样式、单元格样式合并以及 60 项回归校验。
 * 此文件由浏览器原生 ES Modules 直接加载，不依赖 npm、打包器或构建脚本。
 */

import {
  clamp,
  parseRangeAddress,
  DEFAULT_THEME_COLORS,
  normalizeHexColor,
  applyExcelTint,
  colorToCss,
  normalizeCellStyle
} from "../core.js";

import {
  BUILT_IN_TABLE_STYLE_PRESETS,
  TABLE_STYLE_ELEMENT_ORDER
} from "./built-in-table-style-presets.js";

/* ======================================================================== */
/* Excel“超级表”（Table）样式                                                */
/* ======================================================================== */

/**
 * Excel 的内置超级表样式不会逐格写入 fill/font，而只在 table XML 中保存
 * TableStyleLight/Medium/Dark + 编号。ExcelJS 会保留这段表模型，却不会把
 * 它自动合并到 cell.style；因此查看器需要根据表区域和样式选项自行展开。
 *
 * 60 项精确区域定义已拆到 built-in-table-style-presets.js；这里负责把其中的
 * theme+tint 引用解析成当前工作簿颜色，并按规范顺序组合到具体单元格。
 */
function resolvePresetTableColor(reference, themeColors) {
  if (!reference || !Number.isInteger(reference.theme)) return "";
  const colors = themeColors || DEFAULT_THEME_COLORS;
  const fallback = DEFAULT_THEME_COLORS[reference.theme] || "000000";
  const base = normalizeHexColor(colors[reference.theme] || fallback).slice(-6);
  const resolved = applyExcelTint(base, Number(reference.tint) || 0);
  return resolved ? `#${resolved}` : "";
}

/**
 * Excel 边框粗细转换为浏览器边框。double 至少需要 3px，CSS 才会真正绘制
 * 双线；medium/thick 分别映射为 2px/3px，与原始视图的缩放逻辑保持一致。
 */
function resolvePresetTableBorder(value, themeColors) {
  if (!value) return null;
  const widthByStyle = { thin: 1, medium: 2, thick: 3, double: 3 };
  return {
    width: widthByStyle[value.style] || 1,
    style: value.style === "double" ? "double" : "solid",
    color: resolvePresetTableColor(value.color, themeColors)
  };
}

/** 把一个 tableStyleElement 解析成可直接按区域叠加的浏览器样式。 */
function resolvePresetTableElement(value, themeColors) {
  if (!value) return null;
  const result = {};
  if (value.fill) result.fillColor = resolvePresetTableColor(value.fill, themeColors);
  if (value.font) {
    if (value.font.color) result.fontColor = resolvePresetTableColor(value.font.color, themeColors);
    if (value.font.bold === true) result.bold = true;
  }
  if (value.borders) {
    result.borders = {};
    for (const side of ["top", "right", "bottom", "left", "vertical", "horizontal"]) {
      if (value.borders[side]) {
        result.borders[side] = resolvePresetTableBorder(value.borders[side], themeColors);
      }
    }
  }
  return result;
}

/**
 * 把精确的 60 项预设展开为当前工作簿主题色。palette 保留“区域元素”结构，
 * 渲染时才能正确处理行列条纹交叉、首末列填充和汇总行的继承关系。
 */
function createBuiltInTablePalette(styleName, themeColors) {
  const preset = BUILT_IN_TABLE_STYLE_PRESETS[String(styleName || "").toLowerCase()];
  if (!preset) return null;
  const elements = {};
  for (const type of TABLE_STYLE_ELEMENT_ORDER) {
    if (preset.elements[type]) {
      elements[type] = resolvePresetTableElement(preset.elements[type], themeColors);
    }
  }
  return {
    family: preset.family,
    styleName: preset.name,
    elements,
    rowStripeSize: preset.rowStripeSize || 1,
    columnStripeSize: preset.columnStripeSize || 1
  };
}

/** 自定义表仅有 dxf 时使用的无色基础方案，避免意外混入 Medium2 的规则。 */
function createEmptyTablePalette() {
  return {
    family: "custom",
    styleName: "",
    elements: {},
    rowStripeSize: 1,
    columnStripeSize: 1
  };
}

/** 在带命名空间的 OOXML 中按 localName 读取直接子节点。 */
function xmlDirectChild(node, localName) {
  return Array.from(node && node.children ? node.children : [])
    .find((child) => child.localName === localName) || null;
}

/** 将 OOXML 颜色节点转成 colorToCss 可识别的 ExcelJS 风格对象。 */
function ooxmlColorModel(colorNode) {
  if (!colorNode) return null;
  const result = {};
  const rgb = colorNode.getAttribute("rgb");
  const theme = colorNode.getAttribute("theme");
  const indexed = colorNode.getAttribute("indexed");
  const tint = colorNode.getAttribute("tint");
  if (rgb) result.argb = rgb;
  if (theme != null && theme !== "") result.theme = Number(theme);
  if (indexed != null && indexed !== "") result.indexed = Number(indexed);
  if (tint != null && tint !== "") result.tint = Number(tint);
  return Object.keys(result).length ? result : null;
}

/** OOXML 布尔属性既可能省略 val，也可能使用 0/1 或 true/false。 */
function ooxmlBoolean(node) {
  if (!node) return false;
  const value = node.getAttribute("val");
  return value == null || value === "1" || value === "true";
}

/**
 * 解析 styles.xml 中的 dxf（差异样式）。Excel 会用它记录超级表列的显示
 * 线索，例如用户示例中的 theme + tint 背景；ExcelJS 读取工作簿时不会把
 * dataDxfId 展开到单元格，因此必须从原始 XML 补读。
 */
function parseOoxmlDxfStyle(dxfNode, themeColors) {
  if (!dxfNode) return null;
  const source = {};
  const fillNode = xmlDirectChild(dxfNode, "fill");
  const patternNode = xmlDirectChild(fillNode, "patternFill");
  if (patternNode) {
    source.fill = {
      pattern: patternNode.getAttribute("patternType") || "",
      fgColor: ooxmlColorModel(xmlDirectChild(patternNode, "fgColor")),
      bgColor: ooxmlColorModel(xmlDirectChild(patternNode, "bgColor"))
    };
  }

  const fontNode = xmlDirectChild(dxfNode, "font");
  if (fontNode) {
    const sizeNode = xmlDirectChild(fontNode, "sz");
    const nameNode = xmlDirectChild(fontNode, "name");
    const underlineNode = xmlDirectChild(fontNode, "u");
    const underlineValue = underlineNode && underlineNode.getAttribute("val");
    source.font = {
      name: nameNode ? nameNode.getAttribute("val") || "" : "",
      size: sizeNode ? Number(sizeNode.getAttribute("val")) : null,
      bold: ooxmlBoolean(xmlDirectChild(fontNode, "b")),
      italic: ooxmlBoolean(xmlDirectChild(fontNode, "i")),
      // <u val="none"/> 明确表示无下划线，不能仅凭节点存在就判定为开启。
      underline: Boolean(underlineNode) && underlineValue !== "none" && underlineValue !== "0",
      strike: ooxmlBoolean(xmlDirectChild(fontNode, "strike")),
      color: ooxmlColorModel(xmlDirectChild(fontNode, "color"))
    };
  }

  const alignmentNode = xmlDirectChild(dxfNode, "alignment");
  if (alignmentNode) {
    source.alignment = {
      horizontal: alignmentNode.getAttribute("horizontal") || "",
      vertical: alignmentNode.getAttribute("vertical") || "",
      wrapText: alignmentNode.getAttribute("wrapText") === "1",
      textRotation: Number(alignmentNode.getAttribute("textRotation")) || 0
    };
  }

  const borderNode = xmlDirectChild(dxfNode, "border");
  if (borderNode) {
    source.border = {};
    for (const side of ["top", "right", "bottom", "left"]) {
      const sideNode = xmlDirectChild(borderNode, side);
      const style = sideNode && sideNode.getAttribute("style");
      if (!style) continue;
      source.border[side] = {
        style,
        color: ooxmlColorModel(xmlDirectChild(sideNode, "color"))
      };
    }
  }

  return normalizeCellStyle(source, themeColors);
}

/** 读取一个可选 dxf 编号，非法或越界编号安全地返回 null。 */
function dxfStyleByAttribute(node, attributeName, dxfStyles) {
  const rawId = node && node.getAttribute(attributeName);
  if (rawId == null || rawId === "") return null;
  const id = Number(rawId);
  return Number.isInteger(id) && id >= 0 ? dxfStyles[id] || null : null;
}

/**
 * 从 XLSX ZIP 包直接读取 table*.xml 和 styles.xml。此步骤只解析格式元数据，
 * 不执行宏、外部链接或公式。若 JSZip CDN 不可用或元数据损坏，则返回空映射，
 * 主流程仍使用 ExcelJS 表模型和内置样式回退，不影响数据读取。
 */
async function extractOoxmlTableMetadata(buffer, themeColors) {
  const result = new Map();
  if (typeof window.JSZip === "undefined") return result;

  try {
    const zip = await window.JSZip.loadAsync(buffer);
    const parser = new DOMParser();
    const stylesEntry = zip.file("xl/styles.xml");
    const dxfStyles = [];
    if (stylesEntry) {
      const stylesXml = await stylesEntry.async("text");
      const stylesDocument = parser.parseFromString(stylesXml, "application/xml");
      const dxfsNode = Array.from(stylesDocument.getElementsByTagName("*"))
        .find((node) => node.localName === "dxfs");
      if (dxfsNode) {
        for (const dxfNode of Array.from(dxfsNode.children).filter((node) => node.localName === "dxf")) {
          dxfStyles.push(parseOoxmlDxfStyle(dxfNode, themeColors));
        }
      }
    }

    const tablePaths = Object.keys(zip.files)
      .filter((path) => /^xl\/tables\/[^/]+\.xml$/i.test(path) && !zip.files[path].dir);
    for (const path of tablePaths) {
      const tableXml = await zip.files[path].async("text");
      const tableDocument = parser.parseFromString(tableXml, "application/xml");
      if (tableDocument.getElementsByTagName("parsererror").length) continue;
      const tableNode = tableDocument.documentElement;
      if (!tableNode || tableNode.localName !== "table") continue;

      const styleNode = Array.from(tableDocument.getElementsByTagName("*"))
        .find((node) => node.localName === "tableStyleInfo");
      const columnNodes = Array.from(tableDocument.getElementsByTagName("*"))
        .filter((node) => node.localName === "tableColumn");
      const name = tableNode.getAttribute("name") || tableNode.getAttribute("displayName") || path;
      result.set(name, {
        name,
        displayName: tableNode.getAttribute("displayName") || name,
        range: parseRangeAddress(tableNode.getAttribute("ref")),
        styleName: styleNode ? styleNode.getAttribute("name") || "" : "",
        headerRow: tableNode.getAttribute("headerRowCount") !== "0",
        totalsRow: tableNode.getAttribute("totalsRowCount") === "1",
        showFirstColumn: styleNode ? styleNode.getAttribute("showFirstColumn") === "1" : false,
        showLastColumn: styleNode ? styleNode.getAttribute("showLastColumn") === "1" : false,
        showRowStripes: styleNode ? styleNode.getAttribute("showRowStripes") !== "0" : true,
        showColumnStripes: styleNode ? styleNode.getAttribute("showColumnStripes") === "1" : false,
        headerStyle: dxfStyleByAttribute(tableNode, "headerRowDxfId", dxfStyles),
        dataStyle: dxfStyleByAttribute(tableNode, "dataDxfId", dxfStyles),
        totalsStyle: dxfStyleByAttribute(tableNode, "totalsRowDxfId", dxfStyles),
        columnStyles: columnNodes.map((node) => dxfStyleByAttribute(node, "dataDxfId", dxfStyles))
      });
    }
  } catch (_metadataError) {
    // 元数据增强失败不能导致整个工作簿回退到低保真的 SheetJS 解析。
    return new Map();
  }
  return result;
}

/**
 * ExcelJS 4.4.0 在读取外部文件后有时把省略的 headerRowCount 当成 false。
 * 标准超级表默认包含表头，所以再用表第一行与 column.name 做一次可靠推断。
 */
function inferTableHeaderRow(table, worksheet, range) {
  if (table.headerRow === true) return true;
  const columns = Array.isArray(table.columns) ? table.columns : [];
  if (!columns.length || columns.length !== range.endCol - range.startCol + 1) {
    return table.headerRow !== false;
  }
  return columns.every((column, offset) => {
    const cell = worksheet.getCell(range.startRow + 1, range.startCol + offset + 1);
    const cellText = cell && cell.text != null
      ? String(cell.text)
      : formatDisplayValue(cell && cell.value, cell && cell.numFmt);
    return cellText.trim() === String(column && column.name != null ? column.name : "").trim();
  });
}

/**
 * 从工作表提取超级表定义。不同来源的 ExcelJS 工作表可能暴露 model.tables
 * 或 getTables()，两种形式都兼容。自定义表样式没有可展开的完整规则，采用
 * 当前主题的兼容配色，并在状态警告中明确说明。
 */
function extractWorksheetTableStyles(worksheet, themeColors, warnings, ooxmlTables) {
  let sourceTables = [];
  if (worksheet && worksheet.model && Array.isArray(worksheet.model.tables)) {
    sourceTables = worksheet.model.tables;
  } else if (worksheet && typeof worksheet.getTables === "function") {
    sourceTables = worksheet.getTables();
  }

  return sourceTables.map((entry, index) => {
    const table = entry && (entry.model || entry.table || entry);
    const tableName = table && (table.name || table.displayName);
    const metadata = tableName && ooxmlTables instanceof Map
      ? ooxmlTables.get(tableName) || null
      : null;
    const range = (metadata && metadata.range)
      || parseRangeAddress(table && (table.tableRef || table.ref));
    if (!table || !range) return null;

    const style = table.style && typeof table.style === "object" ? table.style : {};
    const styleName = (metadata && metadata.styleName) || style.theme || style.name || "";
    let palette = createBuiltInTablePalette(styleName, themeColors);
    const hasExactDifferentialStyle = Boolean(metadata && (
      metadata.headerStyle
      || metadata.dataStyle
      || metadata.totalsStyle
      || (metadata.columnStyles || []).some(Boolean)
    ));
    if (!palette && hasExactDifferentialStyle) {
      // 自定义样式没有内置名称可供推导时，只创建一个无色基础方案；
      // 后续按 header/data/total/column 的真实 dxf 分区逐格应用。
      palette = createEmptyTablePalette();
    } else if (!palette && styleName) {
      // 文件没有提供可读取的 dxf 时才使用通用兼容色；不能让兼容色优先于
      // 工作簿自身携带的差异样式，否则会再次出现整表颜色被替换的问题。
      palette = createBuiltInTablePalette("TableStyleMedium2", themeColors);
      warnings.push(`工作表“${worksheet.name}”中的自定义超级表样式“${styleName}”已按兼容样式显示。`);
    }

    return {
      name: tableName || `Table ${index + 1}`,
      range,
      palette,
      headerRow: metadata ? metadata.headerRow : inferTableHeaderRow(table, worksheet, range),
      totalsRow: metadata ? metadata.totalsRow : Boolean(table.totalsRow),
      showFirstColumn: metadata ? metadata.showFirstColumn : Boolean(style.showFirstColumn),
      showLastColumn: metadata ? metadata.showLastColumn : Boolean(style.showLastColumn),
      showRowStripes: metadata ? metadata.showRowStripes : style.showRowStripes !== false,
      showColumnStripes: metadata ? metadata.showColumnStripes : Boolean(style.showColumnStripes),
      headerStyle: metadata && metadata.headerStyle,
      dataStyle: metadata && metadata.dataStyle,
      totalsStyle: metadata && metadata.totalsStyle,
      columnStyles: metadata ? metadata.columnStyles || [] : []
    };
  }).filter((table) => table && table.palette);
}

/**
 * 返回指定偏移所在的条纹区域。第一、第二条纹可以拥有不同宽度；内置样式
 * 当前均为 1，但保留完整算法可兼容未来从自定义 tableStyleElement 读取 size。
 */
function tableStripeRegion(index, start, end, firstSize, secondSize) {
  if (index < start || index > end) return null;
  const safeFirstSize = Math.max(1, Number(firstSize) || 1);
  const safeSecondSize = Math.max(1, Number(secondSize) || 1);
  const cycleSize = safeFirstSize + safeSecondSize;
  const cycleStart = start + Math.floor((index - start) / cycleSize) * cycleSize;
  const secondStart = Math.min(end + 1, cycleStart + safeFirstSize);
  if (index < secondStart) {
    return {
      type: "first",
      start: cycleStart,
      end: Math.min(end, secondStart - 1)
    };
  }
  return {
    type: "second",
    start: secondStart,
    end: Math.min(end, secondStart + safeSecondSize - 1)
  };
}

/** 建立空白边框容器；内部还会使用 vertical/horizontal，最终只输出四个物理边。 */
function emptyPhysicalBorders() {
  return { top: null, right: null, bottom: null, left: null };
}

/**
 * 将一个区域元素叠加到单元格。区域边框必须按范围解释：left/right/top/bottom
 * 只作用于区域外沿，vertical/horizontal 则作用于区域内部共享边界。
 */
function applyTableElementToCell(result, elementStyle, region, rowIndex, columnIndex) {
  if (!elementStyle || !region) return;
  if (elementStyle.fillColor) result.fillColor = elementStyle.fillColor;
  if (elementStyle.fontColor) result.fontColor = elementStyle.fontColor;
  if (elementStyle.bold === true) result.bold = true;

  const elementBorders = elementStyle.borders;
  if (!elementBorders) return;
  if (!result.borders) result.borders = emptyPhysicalBorders();
  const setBorder = (side, value, keepLeadingBorder) => {
    if (!value) return;
    result.borders[side] = value;
    if (keepLeadingBorder && side === "top") result.keepTopBorder = true;
    if (keepLeadingBorder && side === "left") result.keepLeftBorder = true;
  };

  if (rowIndex === region.startRow) setBorder("top", elementBorders.top, true);
  if (rowIndex === region.endRow) setBorder("bottom", elementBorders.bottom, false);
  if (columnIndex === region.startCol) setBorder("left", elementBorders.left, true);
  if (columnIndex === region.endCol) setBorder("right", elementBorders.right, false);
  if (columnIndex < region.endCol) setBorder("right", elementBorders.vertical, false);
  if (rowIndex < region.endRow) setBorder("bottom", elementBorders.horizontal, false);
}

/**
 * 根据 OOXML 的规范顺序计算单元格最终超级表样式：
 * wholeTable → 列条纹 → 行条纹 → 末列 → 首列 → 表头 → 汇总行。
 * 后应用的元素只覆盖自己显式声明的属性，因此行列条纹交叉处不会误清空颜色，
 * 首末列也能正确得到 Excel 预设中的填充、字体和分隔边框，而不只是加粗。
 */
function getTableCellStyle(tables, rowIndex, columnIndex) {
  const table = tables.find((candidate) => (
    rowIndex >= candidate.range.startRow && rowIndex <= candidate.range.endRow
    && columnIndex >= candidate.range.startCol && columnIndex <= candidate.range.endCol
  ));
  if (!table) return null;

  const { range, palette } = table;
  const elements = palette.elements || {};
  const result = {
    fillColor: "",
    fontColor: "",
    bold: false,
    keepTopBorder: false,
    keepLeftBorder: false,
    borders: null
  };
  const wholeRange = { ...range };

  /*
   * 镶边行和镶边列只属于超级表的数据正文，不包含表头和汇总行。
   * 之前把 range.endRow 直接作为条纹终点，当汇总行前的数据行数量恰好落在
   * 第一条纹周期时，未声明独立填充的 totalRow 会错误继承条纹背景；同时
   * 列条纹也会让汇总行出现交替底色。先算出真正的数据区，后续两种条纹
   * 共用这段范围，才能与 Excel 的 DataBodyRange 表现一致。
   */
  const bodyStart = range.startRow + (table.headerRow ? 1 : 0);
  const bodyEnd = range.endRow - (table.totalsRow ? 1 : 0);
  const hasBody = bodyStart <= bodyEnd;

  // 1. wholeTable 为全部后续区域提供基础字体、填充和外框/内部网格。
  applyTableElementToCell(result, elements.wholeTable, wholeRange, rowIndex, columnIndex);

  // 2. 列条纹按数据列循环，但只绘制在数据正文的行范围内。
  if (table.showColumnStripes && hasBody && rowIndex >= bodyStart && rowIndex <= bodyEnd) {
    const stripe = tableStripeRegion(
      columnIndex,
      range.startCol,
      range.endCol,
      palette.columnStripeSize,
      palette.columnStripeSize
    );
    if (stripe) {
      const type = stripe.type === "first" ? "firstColumnStripe" : "secondColumnStripe";
      applyTableElementToCell(result, elements[type], {
        startRow: bodyStart,
        endRow: bodyEnd,
        startCol: stripe.start,
        endCol: stripe.end
      }, rowIndex, columnIndex);
    }
  }

  // 3. 行条纹只在数据正文内循环；规范顺序晚于列条纹，所以显式属性优先。
  if (table.showRowStripes && hasBody && rowIndex >= bodyStart && rowIndex <= bodyEnd) {
    const stripe = tableStripeRegion(
      rowIndex,
      bodyStart,
      bodyEnd,
      palette.rowStripeSize,
      palette.rowStripeSize
    );
    if (stripe) {
      const type = stripe.type === "first" ? "firstRowStripe" : "secondRowStripe";
      applyTableElementToCell(result, elements[type], {
        startRow: stripe.start,
        endRow: stripe.end,
        startCol: range.startCol,
        endCol: range.endCol
      }, rowIndex, columnIndex);
    }
  }

  // 4~5. Excel 规定末列先于首列；虽然二者不会在正常表格中重叠，仍保持规范顺序。
  if (table.showLastColumn && columnIndex === range.endCol) {
    applyTableElementToCell(result, elements.lastColumn, {
      startRow: range.startRow,
      endRow: range.endRow,
      startCol: range.endCol,
      endCol: range.endCol
    }, rowIndex, columnIndex);
  }
  if (table.showFirstColumn && columnIndex === range.startCol) {
    applyTableElementToCell(result, elements.firstColumn, {
      startRow: range.startRow,
      endRow: range.endRow,
      startCol: range.startCol,
      endCol: range.startCol
    }, rowIndex, columnIndex);
  }

  // 6~7. 表头和汇总行最后应用，确保不会被条纹或首末列样式错误覆盖。
  const isHeader = table.headerRow && rowIndex === range.startRow;
  const isTotal = table.totalsRow && rowIndex === range.endRow;
  if (isHeader) {
    applyTableElementToCell(result, elements.headerRow, {
      startRow: range.startRow,
      endRow: range.startRow,
      startCol: range.startCol,
      endCol: range.endCol
    }, rowIndex, columnIndex);
  }
  if (isTotal) {
    applyTableElementToCell(result, elements.totalRow, {
      startRow: range.endRow,
      endRow: range.endRow,
      startCol: range.startCol,
      endCol: range.endCol
    }, rowIndex, columnIndex);
  }

  const columnStyle = table.columnStyles[columnIndex - range.startCol] || null;
  const differentialStyle = isHeader
    ? table.headerStyle
    : isTotal
      ? table.totalsStyle
      : columnStyle || table.dataStyle;
  if (differentialStyle) {
    if (differentialStyle.fillColor) result.fillColor = differentialStyle.fillColor;
    if (differentialStyle.fontColor) result.fontColor = differentialStyle.fontColor;
    if (differentialStyle.bold) result.bold = true;
    for (const key of [
      "fontName", "fontSize", "italic", "underline", "strike", "horizontal",
      "vertical", "wrapText", "rotation", "fillPattern", "fillPatternColor"
    ]) {
      if (differentialStyle[key]) result[key] = differentialStyle[key];
    }
    if (differentialStyle.borders) {
      if (!result.borders) result.borders = emptyPhysicalBorders();
      for (const side of ["top", "right", "bottom", "left"]) {
        if (differentialStyle.borders[side]) result.borders[side] = differentialStyle.borders[side];
      }
    }
  }
  return result;
}

/**
 * 对精确预设做启动回归。测试同时覆盖“超级表测试.xlsx”使用的六种开关组合：
 * 无条纹、行条纹、行列条纹、首列、末列和汇总行，避免只验证一种默认外观。
 */
function verifyBuiltInTableStyleRegression() {
  const fail = (styleName, message) => {
    throw new Error(`内置超级表样式回归失败（${styleName}）：${message}`);
  };
  const assert = (condition, styleName, message) => {
    if (!condition) fail(styleName, message);
  };
  const isCssColor = (value) => !value || /^#[0-9A-F]{6}$/i.test(value);
  const families = [["Light", 21], ["Medium", 28], ["Dark", 11]];
  const makeTable = (styleName, options) => ({
    range: { startRow: 0, endRow: 4, startCol: 0, endCol: 2 },
    palette: createBuiltInTablePalette(styleName, DEFAULT_THEME_COLORS),
    headerRow: true,
    totalsRow: false,
    showFirstColumn: false,
    showLastColumn: false,
    showRowStripes: true,
    showColumnStripes: false,
    headerStyle: null,
    dataStyle: null,
    totalsStyle: null,
    columnStyles: [],
    ...(options || {})
  });
  const at = (table, row, column) => getTableCellStyle([table], row, column);
  const borderAt = (style, side) => style && style.borders && style.borders[side];
  const expectColor = (actual, expected, styleName, message) => {
    assert(actual === expected, styleName, `${message}应为 ${expected}，实际为 ${actual || "无"}`);
  };
  const expectBorder = (style, side, expected, styleName, message) => {
    const actual = borderAt(style, side);
    assert(Boolean(actual), styleName, `${message}缺失`);
    if (expected.width) assert(actual.width === expected.width, styleName, `${message}粗细错误`);
    if (expected.style) assert(actual.style === expected.style, styleName, `${message}线型错误`);
    if (expected.color) expectColor(actual.color, expected.color, styleName, `${message}颜色`);
  };

  let checkedStyles = 0;
  for (const [family, count] of families) {
    for (let number = 1; number <= count; number += 1) {
      const styleName = `TableStyle${family}${number}`;
      const palette = createBuiltInTablePalette(styleName, DEFAULT_THEME_COLORS);
      assert(Boolean(palette), styleName, "无法建立样式方案");
      assert(palette.styleName === styleName, styleName, "名称映射错误");
      assert(Boolean(palette.elements.wholeTable), styleName, "缺少 wholeTable 元素");

      for (const elementStyle of Object.values(palette.elements)) {
        if (elementStyle.fillColor) assert(isCssColor(elementStyle.fillColor), styleName, "填充色无效");
        if (elementStyle.fontColor) assert(isCssColor(elementStyle.fontColor), styleName, "字体色无效");
        for (const value of Object.values(elementStyle.borders || {})) {
          assert(value.width >= 1 && Boolean(value.style), styleName, "边框描述不完整");
          assert(isCssColor(value.color), styleName, "边框颜色无效");
        }
      }

      // 六种表选项同时启用，遍历全部单元格，保证区域组合不会产生非法样式。
      const combined = makeTable(styleName, {
        totalsRow: true,
        showFirstColumn: true,
        showLastColumn: true,
        showRowStripes: true,
        showColumnStripes: true
      });
      for (let row = 0; row <= 4; row += 1) {
        for (let column = 0; column <= 2; column += 1) {
          const value = at(combined, row, column);
          assert(Boolean(value), styleName, "区域组合未生成单元格样式");
          assert(isCssColor(value.fillColor), styleName, "组合后的填充色无效");
          assert(isCssColor(value.fontColor), styleName, "组合后的字体色无效");
        }
      }

      // 关闭全部可选强调后，表体必须只剩 wholeTable 的基础定义。
      const plain = makeTable(styleName, {
        showRowStripes: false,
        showColumnStripes: false,
        showFirstColumn: false,
        showLastColumn: false
      });
      const wholeFill = palette.elements.wholeTable.fillColor || "";
      expectColor(at(plain, 2, 1).fillColor, wholeFill, styleName, "基础表体填充");
      checkedStyles += 1;
    }
  }
  assert(checkedStyles === 60, "全部样式", `只完成了 ${checkedStyles} 项回归`);

  // Light 8~14 的行线/列线由对应条纹开关控制，不能在关闭条纹时凭空出现。
  const light9Plain = makeTable("TableStyleLight9", { showRowStripes: false });
  expectColor(at(light9Plain, 0, 1).fillColor, "#4472C4", "TableStyleLight9", "表头填充");
  expectColor(at(light9Plain, 2, 1).fillColor, "", "TableStyleLight9", "无条纹表体填充");
  assert(!borderAt(at(light9Plain, 2, 1), "bottom"), "TableStyleLight9", "关闭行条纹后仍有内部横线");
  const light9Striped = makeTable("TableStyleLight9", { showColumnStripes: true });
  expectBorder(at(light9Striped, 2, 1), "top", { width: 1, color: "#4472C4" }, "TableStyleLight9", "行条纹顶边");
  expectBorder(at(light9Striped, 2, 1), "left", { width: 1, color: "#4472C4" }, "TableStyleLight9", "列条纹左边");

  const light16 = makeTable("TableStyleLight16", { totalsRow: true });
  expectColor(at(light16, 0, 1).fillColor, "", "TableStyleLight16", "表头填充");
  expectColor(at(light16, 1, 1).fillColor, "#D9E1F2", "TableStyleLight16", "第一行条纹");
  expectBorder(at(light16, 0, 1), "bottom", { width: 2, color: "#4472C4" }, "TableStyleLight16", "表头中边框");
  expectBorder(at(light16, 4, 1), "top", { width: 3, style: "double", color: "#4472C4" }, "TableStyleLight16", "汇总行双边框");

  // Medium 8~14 在关闭条纹时必须显示浅 80% 基础色，而不是浅 60% 条纹色。
  const medium9Plain = makeTable("TableStyleMedium9", { showRowStripes: false });
  const medium9 = makeTable("TableStyleMedium9", { totalsRow: true, showFirstColumn: true });
  expectColor(at(medium9Plain, 2, 1).fillColor, "#D9E1F2", "TableStyleMedium9", "基础表体填充");
  expectColor(at(medium9, 1, 1).fillColor, "#B4C6E7", "TableStyleMedium9", "第一行条纹");
  expectColor(at(medium9, 2, 1).fillColor, "#D9E1F2", "TableStyleMedium9", "第二行条纹");
  expectBorder(at(medium9, 0, 1), "bottom", { width: 3, color: "#FFFFFF" }, "TableStyleMedium9", "表头白色粗线");
  expectColor(at(medium9, 2, 0).fillColor, "#4472C4", "TableStyleMedium9", "首列强调填充");
  expectColor(at(medium9, 2, 0).fontColor, "#FFFFFF", "TableStyleMedium9", "首列强调字体");

  // Medium 15~21 的首末列是完整强调块；汇总行只声明双线，不应强制套表头颜色。
  const medium16 = makeTable("TableStyleMedium16", { totalsRow: true, showFirstColumn: true, showLastColumn: true });
  expectColor(at(medium16, 2, 0).fillColor, "#4472C4", "TableStyleMedium16", "首列强调填充");
  expectColor(at(medium16, 2, 2).fillColor, "#4472C4", "TableStyleMedium16", "末列强调填充");
  expectColor(at(medium16, 4, 1).fillColor, "", "TableStyleMedium16", "汇总行继承填充");
  assert(!at(medium16, 4, 1).bold, "TableStyleMedium16", "汇总行被错误强制加粗");
  expectBorder(at(medium16, 4, 1), "top", { width: 3, style: "double", color: "#000000" }, "TableStyleMedium16", "汇总行双边框");

  // Medium 22~28 的表头继承 wholeTable 浅色填充，汇总行使用主题色中边框。
  const medium23 = makeTable("TableStyleMedium23", { totalsRow: true });
  expectColor(at(medium23, 0, 1).fillColor, "#D9E1F2", "TableStyleMedium23", "表头继承填充");
  expectColor(at(medium23, 1, 1).fillColor, "#B4C6E7", "TableStyleMedium23", "第一行条纹");
  expectBorder(at(medium23, 4, 1), "top", { width: 2, color: "#4472C4" }, "TableStyleMedium23", "汇总行中边框");

  // Dark 1~7：汇总行与首末列都有独立填充和白色中边框，表体没有网格线。
  const dark2 = makeTable("TableStyleDark2", { totalsRow: true, showFirstColumn: true });
  expectColor(at(dark2, 0, 1).fillColor, "#000000", "TableStyleDark2", "表头填充");
  expectColor(at(dark2, 1, 1).fillColor, "#305496", "TableStyleDark2", "第一行条纹");
  expectColor(at(dark2, 2, 1).fillColor, "#4472C4", "TableStyleDark2", "第二行条纹");
  expectColor(at(dark2, 4, 1).fillColor, "#203764", "TableStyleDark2", "汇总行填充");
  expectColor(at(dark2, 2, 0).fillColor, "#305496", "TableStyleDark2", "首列强调填充");
  expectBorder(at(dark2, 2, 0), "right", { width: 2, color: "#FFFFFF" }, "TableStyleDark2", "首列分隔线");
  assert(!borderAt(at(dark2, 2, 1), "bottom"), "TableStyleDark2", "深色表体不应出现横向网格线");

  // Dark 8~11 的汇总行继承表体，不可错误复用表头色；条纹是浅 60% 而非浅 80%。
  const dark8 = makeTable("TableStyleDark8", { totalsRow: true });
  expectColor(at(dark8, 1, 1).fillColor, "#A6A6A6", "TableStyleDark8", "第一行条纹");
  expectColor(at(dark8, 2, 1).fillColor, "#D9D9D9", "TableStyleDark8", "第二行条纹");
  expectColor(at(dark8, 4, 1).fillColor, "#D9D9D9", "TableStyleDark8", "汇总行继承填充");
  const dark9 = makeTable("TableStyleDark9", { totalsRow: true });
  expectColor(at(dark9, 0, 1).fillColor, "#ED7D31", "TableStyleDark9", "复合色表头");
  expectColor(at(dark9, 1, 1).fillColor, "#B4C6E7", "TableStyleDark9", "第一行条纹");
  expectColor(at(dark9, 2, 1).fillColor, "#D9E1F2", "TableStyleDark9", "第二行条纹");
  expectColor(at(dark9, 4, 1).fillColor, "#D9E1F2", "TableStyleDark9", "汇总行继承填充");
  expectBorder(at(dark9, 4, 1), "top", { width: 3, style: "double", color: "#000000" }, "TableStyleDark9", "汇总行黑色双线");

  /*
   * “超级表测试.xlsx”中汇总行前有 16 行数据。偶数个数据行会让紧随其后的
   * 行号再次落到第一条纹周期；这里同时开启行列条纹，专门防止条纹背景越过
   * DataBodyRange 污染汇总行。汇总行没有显式填充时只能继承 wholeTable。
   */
  for (const [styleName, expectedFill] of [
    ["TableStyleLight16", ""],
    ["TableStyleMedium23", "#D9E1F2"],
    ["TableStyleDark8", "#D9D9D9"]
  ]) {
    const fixtureLikeTable = makeTable(styleName, {
      range: { startRow: 0, endRow: 17, startCol: 0, endCol: 2 },
      totalsRow: true,
      showRowStripes: true,
      showColumnStripes: true
    });
    expectColor(
      at(fixtureLikeTable, 17, 0).fillColor,
      expectedFill,
      styleName,
      "汇总行不得继承行列条纹背景"
    );
  }

  // 使用非默认主题确认预设始终从工作簿 theme1.xml 取色。
  const customTheme = [
    "FFFDF6", "18212B", "EEE8DD", "34495E", "1F4E78", "A23E48",
    "5B7553", "B8860B", "287D8E", "6E4B8B", "1261A0", "7C365A"
  ];
  expectColor(
    createBuiltInTablePalette("TableStyleLight9", customTheme).elements.headerRow.fillColor,
    "#1F4E78",
    "TableStyleLight9",
    "自定义主题 accent1"
  );

  // 人工构造不同的行列条纹颜色，验证规范要求的“列条纹先、行条纹后”。
  const crossing = makeTable("TableStyleMedium9", { showColumnStripes: true });
  crossing.palette = {
    ...crossing.palette,
    elements: {
      wholeTable: { fillColor: "#FFFFFF" },
      firstColumnStripe: { fillColor: "#4472C4" },
      firstRowStripe: { fillColor: "#ED7D31" }
    }
  };
  expectColor(at(crossing, 1, 0).fillColor, "#ED7D31", "条纹优先级", "行列交叉填充");
}

verifyBuiltInTableStyleRegression();

/**
 * 超级表样式属于区域级默认样式；单元格自身显式设置的字体和填充优先。
 * 仅填补 cell.style 中缺失的字段，可避免覆盖用户手工标记的背景色。
 */
function mergeTableAndCellStyle(tableStyle, cellStyle) {
  if (!tableStyle) return cellStyle;
  const result = cellStyle ? { ...cellStyle } : {
    fontName: "",
    fontSize: null,
    bold: false,
    italic: false,
    underline: false,
    strike: false,
    fontColor: "",
    fillColor: "",
    fillPattern: "",
    fillPatternColor: "",
    horizontal: "",
    vertical: "",
    wrapText: false,
    rotation: 0,
    borders: { top: null, right: null, bottom: null, left: null }
  };
  if (!result.fillColor && tableStyle.fillColor) result.fillColor = tableStyle.fillColor;
  if (!result.fontColor && tableStyle.fontColor) result.fontColor = tableStyle.fontColor;
  if (tableStyle.bold) result.bold = true;
  if (tableStyle.keepTopBorder) result.keepTopBorder = true;
  if (tableStyle.keepLeftBorder) result.keepLeftBorder = true;
  // 差异样式中的对齐、字体和边框也属于超级表视觉的一部分，但仍只填补
  // 普通单元格未声明的字段，确保手工单元格样式拥有最高优先级。
  for (const key of [
    "fontName", "fontSize", "italic", "underline", "strike", "horizontal",
    "vertical", "wrapText", "rotation", "fillPattern", "fillPatternColor"
  ]) {
    if (!result[key] && tableStyle[key]) result[key] = tableStyle[key];
  }
  if (tableStyle.borders) {
    result.borders = result.borders || { top: null, right: null, bottom: null, left: null };
    for (const side of ["top", "right", "bottom", "left"]) {
      if (!result.borders[side] && tableStyle.borders[side]) {
        result.borders[side] = tableStyle.borders[side];
      }
    }
  }
  return result;
}

/**
 * 只通过 element.style 设置经过白名单筛选的属性。
 * 单元格值始终使用 textContent，不能借样式或内容注入 HTML。
 */
function applyCellStyle(element, style, options) {
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
  if (style.fillPattern && style.fillPatternColor) {
    // CSS 没有 Excel 图案填充的直接等价物，使用小尺寸渐变近似常见纹理。
    // solid 填充不经过这里，因此不会影响绝大多数工作簿的精确底色。
    const color = style.fillPatternColor;
    const patternMap = {
      darkHorizontal: `repeating-linear-gradient(0deg, ${color} 0 2px, transparent 2px 4px)`,
      lightHorizontal: `repeating-linear-gradient(0deg, ${color} 0 1px, transparent 1px 5px)`,
      darkVertical: `repeating-linear-gradient(90deg, ${color} 0 2px, transparent 2px 4px)`,
      lightVertical: `repeating-linear-gradient(90deg, ${color} 0 1px, transparent 1px 5px)`,
      darkDown: `repeating-linear-gradient(45deg, ${color} 0 2px, transparent 2px 5px)`,
      lightDown: `repeating-linear-gradient(45deg, ${color} 0 1px, transparent 1px 6px)`,
      darkUp: `repeating-linear-gradient(-45deg, ${color} 0 2px, transparent 2px 5px)`,
      lightUp: `repeating-linear-gradient(-45deg, ${color} 0 1px, transparent 1px 6px)`,
      darkGrid: `repeating-linear-gradient(0deg, ${color} 0 1px, transparent 1px 5px), repeating-linear-gradient(90deg, ${color} 0 1px, transparent 1px 5px)`,
      lightGrid: `repeating-linear-gradient(0deg, ${color} 0 1px, transparent 1px 7px), repeating-linear-gradient(90deg, ${color} 0 1px, transparent 1px 7px)`,
      darkTrellis: `repeating-linear-gradient(45deg, ${color} 0 1px, transparent 1px 6px), repeating-linear-gradient(-45deg, ${color} 0 1px, transparent 1px 6px)`,
      lightTrellis: `repeating-linear-gradient(45deg, ${color} 0 1px, transparent 1px 8px), repeating-linear-gradient(-45deg, ${color} 0 1px, transparent 1px 8px)`
    };
    element.style.backgroundImage = patternMap[style.fillPattern] || "";
  }
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
    /*
     * 原始视图中的相邻单元格通常同时带有“上+下”或“左+右”边框。
     * separate 表格若四边都画，会把共享线叠成双倍宽；而 collapse 又会让
     * 表体边框影响粘性列标题。原始单元格因此统一由前一个单元格的右边框、
     * 上一行的下边框表示共享线，本单元格跳过 top/left。
     */
    if (options && options.omitSharedLeadingBorders) {
      if (side === "top" && !style.keepTopBorder) continue;
      if (side === "left" && !style.keepLeftBorder) continue;
    }
    const border = style.borders && style.borders[side];
    if (border) {
      element.style[`border${side[0].toUpperCase()}${side.slice(1)}`] =
        `${border.width}px ${border.style} ${border.color}`;
    }
  }
}

/**
 * Excel 的工作表网格线不是单元格边框：
 * - 单元格存在填充色时，Excel 不在填充区域上叠加默认网格线；
 * - 单元格显式边框及超级表分隔线由 applyCellStyle 单独绘制；
 * - 只有无填充单元格才根据 Sheet 的 showGridLines 决定是否显示网格。
 *
 * 使用 class 而不是直接写 border，可以让显式内联边框保持最高优先级。
 */
function applyRawGridlineState(element, style, sheet) {
  const hasFill = Boolean(style && (style.fillColor || style.fillPattern));
  element.classList.toggle("has-sheet-gridline", sheet.showGridLines !== false && !hasFill);
}

function buildCell(text, raw, style, options) {
  return {
    text: text == null ? "" : String(text),
    raw: raw == null ? "" : raw,
    style: style || null,
    formulaMissing: Boolean(options && options.formulaMissing)
  };
}


export {
  extractOoxmlTableMetadata,
  extractWorksheetTableStyles,
  getTableCellStyle,
  mergeTableAndCellStyle,
  applyCellStyle,
  applyRawGridlineState,
  buildCell
};
