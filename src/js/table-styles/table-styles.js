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

/* ======================================================================== */
/* Excel“超级表”（Table）样式                                                */
/* ======================================================================== */

/**
 * Excel 的内置超级表样式不会逐格写入 fill/font，而只在 table XML 中保存
 * TableStyleLight/Medium/Dark + 编号。ExcelJS 会保留这段表模型，却不会把
 * 它自动合并到 cell.style；因此查看器需要根据表区域和样式选项自行展开。
 *
 * 下面的 7 个颜色槽对应“深色 1 + 强调色 1~6”。第一槽不是固定黑色：
 * 当工作簿替换了主题时，它也会跟随 theme1.xml 中的 dk1 变化。
 */
const TABLE_THEME_SLOTS = Object.freeze([
  { name: "dark1", themeIndex: 1 },
  { name: "accent1", themeIndex: 4 },
  { name: "accent2", themeIndex: 5 },
  { name: "accent3", themeIndex: 6 },
  { name: "accent4", themeIndex: 7 },
  { name: "accent5", themeIndex: 8 },
  { name: "accent6", themeIndex: 9 }
]);

/** 颜色引用可指定另一主题槽，Dark 9~11 等复合色样式会用到。 */
function tableColorReference(slot, tone) {
  return Object.freeze({ slot, tone });
}

/** 用于规则表的边框描述；颜色在读取具体工作簿主题后才解析。 */
function tableBorder(color, width, style) {
  return Object.freeze({
    color: color || "base",
    width: width || 1,
    style: style || "solid"
  });
}

const WHITE_TABLE_DIVIDER = tableBorder("white", 1, "solid");

/**
 * 建立 Excel 桌面端的 60 种内置超级表规则：Light 21 + Medium 28 + Dark 11。
 * 规则保存的是“主题槽 + 色阶 + 区域结构”，而不是最终 RGB，所以同一条规则
 * 能正确适配默认 Office 主题和工作簿自定义主题。
 */
function createBuiltInTableStyleRules() {
  const rules = {};
  const addRule = (family, number, template) => {
    const name = `TableStyle${family}${number}`;
    const colorSlot = (number - 1) % 7;
    rules[name.toLowerCase()] = Object.freeze({
      name,
      family: family.toLowerCase(),
      number,
      colorSlot,
      rowStripeSize: 1,
      columnStripeSize: 1,
      firstColumnBold: true,
      lastColumnBold: true,
      ...template
    });
  };

  // Light 1~7：无表头填充，表头文字使用主题基色的 25% 深色，表体隔行浅填充。
  for (let number = 1; number <= 7; number += 1) {
    addRule("Light", number, {
      headerFill: "",
      headerFont: "dark25",
      bodyFill: "",
      bodyFont: "textDark",
      firstRowStripeFill: number === 1 ? "neutralLight" : "light80",
      secondRowStripeFill: "",
      firstColumnStripeFill: number === 1 ? "neutralLight" : "light80",
      secondColumnStripeFill: "",
      totalFill: "",
      totalFont: "dark25",
      tableTopBorder: tableBorder("base"),
      headerBottomBorder: tableBorder("base"),
      tableBottomBorder: tableBorder("base")
    });
  }

  /**
   * Light 8~14：实色表头、主题色外框和逐行分隔线。
   * 这一组即使启用了 showRowStripes 也没有条纹填充；该开关只表示“允许应用
   * 样式中的条纹元素”，不能反过来凭空生成浅色条纹。Light10 的原生效果
   * 因此是橙色表头、白色表体和橙色细线，而不是浅橙色隔行背景。
   */
  for (let number = 8; number <= 14; number += 1) {
    addRule("Light", number, {
      headerFill: "base",
      headerFont: "textLight",
      // 表体使用 lt1 的显式填充，而不是透明背景；这样与 Excel 一样会遮住
      // 工作表默认竖向网格线，只留下该样式声明的横向行边框。
      bodyFill: "textLight",
      bodyFont: "textDark",
      firstRowStripeFill: "",
      secondRowStripeFill: "",
      firstColumnStripeFill: "",
      secondColumnStripeFill: "",
      totalFill: "textLight",
      totalFont: "textDark",
      tableTopBorder: tableBorder("base"),
      headerBottomBorder: tableBorder("base"),
      rowSeparatorBorder: tableBorder("base"),
      tableBottomBorder: tableBorder("base"),
      outerLeftBorder: tableBorder("base"),
      outerRightBorder: tableBorder("base")
    });
  }

  /**
   * Light 15~21：无表头填充、浅色隔行，并用主题色细线完整勾勒单元格。
   * Excel 样式库中这一组的表头底色是“无填充”，并不是 60% 浅色；此前把
   * headerFill 写成 light60 会让整组样式看起来更像 Medium 22~28。
   */
  for (let number = 15; number <= 21; number += 1) {
    addRule("Light", number, {
      headerFill: "",
      headerFont: "textDark",
      bodyFill: "",
      bodyFont: "textDark",
      firstRowStripeFill: "light80",
      secondRowStripeFill: "",
      firstColumnStripeFill: "light80",
      secondColumnStripeFill: "",
      totalFill: "",
      totalFont: "textDark",
      // 统一把共享边界只落到一个单元格侧面，避免 separate 边框叠成双线。
      tableTopBorder: tableBorder("base"),
      headerBottomBorder: tableBorder("base"),
      rowSeparatorBorder: tableBorder("base"),
      columnSeparatorBorder: tableBorder("base"),
      tableBottomBorder: tableBorder("base"),
      outerLeftBorder: tableBorder("base"),
      outerRightBorder: tableBorder("base")
    });
  }

  // Medium 1~7：实色表头、白色基础行和 80% 浅色隔行。
  for (let number = 1; number <= 7; number += 1) {
    const isNeutralStyle = number === 1;
    addRule("Medium", number, {
      headerFill: "base",
      headerFont: "textLight",
      bodyFill: "",
      bodyFont: "textDark",
      firstRowStripeFill: isNeutralStyle ? "neutralLight" : "light80",
      secondRowStripeFill: "",
      firstColumnStripeFill: isNeutralStyle ? "neutralLight" : "light80",
      secondColumnStripeFill: "",
      totalFill: "base",
      totalFont: "textLight",
      tableTopBorder: tableBorder(isNeutralStyle ? "base" : "light40"),
      headerBottomBorder: tableBorder(isNeutralStyle ? "base" : "light40"),
      rowSeparatorBorder: tableBorder(isNeutralStyle ? "base" : "light40"),
      tableBottomBorder: tableBorder(isNeutralStyle ? "base" : "light40"),
      outerLeftBorder: tableBorder(isNeutralStyle ? "base" : "light40"),
      outerRightBorder: tableBorder(isNeutralStyle ? "base" : "light40")
    });
  }

  // Medium 8~14：两种数据行都着色，分别使用 60% 与 80% 浅色。
  // 例如 Medium14 会由 accent6 得到 #70AD47/#C6E0B4/#E2EFDA。
  for (let number = 8; number <= 14; number += 1) {
    const isNeutralStyle = number === 8;
    addRule("Medium", number, {
      headerFill: "base",
      headerFont: "textLight",
      bodyFill: isNeutralStyle ? "neutralDark" : "light60",
      bodyFont: "textDark",
      firstRowStripeFill: isNeutralStyle ? "neutralDark" : "light60",
      secondRowStripeFill: isNeutralStyle ? "neutralLight" : "light80",
      firstColumnStripeFill: isNeutralStyle ? "neutralDark" : "light60",
      secondColumnStripeFill: isNeutralStyle ? "neutralLight" : "light80",
      totalFill: "base",
      totalFont: "textLight",
      headerBottomBorder: tableBorder("white", 3),
      rowSeparatorBorder: WHITE_TABLE_DIVIDER,
      columnSeparatorBorder: WHITE_TABLE_DIVIDER
    });
  }

  // Medium 15~21：主题色表头配中性灰条纹，并以黑色粗线封闭表头和表尾。
  for (let number = 15; number <= 21; number += 1) {
    const isNeutralStyle = number === 15;
    addRule("Medium", number, {
      headerFill: "base",
      headerFont: "textLight",
      bodyFill: "",
      bodyFont: "textDark",
      firstRowStripeFill: "neutralLight",
      secondRowStripeFill: "",
      firstColumnStripeFill: "neutralLight",
      secondColumnStripeFill: "",
      totalFill: "",
      totalFont: "textDark",
      tableTopBorder: tableBorder("textDark", 2),
      headerBottomBorder: tableBorder("textDark", 2),
      rowSeparatorBorder: isNeutralStyle ? tableBorder("textDark") : null,
      columnSeparatorBorder: isNeutralStyle ? tableBorder("textDark") : null,
      tableBottomBorder: tableBorder("textDark", 2),
      outerLeftBorder: isNeutralStyle ? tableBorder("textDark") : null,
      outerRightBorder: isNeutralStyle ? tableBorder("textDark") : null
    });
  }

  // Medium 22~28：浅色表头配双层浅色表体，并在每个行列边界绘制主题色细线。
  for (let number = 22; number <= 28; number += 1) {
    const isNeutralStyle = number === 22;
    const headerFill = isNeutralStyle ? "neutralLight" : "light80";
    const strongBodyFill = isNeutralStyle ? "neutralDark" : "light60";
    const lightBodyFill = isNeutralStyle ? "neutralLight" : "light80";
    const ruleColor = isNeutralStyle ? "textDark" : "light40";
    addRule("Medium", number, {
      headerFill,
      headerFont: "textDark",
      bodyFill: strongBodyFill,
      bodyFont: "textDark",
      firstRowStripeFill: strongBodyFill,
      secondRowStripeFill: lightBodyFill,
      firstColumnStripeFill: strongBodyFill,
      secondColumnStripeFill: lightBodyFill,
      totalFill: headerFill,
      totalFont: "textDark",
      tableTopBorder: tableBorder(ruleColor),
      headerBottomBorder: tableBorder(ruleColor),
      rowSeparatorBorder: tableBorder(ruleColor),
      columnSeparatorBorder: tableBorder(ruleColor),
      tableBottomBorder: tableBorder(ruleColor),
      outerLeftBorder: tableBorder(ruleColor),
      outerRightBorder: tableBorder(ruleColor)
    });
  }

  /**
   * Dark 1~7：表头统一使用主题“深色 1”，而不是各强调色再压暗 50%。
   * Dark 2~7 的第一条纹是强调色的 25% 深色，第二条纹才是强调色本身；
   * Dark 1 没有可继续压暗的黑色，因此使用黑色向白色提升 25%/45% 得到
   * Excel 默认主题中的 #404040/#737373。
   *
   * Excel 的 Dark 系列表体只依靠深浅填充区分数据行，既没有横向网格线，
   * 也没有纵向网格线。仅表头底部和汇总行顶部需要与表体形成边界。
   * 工作表默认网格线也会被深色填充遮住，因此 bodyBorders 必须完全省略；
   * 否则表体会出现 Excel 原表中不存在的白色横线或竖线。
   */
  for (let number = 1; number <= 7; number += 1) {
    const isNeutralStyle = number === 1;
    addRule("Dark", number, {
      headerFill: "textDark",
      headerFont: "textLight",
      headerBorders: { bottom: WHITE_TABLE_DIVIDER },
      bodyFill: isNeutralStyle ? "light45" : "base",
      bodyFont: "textLight",
      firstRowStripeFill: isNeutralStyle ? "light25" : "dark25",
      secondRowStripeFill: isNeutralStyle ? "light45" : "base",
      firstColumnStripeFill: isNeutralStyle ? "light25" : "dark25",
      secondColumnStripeFill: isNeutralStyle ? "light45" : "base",
      totalFill: "textDark",
      totalFont: "textLight",
      totalBorders: { top: WHITE_TABLE_DIVIDER }
    });
  }

  // Dark 8~11 是 Excel 图库中的复合色样式，不能用编号取模简单推导。
  const darkCompositeRules = [
    {
      number: 8,
      colorSlot: 0,
      headerFill: tableColorReference(0, "base"),
      bodyFill: "neutralDark",
      stripeFill: "neutralLight",
      bodyFont: "textDark"
    },
    {
      number: 9,
      colorSlot: 2,
      headerFill: tableColorReference(2, "base"),
      bodyFill: tableColorReference(1, "light60"),
      stripeFill: tableColorReference(1, "light80"),
      bodyFont: "textDark"
    },
    {
      number: 10,
      colorSlot: 4,
      headerFill: tableColorReference(4, "base"),
      bodyFill: tableColorReference(3, "light60"),
      stripeFill: tableColorReference(3, "light80"),
      bodyFont: "textDark"
    },
    {
      number: 11,
      colorSlot: 6,
      headerFill: tableColorReference(6, "base"),
      bodyFill: tableColorReference(5, "light60"),
      stripeFill: tableColorReference(5, "light80"),
      bodyFont: "textDark"
    }
  ];
  for (const composite of darkCompositeRules) {
    addRule("Dark", composite.number, {
      colorSlot: composite.colorSlot,
      headerFill: composite.headerFill,
      headerFont: "textLight",
      headerBorders: { bottom: WHITE_TABLE_DIVIDER },
      bodyFill: composite.bodyFill,
      bodyFont: composite.bodyFont,
      firstRowStripeFill: composite.bodyFill,
      secondRowStripeFill: composite.stripeFill,
      firstColumnStripeFill: composite.bodyFill,
      secondColumnStripeFill: composite.stripeFill,
      totalFill: composite.headerFill,
      totalFont: "textLight",
      totalBorders: { top: WHITE_TABLE_DIVIDER }
    });
  }

  /**
   * 除了检查总数，还逐项检查合法名称。只检查 60 这个数字无法发现“漏掉一个、
   * 又重复写入另一个”的问题；逐项校验能保证 Excel 的每个内置名称都可命中。
   */
  const expectedFamilies = [
    ["Light", 21],
    ["Medium", 28],
    ["Dark", 11]
  ];
  for (const [family, lastNumber] of expectedFamilies) {
    for (let number = 1; number <= lastNumber; number += 1) {
      const expectedName = `TableStyle${family}${number}`;
      const rule = rules[expectedName.toLowerCase()];
      if (!rule || rule.name !== expectedName) {
        throw new Error(`缺少内置超级表样式规则：${expectedName}`);
      }
    }
  }
  if (Object.keys(rules).length !== 60) {
    throw new Error("内置超级表样式规则必须且只能包含 60 项。");
  }
  return Object.freeze(rules);
}

const BUILT_IN_TABLE_STYLE_RULES = createBuiltInTableStyleRules();

/** 按颜色槽取得当前工作簿的主题基色。 */
function tableThemeBaseColor(themeColors, colorSlot) {
  const safeSlot = clamp(Number(colorSlot) || 0, 0, TABLE_THEME_SLOTS.length - 1);
  const themeSlot = TABLE_THEME_SLOTS[safeSlot];
  return (themeColors || DEFAULT_THEME_COLORS)[themeSlot.themeIndex]
    || DEFAULT_THEME_COLORS[themeSlot.themeIndex];
}

/**
 * 将规则中的色阶引用转换为 CSS 颜色。默认主题和自定义主题都从当前文件
 * 的 theme1.xml 取色，再走同一套 Excel tint 算法，不使用固定配色覆盖文件。
 */
function resolveBuiltInTableColor(reference, themeColors, defaultColorSlot) {
  if (!reference) return "";
  if (typeof reference === "string" && reference.startsWith("#")) return reference;
  if (reference === "white" || reference === "textLight") {
    const light = (themeColors || DEFAULT_THEME_COLORS)[0] || "FFFFFF";
    return `#${normalizeHexColor(light).slice(-6) || "FFFFFF"}`;
  }
  if (reference === "black" || reference === "textDark") {
    const dark = (themeColors || DEFAULT_THEME_COLORS)[1] || "000000";
    return `#${normalizeHexColor(dark).slice(-6) || "000000"}`;
  }
  if (reference === "neutralLight" || reference === "neutralDark") {
    // 内置样式的中性条纹来自 lt1 向黑色压暗 15%/35%，并不是 accent3 灰色。
    const lightBase = normalizeHexColor(
      (themeColors || DEFAULT_THEME_COLORS)[0] || "FFFFFF"
    ).slice(-6);
    return `#${applyExcelTint(lightBase, reference === "neutralLight" ? -0.15 : -0.35)}`;
  }

  const colorSlot = typeof reference === "object" && Number.isInteger(reference.slot)
    ? reference.slot
    : defaultColorSlot;
  const tone = typeof reference === "object" ? reference.tone : reference;
  const baseColor = normalizeHexColor(tableThemeBaseColor(themeColors, colorSlot)).slice(-6);
  const tintByTone = {
    base: 0,
    light20: 0.2,
    light25: 0.25,
    light40: 0.4,
    light45: 0.45,
    light60: 0.6,
    light80: 0.8,
    dark25: -0.25,
    dark50: -0.5
  };
  const tint = tintByTone[tone];
  const color = applyExcelTint(baseColor, Number.isFinite(tint) ? tint : 0);
  return color ? `#${color}` : "";
}

/** 将规则边框中的主题色引用转换为 applyCellStyle 可直接使用的边框。 */
function resolveBuiltInTableBorders(borders, themeColors, defaultColorSlot) {
  if (!borders) return null;
  const result = { top: null, right: null, bottom: null, left: null };
  for (const side of ["top", "right", "bottom", "left"]) {
    const border = borders[side];
    if (!border) continue;
    result[side] = {
      width: border.width || 1,
      style: border.style || "solid",
      color: resolveBuiltInTableColor(border.color || "base", themeColors, defaultColorSlot)
    };
  }
  return result;
}

/** 解析一个用于表格外框或内部边界的单独边框。 */
function resolveBuiltInTableBorder(border, themeColors, defaultColorSlot) {
  if (!border) return null;
  return {
    width: border.width || 1,
    style: border.style || "solid",
    color: resolveBuiltInTableColor(border.color || "base", themeColors, defaultColorSlot)
  };
}

/** 把 60 项规则中的主题引用展开成当前工作簿对应的实际配色方案。 */
function createBuiltInTablePalette(styleName, themeColors) {
  const rule = BUILT_IN_TABLE_STYLE_RULES[String(styleName || "").toLowerCase()];
  if (!rule) return null;
  const color = (reference) => resolveBuiltInTableColor(reference, themeColors, rule.colorSlot);
  const borders = (value) => resolveBuiltInTableBorders(value, themeColors, rule.colorSlot);
  return {
    family: rule.family,
    styleName: rule.name,
    bodyFill: color(rule.bodyFill),
    bodyFont: color(rule.bodyFont),
    bodyBorders: borders(rule.bodyBorders),
    headerFill: color(rule.headerFill),
    headerFont: color(rule.headerFont),
    headerBorders: borders(rule.headerBorders),
    totalFill: color(rule.totalFill),
    totalFont: color(rule.totalFont),
    totalBorders: borders(rule.totalBorders),
    firstRowStripeFill: color(rule.firstRowStripeFill),
    firstRowStripeFont: color(rule.firstRowStripeFont),
    firstRowStripeBorders: borders(rule.firstRowStripeBorders),
    secondRowStripeFill: color(rule.secondRowStripeFill),
    secondRowStripeFont: color(rule.secondRowStripeFont),
    secondRowStripeBorders: borders(rule.secondRowStripeBorders),
    firstColumnStripeFill: color(rule.firstColumnStripeFill),
    firstColumnStripeFont: color(rule.firstColumnStripeFont),
    firstColumnStripeBorders: borders(rule.firstColumnStripeBorders),
    secondColumnStripeFill: color(rule.secondColumnStripeFill),
    secondColumnStripeFont: color(rule.secondColumnStripeFont),
    secondColumnStripeBorders: borders(rule.secondColumnStripeBorders),
    tableTopBorder: resolveBuiltInTableBorder(rule.tableTopBorder, themeColors, rule.colorSlot),
    headerBottomBorder: resolveBuiltInTableBorder(rule.headerBottomBorder, themeColors, rule.colorSlot),
    rowSeparatorBorder: resolveBuiltInTableBorder(rule.rowSeparatorBorder, themeColors, rule.colorSlot),
    columnSeparatorBorder: resolveBuiltInTableBorder(rule.columnSeparatorBorder, themeColors, rule.colorSlot),
    tableBottomBorder: resolveBuiltInTableBorder(rule.tableBottomBorder, themeColors, rule.colorSlot),
    outerLeftBorder: resolveBuiltInTableBorder(rule.outerLeftBorder, themeColors, rule.colorSlot),
    outerRightBorder: resolveBuiltInTableBorder(rule.outerRightBorder, themeColors, rule.colorSlot),
    rowStripeSize: rule.rowStripeSize || 1,
    columnStripeSize: rule.columnStripeSize || 1,
    firstColumnBold: rule.firstColumnBold !== false,
    lastColumnBold: rule.lastColumnBold !== false
  };
}

/** 自定义表仅有 dxf 时使用的无色基础方案，避免意外混入 Medium2 的规则。 */
function createEmptyTablePalette() {
  return {
    family: "custom",
    styleName: "",
    bodyFill: "",
    bodyFont: "",
    bodyBorders: null,
    headerFill: "",
    headerFont: "",
    headerBorders: null,
    totalFill: "",
    totalFont: "",
    totalBorders: null,
    firstRowStripeFill: "",
    firstRowStripeFont: "",
    firstRowStripeBorders: null,
    secondRowStripeFill: "",
    secondRowStripeFont: "",
    secondRowStripeBorders: null,
    firstColumnStripeFill: "",
    firstColumnStripeFont: "",
    firstColumnStripeBorders: null,
    secondColumnStripeFill: "",
    secondColumnStripeFont: "",
    secondColumnStripeBorders: null,
    tableTopBorder: null,
    headerBottomBorder: null,
    rowSeparatorBorder: null,
    columnSeparatorBorder: null,
    tableBottomBorder: null,
    outerLeftBorder: null,
    outerRightBorder: null,
    rowStripeSize: 1,
    columnStripeSize: 1,
    firstColumnBold: true,
    lastColumnBold: true
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

/** 根据行列位置计算某个单元格应继承的超级表视觉样式。 */
function getTableCellStyle(tables, rowIndex, columnIndex) {
  const table = tables.find((candidate) => (
    rowIndex >= candidate.range.startRow && rowIndex <= candidate.range.endRow
    && columnIndex >= candidate.range.startCol && columnIndex <= candidate.range.endCol
  ));
  if (!table) return null;

  const { range, palette } = table;
  const isHeader = table.headerRow && rowIndex === range.startRow;
  const isTotal = table.totalsRow && rowIndex === range.endRow;
  const bodyStart = range.startRow + (table.headerRow ? 1 : 0);
  const bodyEnd = range.endRow - (table.totalsRow ? 1 : 0);
  const isBody = rowIndex >= bodyStart && rowIndex <= bodyEnd;

  let fillColor = "";
  let fontColor = "";
  let borders = null;
  let keepTopBorder = false;
  let keepLeftBorder = false;
  let bold = false;
  if (isHeader) {
    fillColor = palette.headerFill;
    fontColor = palette.headerFont;
    borders = palette.headerBorders;
    bold = true;
  } else if (isTotal) {
    fillColor = palette.totalFill;
    fontColor = palette.totalFont;
    borders = palette.totalBorders;
    bold = true;
  } else if (isBody) {
    fillColor = palette.bodyFill;
    fontColor = palette.bodyFont;
    borders = palette.bodyBorders;
    const rowOffset = rowIndex - bodyStart;
    const columnOffset = columnIndex - range.startCol;

    /**
     * 条纹宽度虽然内置样式目前都是 1，但这里仍按 OOXML 的 size 语义计算，
     * 以后接入自定义 tableStyleElement 时无需重写渲染逻辑。
     */
    const stripePart = (offset, firstSize, secondSize) => {
      const safeFirstSize = Math.max(1, Number(firstSize) || 1);
      const safeSecondSize = Math.max(1, Number(secondSize) || 1);
      return offset % (safeFirstSize + safeSecondSize) < safeFirstSize ? "first" : "second";
    };
    const applyStripe = (kind, axis) => {
      const prefix = `${kind}${axis}Stripe`;
      if (palette[`${prefix}Fill`]) fillColor = palette[`${prefix}Fill`];
      if (palette[`${prefix}Font`]) fontColor = palette[`${prefix}Font`];
      if (palette[`${prefix}Borders`]) borders = palette[`${prefix}Borders`];
    };

    /**
     * OOXML 规定先应用奇/偶行条纹，再应用奇/偶列条纹；后应用的列条纹在
     * 交叉单元格中拥有更高优先级。这里只让条纹实际声明的属性覆盖前一层，
     * 空的 fill/font/border 表示“不覆盖”，不能误清除整表基础样式。
     */
    if (table.showRowStripes) {
      applyStripe(
        stripePart(rowOffset, palette.rowStripeSize, palette.rowStripeSize),
        "Row"
      );
    }
    if (table.showColumnStripes) {
      applyStripe(
        stripePart(columnOffset, palette.columnStripeSize, palette.columnStripeSize),
        "Column"
      );
    }
  }

  /**
   * 外框和内部边界必须根据当前单元格的位置落到正确一侧。若把同一组边框
   * 无条件套给每格，就会在表格右侧/底部多画线，并产生粗细不一的双边框。
   */
  const setBorder = (side, border) => {
    if (!border) return;
    if (!borders) borders = { top: null, right: null, bottom: null, left: null };
    borders = { ...borders, [side]: border };
  };
  if (rowIndex === range.startRow && palette.tableTopBorder) {
    setBorder("top", palette.tableTopBorder);
    keepTopBorder = true;
  }
  if (isHeader) setBorder("bottom", palette.headerBottomBorder);
  if (isBody && rowIndex < bodyEnd) setBorder("bottom", palette.rowSeparatorBorder);
  // 原始视图用前一格的 right 表示共享竖线，避免 separate 边框叠成双线。
  if (columnIndex < range.endCol) setBorder("right", palette.columnSeparatorBorder);
  if (rowIndex === range.endRow) setBorder("bottom", palette.tableBottomBorder);
  if (columnIndex === range.startCol && palette.outerLeftBorder) {
    setBorder("left", palette.outerLeftBorder);
    keepLeftBorder = true;
  }
  if (columnIndex === range.endCol) setBorder("right", palette.outerRightBorder);

  if (table.showFirstColumn && columnIndex === range.startCol && palette.firstColumnBold) {
    bold = true;
  }
  if (table.showLastColumn && columnIndex === range.endCol && palette.lastColumnBold) {
    bold = true;
  }

  const columnStyle = table.columnStyles[columnIndex - range.startCol] || null;
  const differentialStyle = isHeader
    ? table.headerStyle
    : isTotal
      ? table.totalsStyle
      : isBody
        ? columnStyle || table.dataStyle
        : null;
  const result = {
    fillColor,
    fontColor,
    bold,
    keepTopBorder,
    keepLeftBorder,
    // 内置超级表的分隔线也属于样式；单元格自身边框仍会在合并阶段优先。
    borders
  };
  if (differentialStyle) {
    // dxf 是表头、汇总行、整段数据或特定表列的显式差异格式，作用域已经
    // 由 table/column 元数据确定；因此只覆盖其所属区域，不能扩散到整张表。
    if (differentialStyle.fillColor) {
      result.fillColor = differentialStyle.fillColor;
    }
    if (differentialStyle.fontColor) result.fontColor = differentialStyle.fontColor;
    if (differentialStyle.bold) result.bold = true;
    for (const key of [
      "fontName", "fontSize", "italic", "underline", "strike", "horizontal",
      "vertical", "wrapText", "rotation", "fillPattern", "fillPatternColor"
    ]) {
      if (differentialStyle[key]) result[key] = differentialStyle[key];
    }
    if (differentialStyle.borders) {
      // normalizeCellStyle 即使未声明边框也会返回四个 null 槽位；因此不能
      // 整体替换内置样式边框，只让 dxf 中真正存在的边覆盖对应一侧。
      result.borders = result.borders
        ? { ...result.borders }
        : { top: null, right: null, bottom: null, left: null };
      for (const side of ["top", "right", "bottom", "left"]) {
        if (differentialStyle.borders[side]) {
          result.borders[side] = differentialStyle.borders[side];
        }
      }
    }
  }
  return result;
}

/**
 * 启动时对 Excel 的 60 种内置超级表样式做结构化回归。
 *
 * 这不是只检查“规则数量等于 60”：每一项都会实际创建一个 5 行 3 列的
 * 虚拟超级表，再分别读取表头、奇偶数据行、内部分隔线、外框和汇总行。
 * 一旦以后调整配色或边框时破坏了某个样式族，页面会立即抛出包含样式名的
 * 中文错误，而不会悄悄把错误颜色展示给用户。
 */
function verifyBuiltInTableStyleRegression() {
  const fail = (styleName, message) => {
    throw new Error(`内置超级表样式回归失败（${styleName}）：${message}`);
  };
  const assert = (condition, styleName, message) => {
    if (!condition) fail(styleName, message);
  };
  const isCssColor = (value) => !value || /^#[0-9A-F]{6}$/i.test(value)
    || /^rgba\(/i.test(value);
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
  const styleAt = (table, row, column) => getTableCellStyle([table], row, column);
  const borderAt = (style, side) => style && style.borders && style.borders[side];
  const expectBorder = (styleName, style, side, message, width) => {
    const border = borderAt(style, side);
    assert(Boolean(border), styleName, message);
    if (width) assert(border.width === width, styleName, `${message}应为 ${width}px`);
    assert(isCssColor(border.color), styleName, `${message}颜色无效`);
  };
  const expectNoBorder = (styleName, style, side, message) => {
    assert(!borderAt(style, side), styleName, message);
  };

  const families = [
    ["Light", 21],
    ["Medium", 28],
    ["Dark", 11]
  ];
  let checkedStyles = 0;

  for (const [family, lastNumber] of families) {
    for (let number = 1; number <= lastNumber; number += 1) {
      const styleName = `TableStyle${family}${number}`;
      const table = makeTable(styleName);
      const palette = table.palette;
      assert(Boolean(palette), styleName, "无法生成主题色方案");
      assert(palette.styleName === styleName, styleName, "样式名称映射错误");

      // 所有颜色必须已经从“主题槽 + tint”解析为浏览器可安全应用的 CSS 颜色。
      for (const [key, value] of Object.entries(palette)) {
        if ((key.endsWith("Fill") || key.endsWith("Font")) && typeof value === "string") {
          assert(isCssColor(value), styleName, `${key} 不是合法颜色`);
        }
        if (key.endsWith("Border") && value) {
          assert(value.width >= 1 && Boolean(value.style), styleName, `${key} 边框描述不完整`);
          assert(isCssColor(value.color), styleName, `${key} 边框颜色无效`);
        }
        if (key.endsWith("Borders") && value) {
          for (const [side, border] of Object.entries(value)) {
            if (!border) continue;
            assert(border.width >= 1 && Boolean(border.style), styleName, `${key}.${side} 描述不完整`);
            assert(isCssColor(border.color), styleName, `${key}.${side} 颜色无效`);
          }
        }
      }

      const headerLeft = styleAt(table, 0, 0);
      const headerMiddle = styleAt(table, 0, 1);
      const firstBody = styleAt(table, 1, 1);
      const secondBody = styleAt(table, 2, 1);
      const lastBody = styleAt(table, 4, 1);
      assert(headerMiddle.bold, styleName, "表头必须加粗");
      assert(
        firstBody.fillColor === (palette.firstRowStripeFill || palette.bodyFill),
        styleName,
        "第一行条纹填充错误"
      );
      assert(secondBody.fillColor === (palette.secondRowStripeFill || palette.bodyFill), styleName, "第二行条纹填充错误");

      // 有填充的表格区域必须遮住工作表默认网格线；无填充区域则保留 Sheet 网格线。
      for (const [label, sample] of [["表头", headerMiddle], ["奇数行", firstBody], ["偶数行", secondBody]]) {
        const gridlineProbe = document.createElement("div");
        applyRawGridlineState(gridlineProbe, sample, { showGridLines: true });
        assert(
          gridlineProbe.classList.contains("has-sheet-gridline") === !Boolean(sample.fillColor),
          styleName,
          `${label}的填充与工作表网格线状态不一致`
        );
      }

      // 关闭条纹后，所有数据行都必须回落到 wholeTable 的基础填充。
      const plainTable = makeTable(styleName, { showRowStripes: false });
      assert(
        styleAt(plainTable, 1, 1).fillColor === palette.bodyFill
          && styleAt(plainTable, 2, 1).fillColor === palette.bodyFill,
        styleName,
        "关闭行条纹后仍残留条纹填充"
      );

      // 首列、末列选项只强调对应列，不能把中间列一起加粗。
      const emphasizedTable = makeTable(styleName, {
        showFirstColumn: true,
        showLastColumn: true
      });
      assert(styleAt(emphasizedTable, 2, 0).bold, styleName, "首列强调未生效");
      assert(styleAt(emphasizedTable, 2, 2).bold, styleName, "末列强调未生效");
      assert(!styleAt(emphasizedTable, 2, 1).bold, styleName, "首末列强调错误扩散到中间列");

      // 汇总行使用独立的填充、字体和粗体，但仍保留整个表格的底边界。
      const totalsTable = makeTable(styleName, { totalsRow: true });
      const totals = styleAt(totalsTable, 4, 1);
      assert(totals.bold, styleName, "汇总行必须加粗");
      assert(totals.fillColor === palette.totalFill, styleName, "汇总行填充错误");
      assert(totals.fontColor === palette.totalFont, styleName, "汇总行字体颜色错误");

      if (family === "Light" && number <= 7) {
        assert(!headerMiddle.fillColor, styleName, "Light 1~7 表头不应填充");
        assert(Boolean(firstBody.fillColor) && !secondBody.fillColor, styleName, "Light 1~7 条纹明暗顺序错误");
        expectBorder(styleName, headerMiddle, "top", "缺少表格顶边界");
        expectBorder(styleName, headerMiddle, "bottom", "缺少表头底边界");
        expectBorder(styleName, lastBody, "bottom", "缺少表格底边界");
        expectNoBorder(styleName, firstBody, "right", "Light 1~7 不应生成内部竖线");
      } else if (family === "Light" && number <= 14) {
        assert(Boolean(headerMiddle.fillColor), styleName, "Light 8~14 缺少实色表头");
        assert(firstBody.fillColor === secondBody.fillColor, styleName, "Light 8~14 不应生成隔行底色");
        expectBorder(styleName, firstBody, "bottom", "缺少横向行分隔线");
        expectNoBorder(styleName, firstBody, "right", "Light 8~14 不应生成内部竖线");
        expectBorder(styleName, headerLeft, "left", "缺少表格左外框");
      } else if (family === "Light") {
        assert(!headerMiddle.fillColor, styleName, "Light 15~21 表头不应填充");
        assert(Boolean(firstBody.fillColor) && !secondBody.fillColor, styleName, "Light 15~21 条纹明暗顺序错误");
        expectBorder(styleName, firstBody, "bottom", "缺少单元格横向边框");
        expectBorder(styleName, firstBody, "right", "缺少单元格竖向边框");
        expectBorder(styleName, headerLeft, "left", "缺少表格左外框");
      } else if (family === "Medium" && number <= 7) {
        assert(Boolean(headerMiddle.fillColor), styleName, "Medium 1~7 缺少实色表头");
        assert(Boolean(firstBody.fillColor) && !secondBody.fillColor, styleName, "Medium 1~7 条纹明暗顺序错误");
        expectBorder(styleName, firstBody, "bottom", "缺少横向行分隔线");
        expectNoBorder(styleName, firstBody, "right", "Medium 1~7 不应生成内部竖线");
        expectBorder(styleName, headerLeft, "left", "缺少表格左外框");
      } else if (family === "Medium" && number <= 14) {
        assert(Boolean(firstBody.fillColor) && Boolean(secondBody.fillColor), styleName, "Medium 8~14 两组条纹都应填充");
        assert(firstBody.fillColor !== secondBody.fillColor, styleName, "Medium 8~14 两组条纹颜色不应相同");
        expectBorder(styleName, headerMiddle, "bottom", "缺少白色粗表头分隔线", 3);
        expectBorder(styleName, firstBody, "bottom", "缺少白色行分隔线");
        expectBorder(styleName, firstBody, "right", "缺少白色列分隔线");
      } else if (family === "Medium" && number <= 21) {
        assert(Boolean(firstBody.fillColor) && !secondBody.fillColor, styleName, "Medium 15~21 中性条纹错误");
        expectBorder(styleName, headerMiddle, "top", "缺少黑色粗顶边界", 2);
        expectBorder(styleName, headerMiddle, "bottom", "缺少黑色粗表头边界", 2);
        expectBorder(styleName, lastBody, "bottom", "缺少黑色粗底边界", 2);
        if (number === 15) {
          expectBorder(styleName, firstBody, "right", "Medium 15 缺少内部单元格边框");
        } else {
          expectNoBorder(styleName, firstBody, "right", "Medium 16~21 不应额外生成内部竖线");
        }
      } else if (family === "Medium") {
        assert(Boolean(firstBody.fillColor) && Boolean(secondBody.fillColor), styleName, "Medium 22~28 两组条纹都应填充");
        assert(firstBody.fillColor !== secondBody.fillColor, styleName, "Medium 22~28 两组条纹颜色不应相同");
        expectBorder(styleName, firstBody, "bottom", "缺少单元格横向边框");
        expectBorder(styleName, firstBody, "right", "缺少单元格竖向边框");
      } else {
        assert(Boolean(headerMiddle.fillColor), styleName, "Dark 样式缺少深色表头");
        assert(headerMiddle.fontColor === palette.headerFont, styleName, "Dark 表头字体颜色错误");
        assert(Boolean(firstBody.fillColor) && Boolean(secondBody.fillColor), styleName, "Dark 两组条纹都应填充");
        assert(firstBody.fillColor !== secondBody.fillColor, styleName, "Dark 两组条纹颜色不应相同");
        expectBorder(styleName, headerMiddle, "bottom", "Dark 表头缺少白色横向分隔线");
        expectNoBorder(styleName, firstBody, "bottom", "Dark 表体不应生成白色横向网格线");
        expectNoBorder(styleName, headerMiddle, "right", "Dark 表头不应生成白色纵向网格线");
        expectNoBorder(styleName, firstBody, "right", "Dark 表体不应生成白色纵向网格线");
        expectNoBorder(styleName, lastBody, "bottom", "Dark 表体底部不应生成额外边界");
        expectNoBorder(styleName, totals, "right", "Dark 汇总行不应生成白色纵向网格线");
      }
      checkedStyles += 1;
    }
  }

  /**
   * 使用默认 Office 主题的实测颜色做哨兵校验，既能发现主题槽映射错位，也能
   * 发现 tint 舍入、深浅条纹次序或复合色来源被改坏。这里只用于回归断言，
   * 实际渲染仍从每个工作簿的 theme1.xml 动态取色。
   */
  const exactCases = [
    ["TableStyleLight2", "headerFont", "#305496"],
    ["TableStyleLight2", "firstRowStripeFill", "#D9E1F2"],
    ["TableStyleLight9", "headerFill", "#4472C4"],
    ["TableStyleLight16", "firstRowStripeFill", "#D9E1F2"],
    ["TableStyleMedium9", "firstRowStripeFill", "#B4C6E7"],
    ["TableStyleMedium9", "secondRowStripeFill", "#D9E1F2"],
    ["TableStyleDark1", "firstRowStripeFill", "#404040"],
    ["TableStyleDark1", "secondRowStripeFill", "#737373"],
    ["TableStyleDark2", "headerFill", "#000000"],
    ["TableStyleDark2", "firstRowStripeFill", "#305496"],
    ["TableStyleDark2", "secondRowStripeFill", "#4472C4"],
    ["TableStyleDark8", "firstRowStripeFill", "#A6A6A6"],
    ["TableStyleDark8", "secondRowStripeFill", "#D9D9D9"],
    ["TableStyleDark9", "headerFill", "#ED7D31"],
    ["TableStyleDark9", "firstRowStripeFill", "#B4C6E7"],
    ["TableStyleDark9", "secondRowStripeFill", "#D9E1F2"],
    ["TableStyleDark10", "firstRowStripeFill", "#DBDBDB"],
    ["TableStyleDark10", "secondRowStripeFill", "#EDEDED"],
    ["TableStyleDark11", "headerFill", "#70AD47"],
    ["TableStyleDark11", "firstRowStripeFill", "#BDD7EE"],
    ["TableStyleDark11", "secondRowStripeFill", "#DDEBF7"]
  ];
  for (const [styleName, property, expected] of exactCases) {
    const actual = createBuiltInTablePalette(styleName, DEFAULT_THEME_COLORS)[property];
    assert(actual === expected, styleName, `${property} 应为 ${expected}，实际为 ${actual}`);
  }

  /**
   * 再用一套刻意避开 Office 默认值的主题跑完 60 项。每个样式都必须随主题
   * 改变；否则说明某处又把图库颜色写成了固定 RGB，换主题的工作簿就会偏色。
   */
  const customTheme = [
    "FFFDF6", "18212B", "EEE8DD", "34495E",
    "1F4E78", "A23E48", "5B7553", "B8860B",
    "287D8E", "6E4B8B", "1261A0", "7C365A"
  ];
  for (const [family, lastNumber] of families) {
    for (let number = 1; number <= lastNumber; number += 1) {
      const styleName = `TableStyle${family}${number}`;
      const officePalette = createBuiltInTablePalette(styleName, DEFAULT_THEME_COLORS);
      const themedPalette = createBuiltInTablePalette(styleName, customTheme);
      assert(
        JSON.stringify(officePalette) !== JSON.stringify(themedPalette),
        styleName,
        "更换工作簿主题后样式没有变化，疑似使用了固定颜色"
      );
    }
  }
  assert(
    createBuiltInTablePalette("TableStyleLight9", customTheme).headerFill === "#1F4E78",
    "TableStyleLight9",
    "自定义主题的 accent1 没有映射到表头"
  );
  assert(
    createBuiltInTablePalette("TableStyleDark10", customTheme).headerFill === "#B8860B",
    "TableStyleDark10",
    "复合样式的 accent4 表头映射错误"
  );

  // 同时开启行、列条纹时，列条纹按 OOXML 顺序后应用，并覆盖交叉处的行条纹。
  const crossingTable = makeTable("TableStyleMedium9", { showColumnStripes: true });
  assert(
    styleAt(crossingTable, 1, 1).fillColor === crossingTable.palette.secondColumnStripeFill,
    "TableStyleMedium9",
    "行列条纹交叉处没有按规范让列条纹优先"
  );
  assert(checkedStyles === 60, "全部样式", `只完成了 ${checkedStyles} 项回归`);
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
