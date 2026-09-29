/**
 * Excel 60 种内置超级表样式的精确区域定义。
 *
 * XLSX 文件只会在 table*.xml 中保存 `TableStyleMedium2` 这样的名称，
 * 不会把内置样式的 fill/font/border 逐项写进 styles.xml。这里按 OOXML
 * 内置预设的 tableStyleElement 结构保存规则，并在运行时使用工作簿自己的
 * theme1.xml 解出最终颜色，因此不是一组固定 RGB 配色。
 *
 * 规则已与 Apache POI 官方 presetTableStyles.xml 逐项核对（Apache-2.0）：
 * https://github.com/apache/poi/blob/trunk/poi-ooxml/src/main/resources/org/apache/poi/xssf/usermodel/presetTableStyles.xml
 */

/** Excel 内置样式使用的精确 tint 常量；保留原始精度以匹配 Excel 的 HLS 舍入。 */
const TINT = Object.freeze({
  dark15: -0.14999847407452621,
  dark25: -0.249977111117893,
  dark35: -0.34998626667073579,
  dark50: -0.499984740745262,
  light15: 0.14999847407452621,
  light25: 0.249977111117893,
  light40: 0.39997558519241921,
  light45: 0.44999542222357858,
  light60: 0.59999389629810485,
  light80: 0.79998168889431442
});

/**
 * theme1.xml 的颜色顺序：0=lt1、1=dk1、4~9=accent1~accent6。
 * 规则只保存主题索引和 tint，不在源码中固化 Office 默认主题的 RGB。
 */
const THEME = Object.freeze({
  light1: 0,
  dark1: 1,
  accent1: 4,
  accent2: 5,
  accent3: 6,
  accent4: 7,
  accent5: 8,
  accent6: 9
});

function color(theme, tint) {
  return Object.freeze({ theme, tint: Number.isFinite(tint) ? tint : 0 });
}

function border(style, colorReference) {
  return Object.freeze({ style, color: colorReference });
}

function borders(definition) {
  const result = {};
  for (const [side, value] of Object.entries(definition || {})) {
    if (value) result[side] = value;
  }
  return Object.freeze(result);
}

/**
 * 一个元素只声明自己真正覆盖的属性。未声明的填充、字体或边框必须继续继承
 * 前一个区域元素，不能用空值清除；这正是表格样式交叉区域能够正确工作的关键。
 */
function element(options) {
  const result = {};
  if (options && options.fill) result.fill = options.fill;
  if (options && (options.fontColor || options.bold)) {
    result.font = Object.freeze({
      ...(options.fontColor ? { color: options.fontColor } : {}),
      ...(options.bold ? { bold: true } : {})
    });
  }
  if (options && options.borders) result.borders = options.borders;
  return Object.freeze(result);
}

const thin = (value) => border("thin", value);
const medium = (value) => border("medium", value);
const thick = (value) => border("thick", value);
const double = (value) => border("double", value);

function outerBorder(value, topStyle, bottomStyle) {
  const top = topStyle === "medium" ? medium(value) : thin(value);
  const bottom = bottomStyle === "medium" ? medium(value) : thin(value);
  return borders({ left: thin(value), right: thin(value), top, bottom });
}

function fullGrid(value, topStyle, bottomStyle) {
  return borders({
    ...outerBorder(value, topStyle, bottomStyle),
    vertical: thin(value),
    horizontal: thin(value)
  });
}

/** 按规范顺序应用；越靠后的区域只覆盖自己显式声明的属性。 */
const TABLE_STYLE_ELEMENT_ORDER = Object.freeze([
  "wholeTable",
  "firstColumnStripe",
  "secondColumnStripe",
  "firstRowStripe",
  "secondRowStripe",
  "lastColumn",
  "firstColumn",
  "headerRow",
  "totalRow"
]);

function createBuiltInTableStylePresets() {
  const presets = {};
  const white = color(THEME.light1);
  const black = color(THEME.dark1);
  const accents = [black, ...Array.from({ length: 6 }, (_, index) => color(THEME.accent1 + index))];

  const add = (family, number, elements) => {
    const name = `TableStyle${family}${number}`;
    presets[name.toLowerCase()] = Object.freeze({
      name,
      family: family.toLowerCase(),
      number,
      rowStripeSize: 1,
      columnStripeSize: 1,
      elements: Object.freeze(elements)
    });
  };

  // Light 1：黑色字体、上下细边界，以及 15% 中性灰第一条纹。
  add("Light", 1, {
    wholeTable: element({ fontColor: black, borders: borders({ top: thin(black), bottom: thin(black) }) }),
    headerRow: element({ fontColor: black, bold: true, borders: borders({ bottom: thin(black) }) }),
    totalRow: element({ fontColor: black, bold: true, borders: borders({ top: thin(black) }) }),
    firstColumn: element({ fontColor: black, bold: true }),
    lastColumn: element({ fontColor: black, bold: true }),
    firstRowStripe: element({ fill: color(THEME.light1, TINT.dark15) }),
    firstColumnStripe: element({ fill: color(THEME.light1, TINT.dark15) })
  });

  // Light 2~7：强调色深 25% 字体、强调色边界、强调色浅 80% 条纹。
  for (let number = 2; number <= 7; number += 1) {
    const accent = accents[number - 1];
    const darkFont = color(accent.theme, TINT.dark25);
    add("Light", number, {
      wholeTable: element({ fontColor: darkFont, borders: borders({ top: thin(accent), bottom: thin(accent) }) }),
      headerRow: element({ fontColor: darkFont, bold: true, borders: borders({ bottom: thin(accent) }) }),
      totalRow: element({ fontColor: darkFont, bold: true, borders: borders({ top: thin(accent) }) }),
      firstColumn: element({ fontColor: darkFont, bold: true }),
      lastColumn: element({ fontColor: darkFont, bold: true }),
      firstRowStripe: element({ fill: color(accent.theme, TINT.light80) }),
      firstColumnStripe: element({ fill: color(accent.theme, TINT.light80) })
    });
  }

  // Light 8~14：实色表头；行线/列线分别由行条纹和列条纹开关控制。
  for (let number = 8; number <= 14; number += 1) {
    const accent = accents[number - 8];
    const rowStripeBorders = borders({ top: thin(accent) });
    const columnStripeBorders = borders({ left: thin(accent) });
    add("Light", number, {
      wholeTable: element({ fontColor: black, borders: outerBorder(accent) }),
      headerRow: element({ fontColor: white, bold: true, fill: accent }),
      totalRow: element({ fontColor: black, bold: true, borders: borders({ top: double(accent) }) }),
      firstColumn: element({ fontColor: black, bold: true }),
      lastColumn: element({ fontColor: black, bold: true }),
      firstRowStripe: element({ borders: rowStripeBorders }),
      secondRowStripe: element({ borders: rowStripeBorders }),
      firstColumnStripe: element({ borders: columnStripeBorders }),
      secondColumnStripe: element({ borders: columnStripeBorders })
    });
  }

  // Light 15~21：完整细网格、表头中边框、汇总行双边框和浅色条纹。
  for (let number = 15; number <= 21; number += 1) {
    const accent = accents[number - 15];
    const stripeFill = number === 15
      ? color(THEME.light1, TINT.dark15)
      : color(accent.theme, TINT.light80);
    add("Light", number, {
      wholeTable: element({ fontColor: black, borders: fullGrid(accent) }),
      headerRow: element({ fontColor: black, bold: true, borders: borders({ bottom: medium(accent) }) }),
      totalRow: element({ fontColor: black, bold: true, borders: borders({ top: double(accent) }) }),
      firstColumn: element({ fontColor: black, bold: true }),
      lastColumn: element({ fontColor: black, bold: true }),
      firstRowStripe: element({ fill: stripeFill }),
      firstColumnStripe: element({ fill: stripeFill })
    });
  }

  // Medium 1~7：实色表头、外框与横向行线，汇总行使用强调色双线。
  for (let number = 1; number <= 7; number += 1) {
    const accent = accents[number - 1];
    const ruleColor = number === 1 ? black : color(accent.theme, TINT.light40);
    const stripeFill = number === 1
      ? color(THEME.light1, TINT.dark15)
      : color(accent.theme, TINT.light80);
    add("Medium", number, {
      wholeTable: element({
        fontColor: black,
        borders: borders({
          left: thin(ruleColor), right: thin(ruleColor), top: thin(ruleColor),
          bottom: thin(ruleColor), horizontal: thin(ruleColor)
        })
      }),
      headerRow: element({ fontColor: white, bold: true, fill: accent }),
      totalRow: element({ fontColor: black, bold: true, borders: borders({ top: double(accent) }) }),
      firstColumn: element({ fontColor: black, bold: true }),
      lastColumn: element({ fontColor: black, bold: true }),
      firstRowStripe: element({ fill: stripeFill }),
      firstColumnStripe: element({ fill: stripeFill })
    });
  }

  // Medium 8~14：浅 80% 基础填充、浅 60% 第一条纹，以及白色内部网格。
  for (let number = 8; number <= 14; number += 1) {
    const accent = accents[number - 8];
    const baseFill = number === 8
      ? color(THEME.light1, TINT.dark15)
      : color(accent.theme, TINT.light80);
    const stripeFill = number === 8
      ? color(THEME.light1, TINT.dark35)
      : color(accent.theme, TINT.light60);
    const internalGrid = borders({ vertical: thin(white), horizontal: thin(white) });
    add("Medium", number, {
      wholeTable: element({ fontColor: black, fill: baseFill, borders: internalGrid }),
      headerRow: element({ fontColor: white, bold: true, fill: accent, borders: borders({ bottom: thick(white) }) }),
      totalRow: element({ fontColor: white, bold: true, fill: accent, borders: borders({ top: thick(white) }) }),
      firstColumn: element({ fontColor: white, bold: true, fill: accent }),
      lastColumn: element({ fontColor: white, bold: true, fill: accent }),
      firstRowStripe: element({ fill: stripeFill }),
      firstColumnStripe: element({ fill: stripeFill })
    });
  }

  // Medium 15：黑色完整网格；首末列是与表头相同的深色强调块。
  add("Medium", 15, {
    wholeTable: element({ fontColor: black, borders: fullGrid(black, "medium", "medium") }),
    headerRow: element({ fontColor: white, bold: true, fill: black, borders: borders({ bottom: medium(black) }) }),
    totalRow: element({ borders: borders({ top: double(black) }) }),
    firstColumn: element({ fontColor: white, bold: true, fill: black }),
    lastColumn: element({ fontColor: white, bold: true, fill: black }),
    firstRowStripe: element({ fill: color(THEME.light1, TINT.dark15) }),
    firstColumnStripe: element({ fill: color(THEME.light1, TINT.dark15) })
  });

  // Medium 16~21：黑色上下中边框、强调色表头与首末列、中性灰第一条纹。
  for (let number = 16; number <= 21; number += 1) {
    const accent = accents[number - 15];
    add("Medium", number, {
      wholeTable: element({ fontColor: black, borders: borders({ top: medium(black), bottom: medium(black) }) }),
      headerRow: element({ fontColor: white, bold: true, fill: accent, borders: borders({ bottom: medium(black) }) }),
      totalRow: element({ borders: borders({ top: double(black) }) }),
      firstColumn: element({ fontColor: white, bold: true, fill: accent }),
      lastColumn: element({ fontColor: white, bold: true, fill: accent }),
      firstRowStripe: element({ fill: color(THEME.light1, TINT.dark15) }),
      firstColumnStripe: element({ fill: color(THEME.light1, TINT.dark15) })
    });
  }

  // Medium 22~28：浅 80% 基础填充、浅 60% 条纹和主题色完整细网格。
  for (let number = 22; number <= 28; number += 1) {
    const accent = accents[number - 22];
    const baseFill = number === 22
      ? color(THEME.light1, TINT.dark15)
      : color(accent.theme, TINT.light80);
    const stripeFill = number === 22
      ? color(THEME.light1, TINT.dark35)
      : color(accent.theme, TINT.light60);
    const gridColor = number === 22 ? black : color(accent.theme, TINT.light40);
    add("Medium", number, {
      wholeTable: element({ fontColor: black, fill: baseFill, borders: fullGrid(gridColor) }),
      headerRow: element({ fontColor: black, bold: true }),
      totalRow: element({ fontColor: black, bold: true, borders: borders({ top: medium(accent) }) }),
      firstColumn: element({ fontColor: black, bold: true }),
      lastColumn: element({ fontColor: black, bold: true }),
      firstRowStripe: element({ fill: stripeFill }),
      firstColumnStripe: element({ fill: stripeFill })
    });
  }

  // Dark 1：深色 1 的 45% 基础填充、25% 条纹和 15% 汇总行。
  add("Dark", 1, {
    wholeTable: element({ fontColor: white, fill: color(THEME.dark1, TINT.light45) }),
    headerRow: element({ fontColor: white, bold: true, fill: black, borders: borders({ bottom: medium(white) }) }),
    totalRow: element({ fontColor: white, bold: true, fill: color(THEME.dark1, TINT.light15), borders: borders({ top: medium(white) }) }),
    firstColumn: element({ fontColor: white, bold: true, fill: color(THEME.dark1, TINT.light25), borders: borders({ right: medium(white) }) }),
    lastColumn: element({ fontColor: white, bold: true, fill: color(THEME.dark1, TINT.light25), borders: borders({ left: medium(white) }) }),
    firstRowStripe: element({ fill: color(THEME.dark1, TINT.light25) }),
    firstColumnStripe: element({ fill: color(THEME.dark1, TINT.light25) })
  });

  // Dark 2~7：强调色表体、深 25% 条纹/首末列、深 50% 汇总行。
  for (let number = 2; number <= 7; number += 1) {
    const accent = accents[number - 1];
    add("Dark", number, {
      wholeTable: element({ fontColor: white, fill: accent }),
      headerRow: element({ fontColor: white, bold: true, fill: black, borders: borders({ bottom: medium(white) }) }),
      totalRow: element({ fontColor: white, bold: true, fill: color(accent.theme, TINT.dark50), borders: borders({ top: medium(white) }) }),
      firstColumn: element({ fontColor: white, bold: true, fill: color(accent.theme, TINT.dark25), borders: borders({ right: medium(white) }) }),
      lastColumn: element({ fontColor: white, bold: true, fill: color(accent.theme, TINT.dark25), borders: borders({ left: medium(white) }) }),
      firstRowStripe: element({ fill: color(accent.theme, TINT.dark25) }),
      firstColumnStripe: element({ fill: color(accent.theme, TINT.dark25) })
    });
  }

  // Dark 8~11：Excel 图库中的复合色组合；汇总行继承表体填充并使用黑色双线。
  const darkComposite = [
    { number: 8, bodyTheme: THEME.light1, headerTheme: THEME.dark1, neutral: true },
    { number: 9, bodyTheme: THEME.accent1, headerTheme: THEME.accent2 },
    { number: 10, bodyTheme: THEME.accent3, headerTheme: THEME.accent4 },
    { number: 11, bodyTheme: THEME.accent5, headerTheme: THEME.accent6 }
  ];
  for (const item of darkComposite) {
    const baseFill = color(item.bodyTheme, item.neutral ? TINT.dark15 : TINT.light80);
    const stripeFill = color(item.bodyTheme, item.neutral ? TINT.dark35 : TINT.light60);
    add("Dark", item.number, {
      wholeTable: element({ fill: baseFill }),
      headerRow: element({ fontColor: white, fill: color(item.headerTheme) }),
      totalRow: element({ fontColor: black, bold: true, borders: borders({ top: double(black) }) }),
      firstColumn: element({ fontColor: black, bold: true }),
      lastColumn: element({ fontColor: black, bold: true }),
      firstRowStripe: element({ fill: stripeFill }),
      firstColumnStripe: element({ fill: stripeFill })
    });
  }

  const expectedFamilies = [["Light", 21], ["Medium", 28], ["Dark", 11]];
  for (const [family, count] of expectedFamilies) {
    for (let number = 1; number <= count; number += 1) {
      const name = `TableStyle${family}${number}`;
      if (!presets[name.toLowerCase()]) throw new Error(`缺少内置超级表样式：${name}`);
    }
  }
  if (Object.keys(presets).length !== 60) {
    throw new Error("内置超级表样式必须且只能包含 60 项。");
  }
  return Object.freeze(presets);
}

const BUILT_IN_TABLE_STYLE_PRESETS = createBuiltInTableStylePresets();

export {
  BUILT_IN_TABLE_STYLE_PRESETS,
  TABLE_STYLE_ELEMENT_ORDER
};
