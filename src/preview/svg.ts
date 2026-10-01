import { Resvg } from "@resvg/resvg-js";
import { pointsToPixels } from "./geometry.js";
import type { PreviewCell, PreviewModel, PreviewStyle } from "./types.js";

const GRID_COLOR = "#D9E0E7";
const PAGE_BREAK_COLOR = "#6B7280";

function xml(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function n(value: number): string {
  return Number.isFinite(value) ? value.toFixed(3).replace(/\.000$/, "") : "0";
}

function px(value: number): number {
  return pointsToPixels(value);
}

function edgeStyle(style: number): { width: number; dash?: string } {
  if (style === 0) {
    return { width: 0 };
  }
  if (style === 2) {
    return { width: 2 };
  }
  if (style === 3 || style === 9 || style === 11 || style === 13) {
    return { width: 1, dash: "4 3" };
  }
  if (style === 4) {
    return { width: 1, dash: "1 3" };
  }
  if (style === 8 || style === 10 || style === 12) {
    return { width: 2, dash: "4 3" };
  }
  if (style === 5) {
    return { width: 3 };
  }
  if (style === 6) {
    return { width: 2 };
  }
  if (style === 7) {
    return { width: 0.5 };
  }
  return { width: 1 };
}

function positions(lengths: number[]): number[] {
  const result = [0];
  for (const length of lengths) {
    result.push(result[result.length - 1] + length);
  }
  return result;
}

function mergeRect(cell: PreviewCell, model: PreviewModel) {
  const columns = model.columns;
  const rows = model.rows;
  if (!cell.merge) {
    return { x: cell.x, y: cell.y, width: cell.width, height: cell.height };
  }
  const colPos = positions(columns);
  const rowPos = positions(rows);
  const firstCol = cell.merge.firstCol - model.firstCol;
  const firstRow = cell.merge.firstRow - model.firstRow;
  const lastCol = cell.merge.lastCol - model.firstCol;
  const lastRow = cell.merge.lastRow - model.firstRow;
  return {
    x: colPos[firstCol] ?? cell.x,
    y: rowPos[firstRow] ?? cell.y,
    width: colPos[lastCol + 1] - (colPos[firstCol] ?? 0),
    height: rowPos[lastRow + 1] - (rowPos[firstRow] ?? 0),
  };
}

function isCoveredCell(cell: PreviewCell): boolean {
  return Boolean(
    cell.merge && (cell.row !== cell.merge.firstRow || cell.col !== cell.merge.firstCol),
  );
}

function drawBorder(
  output: string[],
  x: number,
  y: number,
  width: number,
  height: number,
  style: PreviewStyle,
  visibleSides = { top: true, bottom: true, left: true, right: true },
): void {
  const segments = [
    [style.borders.top, x, y, x + width, y],
    [style.borders.bottom, x, y + height, x + width, y + height],
    [style.borders.left, x, y, x, y + height],
    [style.borders.right, x + width, y, x + width, y + height],
  ] as const;
  for (const [border, x1, y1, x2, y2] of segments) {
    const side = y1 === y2 ? (y1 === y ? "top" : "bottom") : x1 === x ? "left" : "right";
    if (!visibleSides[side]) {
      continue;
    }
    const shape = edgeStyle(border.style);
    if (shape.width === 0) {
      continue;
    }
    output.push(
      `<line x1="${n(px(x1))}" y1="${n(px(y1))}" x2="${n(px(x2))}" y2="${n(px(y2))}" stroke="${xml(border.color)}" stroke-width="${n(px(shape.width))}"${shape.dash ? ` stroke-dasharray="${shape.dash}"` : ""} />`,
    );
  }
}

function glyphWidth(character: string, fontSize: number, average: number): number {
  if (/\p{Mark}/u.test(character) || character === "\u200d" || /[\uFE00-\uFE0F]/u.test(character)) {
    return 0;
  }
  if (
    /\p{Script=Han}|\p{Script=Hiragana}|\p{Script=Katakana}|\p{Script=Hangul}|\p{Extended_Pictographic}/u.test(
      character,
    ) ||
    /[\uFF01-\uFF60\uFFE0-\uFFE6]/u.test(character)
  ) {
    return fontSize;
  }
  return average;
}

function drawText(
  output: string[],
  cell: PreviewCell,
  rect: { x: number; y: number; width: number; height: number },
  clipId: string,
): void {
  if (!cell.text) {
    return;
  }
  const style = cell.style;
  const fontSize = px(style.fontSize);
  const padding = px(2 + style.indent * 2);
  const available = Math.max(1, px(rect.width) - padding * 2);
  const average = Math.max(1, fontSize * 0.55);
  const maxWidth = Math.max(1, available);
  const wrapLine = (line: string): string[] => {
    if (!line) {
      return [""];
    }
    const lines: string[] = [];
    let current = "";
    let currentWidth = 0;
    for (const character of Array.from(line)) {
      const width = glyphWidth(character, fontSize, average);
      if (current && currentWidth + width > maxWidth) {
        lines.push(current);
        current = "";
        currentWidth = 0;
      }
      current += character;
      currentWidth += width;
    }
    if (current) {
      lines.push(current);
    }
    return lines;
  };
  const lines = style.wrapText
    ? cell.text.split(/\r?\n/).flatMap(wrapLine)
    : [cell.text.replace(/[\r\n]+/g, " ")];
  const lineHeight = fontSize * 1.2;
  const textHeight = lineHeight * lines.length;
  const vertical =
    style.vertical === 1
      ? (px(rect.height) - textHeight) / 2
      : style.vertical === 2
        ? px(rect.height) - textHeight - padding
        : padding;
  const horizontal = style.horizontal === 0 ? (cell.numeric ? 3 : 1) : style.horizontal;
  const anchor = horizontal === 2 ? "middle" : horizontal === 3 ? "end" : "start";
  const x =
    horizontal === 2
      ? px(rect.x + rect.width / 2)
      : horizontal === 3
        ? px(rect.x + rect.width) - padding
        : px(rect.x) + padding;
  const textStyle = [
    `font-family="${xml(style.fontFamily)}"`,
    `font-size="${n(fontSize)}"`,
    `fill="${xml(style.fontColor)}"`,
    `text-anchor="${anchor}"`,
    style.bold ? 'font-weight="700"' : "",
    style.italic ? 'font-style="italic"' : "",
    style.underline ? 'text-decoration="underline"' : "",
  ]
    .filter(Boolean)
    .join(" ");
  output.push(
    `<text x="${n(x)}" y="${n(px(rect.y) + vertical + fontSize * 0.82)}" ${textStyle} clip-path="url(#${clipId})">`,
  );
  lines.forEach((line, index) => {
    const dy = index === 0 ? 0 : lineHeight;
    output.push(`<tspan x="${n(x)}" dy="${n(dy)}">${xml(line)}</tspan>`);
  });
  output.push("</text>");
}

function mergeBoundarySets(model: PreviewModel): {
  vertical: Set<string>;
  horizontal: Set<string>;
} {
  const vertical = new Set<string>();
  const horizontal = new Set<string>();
  const seenMerges = new Set<string>();
  for (const cell of model.cells) {
    if (!cell.merge) {
      continue;
    }
    const mergeKey = `${cell.merge.firstRow},${cell.merge.firstCol},${cell.merge.lastRow},${cell.merge.lastCol}`;
    if (seenMerges.has(mergeKey)) {
      continue;
    }
    seenMerges.add(mergeKey);
    for (let row = cell.merge.firstRow; row <= cell.merge.lastRow; row += 1) {
      for (let boundary = cell.merge.firstCol + 1; boundary <= cell.merge.lastCol; boundary += 1) {
        vertical.add(`${boundary},${row}`);
      }
    }
    for (let col = cell.merge.firstCol; col <= cell.merge.lastCol; col += 1) {
      for (let boundary = cell.merge.firstRow + 1; boundary <= cell.merge.lastRow; boundary += 1) {
        horizontal.add(`${boundary},${col}`);
      }
    }
  }
  return { vertical, horizontal };
}

/** Builds a deterministic SVG workbook snapshot from the prepared model. */
export function renderPreviewSvg(model: PreviewModel): string {
  // Keep the root viewport integral. Resvg rounds its raster dimensions from
  // the root viewport, so using the same ceil here keeps PNG metadata stable
  // at fractional point-to-pixel conversions (for example default row height).
  const width = Math.ceil(px(model.logicalWidth));
  const height = Math.ceil(px(model.logicalHeight));
  const columns = positions(model.columns);
  const rows = positions(model.rows);
  const mergedBoundaries = mergeBoundarySets(model);
  const output = [
    `<svg xmlns="http://www.w3.org/2000/svg" width="${n(width)}" height="${n(height)}" viewBox="0 0 ${n(width)} ${n(height)}">`,
    `<rect width="${n(width)}" height="${n(height)}" fill="#FFFFFF" />`,
    "<defs>",
  ];
  model.cells.forEach((cell, index) => {
    if (!isCoveredCell(cell)) {
      const rect = mergeRect(cell, model);
      output.push(
        `<clipPath id="clip-${index}"><rect x="${n(px(rect.x))}" y="${n(px(rect.y))}" width="${n(px(rect.width))}" height="${n(px(rect.height))}" /></clipPath>`,
      );
    }
  });
  output.push("</defs>");

  for (const cell of model.cells) {
    if (isCoveredCell(cell)) {
      continue;
    }
    const rect = mergeRect(cell, model);
    if (cell.style.fill) {
      output.push(
        `<rect x="${n(px(rect.x))}" y="${n(px(rect.y))}" width="${n(px(rect.width))}" height="${n(px(rect.height))}" fill="${xml(cell.style.fill)}" />`,
      );
    }
  }
  if (model.showGridLines) {
    for (let index = 1; index < columns.length - 1; index += 1) {
      const absolute = model.firstCol + index;
      for (let row = 0; row < rows.length - 1; row += 1) {
        const absoluteRow = model.firstRow + row;
        if (mergedBoundaries.vertical.has(`${absolute},${absoluteRow}`)) {
          continue;
        }
        output.push(
          `<line x1="${n(px(columns[index]))}" y1="${n(px(rows[row]))}" x2="${n(px(columns[index]))}" y2="${n(px(rows[row + 1]))}" stroke="${GRID_COLOR}" stroke-width="${n(px(0.5))}" />`,
        );
      }
    }
    for (let index = 1; index < rows.length - 1; index += 1) {
      const absolute = model.firstRow + index;
      for (let col = 0; col < columns.length - 1; col += 1) {
        const absoluteCol = model.firstCol + col;
        if (mergedBoundaries.horizontal.has(`${absolute},${absoluteCol}`)) {
          continue;
        }
        output.push(
          `<line x1="${n(px(columns[col]))}" y1="${n(px(rows[index]))}" x2="${n(px(columns[col + 1]))}" y2="${n(px(rows[index]))}" stroke="${GRID_COLOR}" stroke-width="${n(px(0.5))}" />`,
        );
      }
    }
  }
  model.cells.forEach((cell, index) => {
    const rect = mergeRect(cell, model);
    if (cell.merge) {
      const merge = cell.merge;
      drawBorder(output, cell.x, cell.y, cell.width, cell.height, cell.style, {
        top: cell.row === merge.firstRow,
        bottom: cell.row === merge.lastRow,
        left: cell.col === merge.firstCol,
        right: cell.col === merge.lastCol,
      });
    } else {
      drawBorder(output, rect.x, rect.y, rect.width, rect.height, cell.style);
    }
    if (!isCoveredCell(cell)) {
      drawText(output, cell, rect, `clip-${index}`);
    }
  });
  if (model.showPageBreaks) {
    for (const row of model.rowBreaks) {
      const index = row - model.firstRow;
      if (index >= 0 && index <= rows.length - 1) {
        output.push(
          `<line x1="0" y1="${n(px(rows[index]))}" x2="${n(width)}" y2="${n(px(rows[index]))}" stroke="${PAGE_BREAK_COLOR}" stroke-width="${n(px(1))}" stroke-dasharray="${n(px(2))} ${n(px(2))}" />`,
        );
      }
    }
    for (const col of model.colBreaks) {
      const index = col - model.firstCol;
      if (index >= 0 && index <= columns.length - 1) {
        output.push(
          `<line x1="${n(px(columns[index]))}" y1="0" x2="${n(px(columns[index]))}" y2="${n(height)}" stroke="${PAGE_BREAK_COLOR}" stroke-width="${n(px(1))}" stroke-dasharray="${n(px(2))} ${n(px(2))}" />`,
        );
      }
    }
  }
  output.push("</svg>");
  return output.join("");
}

/** Rasterizes the SVG through Resvg with deterministic font/loading options. */
export function renderPreviewPng(svg: string, scale: number): Buffer {
  const rendered = new Resvg(svg, {
    font: { loadSystemFonts: true },
    fitTo: { mode: "zoom", value: scale },
    logLevel: "off",
  }).render();
  return rendered.asPng();
}
