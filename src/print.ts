import { indexToCol } from "./a1.js";
import {
  type HeaderFooterInput,
  okStatus,
  type PageSetupInput,
  type PrintOptionsInput,
  rectA1,
  resultToJson,
  type Sheet,
  StatusCode,
  type StatusJson,
  sheetAt,
  statusError,
  type Workbook,
} from "./formulon.js";
import { rangeA1 } from "./session/shared.js";

const POINTS_PER_INCH = 72;

/**
 * Worksheet print settings: page setup, margins, print options, header/footer,
 * print area, print titles and manual page breaks.
 *
 * The engine models these as separate typed setters plus a raw-XML counterpart
 * for the parts it does not model; this module gathers them into one read and
 * one write so a caller can state a printable page in a single call rather than
 * a dozen. `printOptions` and `headerFooter` have typed setters but only a raw
 * getter, so a read reports those two as their stored XML fragment.
 */

/**
 * `<pageSetup orientation>` as a tool states it: leaving the attribute off
 * (the printer default), portrait, or landscape.
 */
const ORIENTATION = ["default", "portrait", "landscape"] as const;

export type OrientationName = (typeof ORIENTATION)[number];

export type PageSetupSpec = {
  orientation?: OrientationName;
  paperSize?: number;
  scale?: number;
  fitToWidth?: number;
  fitToHeight?: number;
  fitToPage?: boolean;
};

export type PageMarginsSpec = {
  left?: number;
  right?: number;
  top?: number;
  bottom?: number;
  header?: number;
  footer?: number;
};

export type PrintTitlesSpec = {
  repeatRows?: string;
  repeatCols?: string;
};

export type PrintSettingsInput = {
  pageSetup?: PageSetupSpec;
  margins?: PageMarginsSpec;
  printOptions?: Partial<PrintOptionsInput>;
  headerFooter?: Partial<HeaderFooterInput>;
  /** Comma-separated A1 ranges, for example `A1:F40` or `A1:B10,D5:E20`. Empty removes it. */
  printArea?: string;
  printTitles?: PrintTitlesSpec;
  /** Zero-based rows each manual break precedes. Replaces the sheet's row breaks. */
  rowBreaks?: number[];
  /** Zero-based columns each manual break precedes. Replaces the sheet's column breaks. */
  colBreaks?: number[];
  /** Raw XML escape hatches for attributes the engine does not model. */
  pageSetupXml?: string;
  pageMarginsXml?: string;
  printOptionsXml?: string;
  headerFooterXml?: string;
  sheetPrXml?: string;
};

/**
 * Named page setups for a document that is meant to be printed.
 *
 * The `-fit` variants scale the sheet to one page wide and leave the height
 * unbounded (`fitToHeight: 0`), which is what a document whose row count is not
 * known in advance needs: a long line-item table flows onto a second page
 * rather than being shrunk until it is unreadable.
 */
export const PRINT_PRESETS: Record<string, PrintSettingsInput> = Object.freeze({
  "a4-portrait": {
    pageSetup: { orientation: "portrait", paperSize: 9 },
    printOptions: { horizontalCentered: true },
  },
  "a4-portrait-fit": {
    pageSetup: {
      orientation: "portrait",
      paperSize: 9,
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
    },
    printOptions: { horizontalCentered: true },
  },
  "a4-landscape": {
    pageSetup: { orientation: "landscape", paperSize: 9 },
    printOptions: { horizontalCentered: true },
  },
  "a4-landscape-fit": {
    pageSetup: {
      orientation: "landscape",
      paperSize: 9,
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
    },
    printOptions: { horizontalCentered: true },
  },
  "letter-portrait": {
    pageSetup: { orientation: "portrait", paperSize: 1 },
    printOptions: { horizontalCentered: true },
  },
  "letter-portrait-fit": {
    pageSetup: {
      orientation: "portrait",
      paperSize: 1,
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
    },
    printOptions: { horizontalCentered: true },
  },
  "letter-landscape": {
    pageSetup: { orientation: "landscape", paperSize: 1 },
    printOptions: { horizontalCentered: true },
  },
  "letter-landscape-fit": {
    pageSetup: {
      orientation: "landscape",
      paperSize: 1,
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
    },
    printOptions: { horizontalCentered: true },
  },
});

export const PRINT_PRESET_NAMES = Object.keys(PRINT_PRESETS) as [string, ...string[]];

const PAGE_ORIENTATION = { portrait: 0, landscape: 1 } as const;

function orientationCode(name: OrientationName): PageSetupInput["orientation"] | null {
  return name === "default" ? null : PAGE_ORIENTATION[name];
}

function orientationName(code: number | null): string {
  if (code === null) {
    return "default";
  }
  return ORIENTATION[code + 1] ?? `unknown(${code})`;
}

/** Pages listed individually in a print-settings read; the count is always complete. */
const MAX_REPORTED_PAGES = 50;

function pageBreaks(breaks: readonly { index: number; manual: boolean }[]) {
  return breaks.map(({ index, manual }) => ({ index, manual }));
}

function spanText(span: { first: number; last: number } | null | undefined, axis: "row" | "col") {
  if (!span) {
    return "";
  }
  const name = (index: number) => (axis === "row" ? `${index + 1}` : indexToCol(index));
  return `${name(span.first)}:${name(span.last)}`;
}

/** Reads every print setting a sheet carries, plus the page count they produce. */
export function readPrintSettings(wb: Workbook, sheetIndex: number) {
  const sheet = sheetAt(wb, sheetIndex);
  const pageSetup = sheet.getPageSetup();
  const margins = sheet.getPageMargins();
  const printArea = sheet.getPrintArea();
  const printTitles = sheet.getPrintTitles();
  const rowBreaks = sheet.listRowBreaks();
  const colBreaks = sheet.listColumnBreaks();
  const pagination = sheet.getPagination();

  return {
    pageSetup: {
      orientation: orientationName(pageSetup.orientation),
      paperSize: pageSetup.paperSize,
      scale: pageSetup.scalePct,
      fitToWidth: pageSetup.fitToWidth,
      fitToHeight: pageSetup.fitToHeight,
      fitToPage: pageSetup.fitToPage,
      orientationStated: pageSetup.orientation !== null,
      paperSizeStated: pageSetup.paperSizeStated,
      scaleStated: pageSetup.scalePctStated,
      fitToWidthStated: pageSetup.fitToWidthStated,
      fitToHeightStated: pageSetup.fitToHeightStated,
      fitToPageStated: pageSetup.fitToPageStated,
    },
    margins: {
      left: margins.leftIn,
      right: margins.rightIn,
      top: margins.topIn,
      bottom: margins.bottomIn,
      header: margins.headerIn,
      footer: margins.footerIn,
      leftStated: margins.leftInStated,
      rightStated: margins.rightInStated,
      topStated: margins.topInStated,
      bottomStated: margins.bottomInStated,
      headerStated: margins.headerInStated,
      footerStated: margins.footerInStated,
    },
    printArea: (printArea ?? []).map(rectA1).join(","),
    printTitles: {
      repeatRows: spanText(printTitles?.rows, "row"),
      repeatCols: spanText(printTitles?.columns, "col"),
    },
    rowBreaks: resultToJson(
      rowBreaks.map((b) => ({ id: b.row, min: b.first, max: b.last, manual: b.manual })),
    ),
    colBreaks: resultToJson(
      colBreaks.map((b) => ({ id: b.col, min: b.first, max: b.last, manual: b.manual })),
    ),
    // Neither of these has a structured getter, so the stored fragment is the
    // only way to read back what a typed setter wrote.
    printOptionsXml: sheet.ooxml.getPrintOptionsXml() ?? "",
    headerFooterXml: sheet.ooxml.getHeaderFooterXml() ?? "",
    sheetPrXml: sheet.ooxml.getSheetPrXml() ?? "",
    pageCount: pagination.pageCount,
    // The layout the engine actually resolves: paper and margins in points.
    layout: {
      paper: pagination.paper,
      marginsPt: {
        left: pagination.margins.leftIn * POINTS_PER_INCH,
        right: pagination.margins.rightIn * POINTS_PER_INCH,
        top: pagination.margins.topIn * POINTS_PER_INCH,
        bottom: pagination.margins.bottomIn * POINTS_PER_INCH,
        header: pagination.margins.headerIn * POINTS_PER_INCH,
        footer: pagination.margins.footerIn * POINTS_PER_INCH,
      },
      printable: {
        x: pagination.printable.x,
        y: pagination.printable.y,
        width: pagination.printable.widthPt,
        height: pagination.printable.heightPt,
      },
      // The effective scale as a factor (1 is 100%).
      scale: pagination.scalePct / 100,
      pageOrder: pagination.pageOrder === 0 ? "downThenOver" : "overThenDown",
      rowPageBreaks: pageBreaks(
        pagination.rowBreaks.map((b) => ({ index: b.row, manual: b.manual })),
      ),
      colPageBreaks: pageBreaks(
        pagination.columnBreaks.map((b) => ({ index: b.col, manual: b.manual })),
      ),
      pages: pagination.pages.slice(0, MAX_REPORTED_PAGES).map((page) => ({
        ref: rangeA1(
          page.range.firstRow,
          page.range.firstCol,
          page.range.lastRow,
          page.range.lastCol,
        ),
        ...(pagination.printArea.length > 1 ? { areaIndex: page.areaIndex } : {}),
      })),
      pagesTruncated: pagination.pages.length > MAX_REPORTED_PAGES,
    },
  };
}

type AppliedStep = { setting: string; status: StatusJson };

const XML_SETTERS = [
  ["pageSetupXml", "setPageSetupXml"],
  ["pageMarginsXml", "setPageMarginsXml"],
  ["printOptionsXml", "setPrintOptionsXml"],
  ["headerFooterXml", "setHeaderFooterXml"],
  ["sheetPrXml", "setSheetPrXml"],
] as const satisfies readonly (readonly [keyof PrintSettingsInput, keyof Sheet["ooxml"]])[];

/**
 * Applies the print settings present in `input`, leaving every other setting
 * alone. Each group is a partial update at the engine level too, so stating one
 * margin does not reset the rest.
 */
export function writePrintSettings(wb: Workbook, sheetIndex: number, input: PrintSettingsInput) {
  const sheet = sheetAt(wb, sheetIndex);
  const applied: AppliedStep[] = [];
  const step = (setting: string, apply: () => void) => {
    apply();
    applied.push({ setting, status: okStatus() });
  };

  if (input.pageSetup) {
    const { orientation, scale, ...rest } = input.pageSetup;
    const patch: Parameters<Sheet["updatePageSetup"]>[0] = {
      ...rest,
      ...(orientation !== undefined && { orientation: orientationCode(orientation) }),
      ...(scale !== undefined && { scalePct: scale }),
    };
    step("pageSetup", () => sheet.updatePageSetup(patch));
  }
  if (input.margins) {
    const { left, right, top, bottom, header, footer } = input.margins;
    const patch: Parameters<Sheet["updatePageMargins"]>[0] = {
      ...(left !== undefined && { leftIn: left }),
      ...(right !== undefined && { rightIn: right }),
      ...(top !== undefined && { topIn: top }),
      ...(bottom !== undefined && { bottomIn: bottom }),
      ...(header !== undefined && { headerIn: header }),
      ...(footer !== undefined && { footerIn: footer }),
    };
    step("margins", () => sheet.updatePageMargins(patch));
  }
  if (input.printOptions) {
    const options = input.printOptions;
    step("printOptions", () => sheet.updatePrintOptions(options));
  }
  if (input.headerFooter) {
    const headerFooter = input.headerFooter;
    step("headerFooter", () => sheet.updateHeaderFooter(headerFooter));
  }
  if (input.printArea !== undefined) {
    const ranges = input.printArea
      .split(",")
      .map((range) => range.trim())
      .filter((range) => range !== "");
    step("printArea", () => sheet.setPrintArea(ranges.length > 0 ? ranges : null));
  }
  if (input.printTitles) {
    const { repeatRows, repeatCols } = input.printTitles;
    step("printTitles", () =>
      sheet.setPrintTitles({ rows: repeatRows || null, columns: repeatCols || null }),
    );
  }
  if (input.rowBreaks !== undefined || input.colBreaks !== undefined) {
    // A stated break list replaces both axes rather than accumulating onto
    // what a loaded file already carried.
    step("clearBreaks", () => sheet.clearBreaks());
    for (const row of input.rowBreaks ?? []) {
      step(`rowBreak:${row}`, () => sheet.setRowBreak(row, { manual: true }));
    }
    for (const col of input.colBreaks ?? []) {
      step(`colBreak:${col}`, () => sheet.setColumnBreak(col, { manual: true }));
    }
  }
  for (const [key, setter] of XML_SETTERS) {
    const xml = input[key];
    if (xml !== undefined) {
      step(key, () => sheet.ooxml[setter](xml));
    }
  }

  if (applied.length === 0) {
    throw statusError(StatusCode.InvalidArgument, "no print setting was given to set");
  }
  return applied;
}
