import {
  assertStatus,
  type HeaderFooterInput,
  type PageMarginsInput,
  type PageSetupInput,
  type PrintOptionsInput,
  resultToJson,
  type Status,
  statusToJson,
  type Workbook,
} from "./formulon.js";

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
 * `<pageSetup orientation>` as the engine encodes it: leaving the attribute off
 * (the printer default), portrait, or landscape. Declared here rather than
 * imported because the engine keeps `Orientation` a compile-time-only union
 * with no runtime table to defer to.
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

export type PrintTitlesSpec = {
  repeatRows?: string;
  repeatCols?: string;
};

export type PrintSettingsInput = {
  pageSetup?: PageSetupSpec;
  margins?: PageMarginsInput;
  printOptions?: PrintOptionsInput;
  headerFooter?: HeaderFooterInput;
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

function orientationCode(name: OrientationName): number {
  const code = ORIENTATION.indexOf(name);
  if (code < 0) {
    throw new Error(`unknown orientation: ${name} (expected ${ORIENTATION.join(", ")})`);
  }
  return code;
}

function orientationName(code: number): string {
  return ORIENTATION[code] ?? `unknown(${code})`;
}

/** Reads every print setting a sheet carries, plus the page count they produce. */
export function readPrintSettings(wb: Workbook, sheet: number) {
  const pageSetup = wb.getSheetPageSetup(sheet);
  assertStatus(pageSetup.status, "read page setup");
  const margins = wb.getSheetPageMargins(sheet);
  assertStatus(margins.status, "read page margins");
  const printArea = wb.getSheetPrintArea(sheet);
  assertStatus(printArea.status, "read print area");
  const printTitles = wb.getSheetPrintTitles(sheet);
  assertStatus(printTitles.status, "read print titles");
  const rowBreaks = wb.getSheetRowBreaks(sheet);
  assertStatus(rowBreaks.status, "read row breaks");
  const colBreaks = wb.getSheetColBreaks(sheet);
  assertStatus(colBreaks.status, "read column breaks");
  const pagination = wb.paginate(sheet);
  assertStatus(pagination.status, "paginate sheet");

  const { status: _setupStatus, orientation, ...setupRest } = pageSetup;
  const { status: _marginStatus, ...marginRest } = margins;
  return {
    pageSetup: { orientation: orientationName(orientation), ...setupRest },
    margins: marginRest,
    printArea: printArea.ranges,
    printTitles: { repeatRows: printTitles.repeatRows, repeatCols: printTitles.repeatCols },
    rowBreaks: resultToJson(rowBreaks.breaks),
    colBreaks: resultToJson(colBreaks.breaks),
    // Neither of these has a structured getter, so the stored fragment is the
    // only way to read back what a typed setter wrote.
    printOptionsXml: wb.getSheetPrintOptionsXml(sheet).xml,
    headerFooterXml: wb.getSheetHeaderFooterXml(sheet).xml,
    sheetPrXml: wb.getSheetSheetPrXml(sheet).xml,
    pageCount: pagination.pageCount,
  };
}

type AppliedStep = { setting: string; status: ReturnType<typeof statusToJson> };

/**
 * Applies the print settings present in `input`, leaving every other setting
 * alone. Each group is a partial update at the engine level too, so stating one
 * margin does not reset the rest.
 */
export function writePrintSettings(wb: Workbook, sheet: number, input: PrintSettingsInput) {
  const applied: AppliedStep[] = [];
  const step = (setting: string, status: Status) => {
    assertStatus(status, `set ${setting}`);
    applied.push({ setting, status: statusToJson(status) });
  };

  if (input.pageSetup) {
    const { orientation, ...rest } = input.pageSetup;
    const setup: PageSetupInput = { ...rest };
    if (orientation !== undefined) {
      setup.orientation = orientationCode(orientation) as PageSetupInput["orientation"];
    }
    step("pageSetup", wb.setSheetPageSetup(sheet, setup));
  }
  if (input.margins) {
    step("margins", wb.setSheetPageMargins(sheet, input.margins));
  }
  if (input.printOptions) {
    step("printOptions", wb.setSheetPrintOptions(sheet, input.printOptions));
  }
  if (input.headerFooter) {
    step("headerFooter", wb.setSheetHeaderFooter(sheet, input.headerFooter));
  }
  if (input.printArea !== undefined) {
    step("printArea", wb.setSheetPrintArea(sheet, input.printArea));
  }
  if (input.printTitles) {
    step(
      "printTitles",
      wb.setSheetPrintTitles(
        sheet,
        input.printTitles.repeatRows ?? "",
        input.printTitles.repeatCols ?? "",
      ),
    );
  }
  if (input.rowBreaks !== undefined || input.colBreaks !== undefined) {
    // The engine upserts one break at a time and has no per-axis clear, so a
    // stated break list replaces both axes rather than accumulating onto what
    // a loaded file already carried.
    step("clearBreaks", wb.clearSheetBreaks(sheet));
    for (const row of input.rowBreaks ?? []) {
      step(`rowBreak:${row}`, wb.addSheetRowBreak(sheet, row, true));
    }
    for (const col of input.colBreaks ?? []) {
      step(`colBreak:${col}`, wb.addSheetColBreak(sheet, col, true));
    }
  }
  for (const [key, setter] of [
    ["pageSetupXml", wb.setSheetPageSetupXml],
    ["pageMarginsXml", wb.setSheetPageMarginsXml],
    ["printOptionsXml", wb.setSheetPrintOptionsXml],
    ["headerFooterXml", wb.setSheetHeaderFooterXml],
    ["sheetPrXml", wb.setSheetSheetPrXml],
  ] as const) {
    const xml = input[key];
    if (xml !== undefined) {
      step(key, setter.call(wb, sheet, xml));
    }
  }

  if (applied.length === 0) {
    throw new Error("no print setting was given to set");
  }
  return applied;
}
