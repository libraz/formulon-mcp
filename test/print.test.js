import assert from "node:assert/strict";
import { test } from "vitest";
import {
  applySheetOperation,
  closeSession,
  inspectSessionLayout,
  openSession,
  sessionPrintSettings,
  setSessionRange,
  setSessionSheetView,
} from "../dist/sessions.js";

test("authors a printable page and reads every setting back", async () => {
  await openSession(undefined, "print-author");
  try {
    setSessionRange(
      "print-author",
      "A1",
      [
        ["Item", "Amount"],
        ["Design", 120000],
      ],
      undefined,
      false,
    );

    const written = sessionPrintSettings("print-author", 0, {
      pageSetup: { orientation: "landscape", paperSize: 9, fitToPage: true, fitToWidth: 1 },
      margins: { left: 0.5, right: 0.5 },
      printOptions: { horizontalCentered: true, gridLines: false },
      printArea: "A1:B20",
      printTitles: { repeatRows: "1:1" },
    });
    assert.deepEqual(
      written.applied.map((step) => step.setting),
      ["pageSetup", "margins", "printOptions", "printArea", "printTitles"],
    );

    const { settings } = written;
    assert.equal(settings.pageSetup.orientation, "landscape");
    assert.equal(settings.pageSetup.paperSize, 9);
    assert.equal(settings.pageSetup.fitToPage, true);
    assert.equal(settings.margins.left, 0.5);
    // A margin that was not stated keeps Excel's default rather than zeroing.
    assert.equal(settings.margins.top, 0.75);
    assert.equal(settings.printArea, "A1:B20");
    assert.equal(settings.printTitles.repeatRows, "1:1");
    assert.match(settings.printOptionsXml, /horizontalCentered="1"/);
    assert.ok(settings.pageCount >= 1);
  } finally {
    closeSession("print-author");
  }
});

test("writes header and footer sections from decoded text", async () => {
  await openSession(undefined, "print-header");
  try {
    const { settings } = sessionPrintSettings("print-header", 0, {
      headerFooter: { oddHeader: "&CQ1 && Q2 Report", oddFooter: "&C&P / &N" },
    });
    // A literal ampersand is spelled `&&` the way Excel's header syntax does,
    // and the engine escapes it for the file, so no caller assembles XML.
    assert.match(settings.headerFooterXml, /<oddHeader>&amp;CQ1 &amp;&amp; Q2 Report<\/oddHeader>/);
    assert.match(settings.headerFooterXml, /<oddFooter>&amp;C&amp;P \/ &amp;N<\/oddFooter>/);
  } finally {
    closeSession("print-header");
  }
});

test("replaces page breaks rather than accumulating them", async () => {
  await openSession(undefined, "print-breaks");
  try {
    sessionPrintSettings("print-breaks", 0, { rowBreaks: [10, 20], colBreaks: [4] });
    const replaced = sessionPrintSettings("print-breaks", 0, { rowBreaks: [30] });
    assert.deepEqual(
      replaced.settings.rowBreaks.map((entry) => entry.id),
      [30],
    );
    // Stating one axis clears both, so the sheet's breaks are exactly what the
    // last call named.
    assert.deepEqual(replaced.settings.colBreaks, []);
  } finally {
    closeSession("print-breaks");
  }
});

test("removes a print area with an empty range list", async () => {
  await openSession(undefined, "print-clear");
  try {
    sessionPrintSettings("print-clear", 0, { printArea: "A1:C5" });
    const cleared = sessionPrintSettings("print-clear", 0, { printArea: "" });
    assert.equal(cleared.settings.printArea, null);
  } finally {
    closeSession("print-clear");
  }
});

test("refuses a set call that states no setting", async () => {
  await openSession(undefined, "print-empty");
  try {
    assert.throws(
      () => sessionPrintSettings("print-empty", 0, {}),
      /no print setting was given to set/,
    );
    // A read leaves the session clean.
    const read = sessionPrintSettings("print-empty", 0);
    assert.equal(read.session.dirty, false);
    assert.equal(read.applied, undefined);
  } finally {
    closeSession("print-empty");
  }
});

test("sets the three-state sheet tab visibility", async () => {
  await openSession(undefined, "print-visibility");
  try {
    assert.throws(
      () => setSessionSheetView("print-visibility", 0, { visibility: "hidden" }),
      /visible sheet/,
    );
    applySheetOperation("print-visibility", "add", { name: "Second" });
    const hidden = setSessionSheetView("print-visibility", 0, { visibility: "veryHidden" });
    assert.equal(hidden.statuses.length, 1);
    assert.equal(hidden.statuses[0].ok, true);
    const layout = inspectSessionLayout("print-visibility", {
      includeCells: false,
      includeStyles: false,
      maxCells: 10,
    });
    assert.deepEqual(
      layout.sheets.map((sheet) => sheet.visibility),
      ["veryHidden", "visible"],
    );
  } finally {
    closeSession("print-visibility");
  }
});

test("reads the resolved page layout of a fit-to-width landscape sheet", async () => {
  await openSession(undefined, "print-layout");
  try {
    const rows = [["Item", "Amount"]];
    for (let row = 1; row <= 120; row += 1) {
      rows.push([`Item ${row}`, row]);
    }
    setSessionRange("print-layout", "A1", rows, undefined, false);

    const { settings } = sessionPrintSettings("print-layout", 0, {
      pageSetup: {
        orientation: "landscape",
        paperSize: 9,
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0,
      },
      printTitles: { repeatRows: "1:1" },
      rowBreaks: [60],
    });
    const { layout } = settings;
    assert.equal(layout.paper.landscape, true);
    assert.equal(layout.paper.known, true);
    assert.ok(layout.paper.widthPt > layout.paper.heightPt);
    assert.equal(layout.pageOrder, "downThenOver");
    assert.ok(layout.scale > 0 && layout.scale <= 1);
    // Points, not inches: every Excel margin preset is at least 0.25 in (18 pt).
    assert.ok(layout.marginsPt.left >= 18);
    assert.ok(layout.printable.width > 0);
    assert.deepEqual(
      layout.rowPageBreaks.filter((entry) => entry.manual).map((entry) => entry.index),
      [60],
    );
    assert.ok(layout.pages.length >= 2);
    assert.equal(layout.pages.length, settings.pageCount);
    assert.equal(layout.pagesTruncated, false);
    assert.match(layout.pages[0].ref, /^A1:B\d+$/);
  } finally {
    closeSession("print-layout");
  }
});

test("a null raw print fragment removes it and the read side reports null", async () => {
  await openSession(undefined, "print-null-xml");
  try {
    const written = sessionPrintSettings("print-null-xml", 0, {
      printOptionsXml: '<printOptions horizontalCentered="1"/>',
    });
    assert.match(written.settings.printOptionsXml, /horizontalCentered="1"/);

    const removed = sessionPrintSettings("print-null-xml", 0, { printOptionsXml: null });
    assert.equal(removed.settings.printOptionsXml, null);
    // A fragment that was never written reads as absent, not as an empty string.
    assert.equal(removed.settings.sheetPrXml, null);
  } finally {
    closeSession("print-null-xml");
  }
});

test("an empty raw print fragment reaches the client as the engine's refusal", async () => {
  await openSession(undefined, "print-empty-xml");
  try {
    assert.throws(
      () => sessionPrintSettings("print-empty-xml", 0, { pageSetupXml: "" }),
      /not a valid element/,
    );
  } finally {
    closeSession("print-empty-xml");
  }
});

test("null clears a print area, the print titles and a header section", async () => {
  await openSession(undefined, "print-null-clear");
  try {
    sessionPrintSettings("print-null-clear", 0, {
      printArea: "A1:B20",
      printTitles: { repeatRows: "1:1", repeatCols: "A:A" },
      headerFooter: { oddHeader: "&CTitle", oddFooter: "&P" },
    });
    const { settings } = sessionPrintSettings("print-null-clear", 0, {
      printArea: null,
      printTitles: { repeatRows: null, repeatCols: null },
      headerFooter: { oddHeader: null },
    });
    assert.equal(settings.printArea, null);
    assert.equal(settings.printTitles.repeatRows, null);
    assert.equal(settings.printTitles.repeatCols, null);
    assert.doesNotMatch(settings.headerFooterXml ?? "", /oddHeader/);
    assert.match(settings.headerFooterXml, /oddFooter/);
  } finally {
    closeSession("print-null-clear");
  }
});
