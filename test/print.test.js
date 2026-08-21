import assert from "node:assert/strict";
import { test } from "vitest";
import {
  closeSession,
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
    assert.match(settings.printOptionsXml, /horizontalCentered="true"/);
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
    assert.equal(cleared.settings.printArea, "");
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
    const hidden = setSessionSheetView("print-visibility", 0, { visibility: "veryHidden" });
    assert.equal(hidden.statuses.length, 1);
    assert.equal(hidden.statuses[0].ok, true);
  } finally {
    closeSession("print-visibility");
  }
});
