import assert from "node:assert/strict";
import { test } from "vitest";
import {
  classifyNumberFormat,
  describeNumberFormat,
  excelSerialToDate,
  formatNumberValue,
} from "../dist/numfmt.js";

test("classifies format codes into semantic kinds", () => {
  assert.equal(classifyNumberFormat(164, "yyyy-mm-dd"), "date");
  assert.equal(classifyNumberFormat(22, "m/d/yy h:mm"), "datetime");
  assert.equal(classifyNumberFormat(21, "h:mm:ss"), "time");
  assert.equal(classifyNumberFormat(164, '"¥"#,##0'), "currency");
  assert.equal(classifyNumberFormat(164, "[$€-407]#,##0.00"), "currency");
  assert.equal(classifyNumberFormat(9, "0%"), "percent");
  assert.equal(classifyNumberFormat(164, "#,##0.00"), "number");
  assert.equal(classifyNumberFormat(0, "General"), "general");
  // A locale-only bracket tag on a date must stay a date, not become currency.
  assert.equal(classifyNumberFormat(164, "[$-409]mmm-yy"), "date");
});

test("converts Excel serials to calendar dates", () => {
  assert.equal(excelSerialToDate(45000).toISOString().slice(0, 10), "2023-03-15");
  assert.equal(excelSerialToDate(25569).toISOString().slice(0, 10), "1970-01-01");
});

test("formats stored numbers per kind", () => {
  assert.equal(formatNumberValue(45000, "date"), "2023-03-15");
  assert.equal(formatNumberValue(0.153, "percent"), "15.3%");
  assert.equal(formatNumberValue(1500, "currency"), undefined);
  assert.equal(formatNumberValue(1500, "number"), undefined);
});

test("describes only formats that add meaning", () => {
  assert.deepEqual(describeNumberFormat(164, "yyyy-mm-dd", 45000), {
    numberFormat: "yyyy-mm-dd",
    formatKind: "date",
    formatted: "2023-03-15",
  });
  assert.deepEqual(describeNumberFormat(164, '"¥"#,##0', 1500), {
    numberFormat: '"¥"#,##0',
    formatKind: "currency",
  });
  assert.equal(describeNumberFormat(164, "#,##0.00", 1500), undefined);
  assert.equal(describeNumberFormat(0, "General", 1), undefined);
});
