/**
 * Excel number-format classification and value decoding.
 *
 * Cells store raw numbers; their human meaning (a date, a currency amount, a
 * percentage) lives in the applied number format. These helpers classify a
 * format code and decode a stored number into a readable string so a range
 * read can surface `2023-03-15` instead of the bare serial `45000`.
 */

/** Broad semantic class of an Excel number format. */
export type NumberFormatKind =
  | "date"
  | "datetime"
  | "time"
  | "currency"
  | "percent"
  | "number"
  | "general";

/** Decoded annotation for a formatted numeric cell. */
export type NumberFormatInfo = {
  numberFormat: string;
  formatKind: NumberFormatKind;
  formatted?: string;
};

/**
 * Removes the parts of a format code that never carry date/time/currency
 * tokens: bracket sections (`[Red]`, `[$-409]`), quoted literals, escaped
 * characters, and `_`/`*` padding directives. What remains is the bare token
 * skeleton that the classifier scans.
 */
function tokenSkeleton(code: string): string {
  return code
    .replace(/\[[^\]]*\]/g, "")
    .replace(/"[^"]*"/g, "")
    .replace(/\\./g, "")
    .replace(/_.?/g, "")
    .replace(/\*.?/g, "");
}

/** Detects a currency symbol or `[$…]` currency-locale tag in a raw format code. */
function looksLikeCurrency(code: string): boolean {
  // `[$¥-411]` etc. carry a symbol before the locale dash; `[$-409]` is a
  // locale-only tag (no leading symbol) and must not count as currency.
  if (/\[\$[^\]\d-][^\]]*\]/.test(code)) {
    return true;
  }
  return /[¥€£₩]/.test(code) || /(^|[^[])\$/.test(code);
}

/**
 * Classifies a format code (with its numeric id as a fallback) into a
 * {@link NumberFormatKind}. Date/time detection wins over currency/percent so a
 * serial is only ever decoded as a calendar value when the format truly is one.
 */
export function classifyNumberFormat(numFmtId: number, formatCode: string): NumberFormatKind {
  const code = formatCode ?? "";
  if (code === "" || /^general$/i.test(code.trim())) {
    return numFmtId === 0 ? "general" : "number";
  }
  const skeleton = tokenSkeleton(code);
  const hasYear = /y/i.test(skeleton);
  const hasDay = /d/i.test(skeleton);
  const hasMonth = /m/i.test(skeleton);
  const hasTime = /[hs]/i.test(skeleton);
  // `m` is minutes when a time token is present with no explicit day/year, so it
  // only counts toward a date when there is a clear date context.
  const hasDate = hasYear || hasDay || (hasMonth && !hasTime);
  if (hasDate && hasTime) {
    return "datetime";
  }
  if (hasDate) {
    return "date";
  }
  if (hasTime) {
    return "time";
  }
  if (looksLikeCurrency(code)) {
    return "currency";
  }
  if (code.includes("%")) {
    return "percent";
  }
  return "number";
}

/**
 * Converts an Excel serial number to a `Date`.
 *
 * Uses the anchor `serial 25569 == 1970-01-01`, which absorbs Excel's spurious
 * 1900 leap-year day for every date on or after 1900-03-01 — i.e. all
 * real-world dates. Serials below 60 (pre-1900-03) are off by one day; those
 * are not decoded in practice.
 */
export function excelSerialToDate(serial: number): Date {
  return new Date(Math.round((serial - 25569) * 86400 * 1000));
}

/**
 * Converts an ISO-8601 date or date-time to an Excel serial, the inverse of
 * {@link excelSerialToDate}. Returns `undefined` for anything that is not a
 * date literal, so a caller can fall back to writing the value as text.
 */
export function isoToExcelSerial(value: string): number | undefined {
  const match = value
    .trim()
    .match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?$/);
  if (!match) {
    return undefined;
  }
  const [, year, month, day, hour, minute, second] = match;
  const utc = Date.UTC(
    Number(year),
    Number(month) - 1,
    Number(day),
    Number(hour ?? 0),
    Number(minute ?? 0),
    Number(second ?? 0),
  );
  if (Number.isNaN(utc)) {
    return undefined;
  }
  return utc / 86400 / 1000 + 25569;
}

function pad2(value: number): string {
  return String(value).padStart(2, "0");
}

function isoDate(date: Date): string {
  return `${date.getUTCFullYear()}-${pad2(date.getUTCMonth() + 1)}-${pad2(date.getUTCDate())}`;
}

function isoTime(date: Date): string {
  return `${pad2(date.getUTCHours())}:${pad2(date.getUTCMinutes())}:${pad2(date.getUTCSeconds())}`;
}

/**
 * Renders a stored number as a readable string for a given format kind, or
 * `undefined` when the kind carries no useful transformation (plain numbers
 * are already readable). Dates below serial 1 that carry a date format are left
 * undecoded because they cannot represent a real calendar day.
 */
export function formatNumberValue(value: number, kind: NumberFormatKind): string | undefined {
  if (!Number.isFinite(value)) {
    return undefined;
  }
  switch (kind) {
    case "date":
      return value < 1 ? undefined : isoDate(excelSerialToDate(value));
    case "datetime": {
      const date = excelSerialToDate(value);
      return `${isoDate(date)}T${isoTime(date)}`;
    }
    case "time":
      return isoTime(excelSerialToDate(value));
    case "percent":
      return `${Number((value * 100).toPrecision(15))}%`;
    default:
      return undefined;
  }
}

/**
 * Builds the format annotation for a numeric cell value. Returns `undefined`
 * for plain/general numbers so unformatted cells stay lean in the output.
 */
export function describeNumberFormat(
  numFmtId: number,
  formatCode: string,
  value: number,
): NumberFormatInfo | undefined {
  const formatKind = classifyNumberFormat(numFmtId, formatCode);
  if (formatKind === "general" || formatKind === "number") {
    return undefined;
  }
  const info: NumberFormatInfo = { numberFormat: formatCode, formatKind };
  const formatted = formatNumberValue(value, formatKind);
  if (formatted !== undefined) {
    info.formatted = formatted;
  }
  return info;
}
