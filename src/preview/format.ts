import { classifyNumberFormat, formatNumberValue, type NumberFormatKind } from "../numfmt.js";

export type PreviewFormatResult = {
  text: string;
  warning?: string;
};

function decimals(code: string): number {
  const match = code.replace(/"[^"]*"/g, "").match(/\.([0#?]+)/);
  return match?.[1].length ?? 0;
}

function grouping(code: string): boolean {
  return /[0#?],[0#?]/.test(code);
}

function currencySymbol(code: string): string | undefined {
  const locale = code.match(/\[\$([^\]-]+)(?:-[^\]]+)?\]/);
  if (locale?.[1]) {
    return locale[1];
  }
  const symbol = code.match(/[¥€£₩$]/);
  return symbol?.[0];
}

function formatPlainNumber(value: number, code: string): string {
  const places = decimals(code);
  return new Intl.NumberFormat("en-US", {
    useGrouping: grouping(code),
    minimumFractionDigits: places,
    maximumFractionDigits: places,
  }).format(value);
}

function formatPercent(value: number, code: string): string {
  const places = decimals(code);
  return `${new Intl.NumberFormat("en-US", {
    useGrouping: false,
    minimumFractionDigits: places,
    maximumFractionDigits: places,
  }).format(value * 100)}%`;
}

function unsupported(code: string, value: number): PreviewFormatResult {
  return {
    text: String(value),
    warning: `number format ${code || "(General)"} is approximated in preview`,
  };
}

function simpleNumericCode(code: string): boolean {
  const skeleton = code
    .replace(/\[[^\]]*\]/g, "")
    .replace(/"[^"]*"/g, "")
    .replace(/\\./g, "")
    .replace(/_.?/g, "")
    .replace(/\*.?/g, "")
    .trim();
  return /^[#,0?]+(?:\.([0#?]{1,20}))?$/.test(skeleton) && !/,\s*$/.test(skeleton);
}

function simplePercentCode(code: string): boolean {
  const skeleton = code.replace(/"[^"]*"/g, "").trim();
  return /%$/.test(skeleton) && /^[#,0?]+(?:\.([0#?]{1,20}))?%$/.test(skeleton);
}

function simpleCurrencyCode(code: string): boolean {
  const skeleton = code
    .replace(/\[[^\]]*\]/g, "")
    .replace(/"[^"]*"/g, "")
    .replace(/\\./g, "")
    .trim();
  return (
    Boolean(currencySymbol(code)) && /^[¥€£₩$]?[#,0?]+(?:\.([0#?]{1,20}))?[¥€£₩$]?$/.test(skeleton)
  );
}

function formatNumericKind(
  value: number,
  kind: NumberFormatKind,
  code: string,
): PreviewFormatResult {
  if (kind === "date" || kind === "datetime" || kind === "time" || kind === "percent") {
    if (kind === "percent" && !simplePercentCode(code)) {
      return unsupported(code, value);
    }
    if (kind === "percent") {
      return { text: formatPercent(value, code) };
    }
    const text = formatNumberValue(value, kind) ?? String(value);
    return kind === "date" || kind === "datetime" || kind === "time"
      ? { text, warning: `date/time format ${code} is rendered approximately in preview` }
      : { text };
  }
  if (kind === "currency") {
    if (!simpleCurrencyCode(code)) {
      return unsupported(code, value);
    }
    const symbol = currencySymbol(code) ?? "¤";
    const text = formatPlainNumber(Math.abs(value), code);
    return { text: value < 0 ? `-${symbol}${text}` : `${symbol}${text}` };
  }
  if (kind === "number") {
    return simpleNumericCode(code)
      ? { text: formatPlainNumber(value, code) }
      : unsupported(code, value);
  }
  return { text: String(value) };
}

/** Formats the bounded set of number formats used by common Excel workbooks. */
export function formatPreviewNumber(
  value: number,
  numFmtId: number,
  formatCode: string,
): PreviewFormatResult {
  const kind = classifyNumberFormat(numFmtId, formatCode);
  if (decimals(formatCode) > 20) {
    return unsupported(formatCode, value);
  }
  const result = formatNumericKind(value, kind, formatCode);
  return result;
}

/** Returns a stable display string for non-numeric engine values. */
export function formatPreviewText(kind: number, value: unknown, errorName?: string): string {
  if (kind === 0) {
    return "";
  }
  if (kind === 1 && typeof value === "number") {
    return String(value);
  }
  if (kind === 2) {
    return value ? "TRUE" : "FALSE";
  }
  if (kind === 3) {
    return typeof value === "string" ? value : String(value ?? "");
  }
  if (kind === 4) {
    return errorName ?? "#ERROR!";
  }
  return String(value ?? "");
}
