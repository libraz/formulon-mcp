import type {
  BlanksAs,
  Color,
  SheetRect,
  SparklineGroup,
  SparklineGroupPatch,
} from "@libraz/formulon";
import { cellToA1, parseCellRef, parseRangeRef } from "../a1.js";
import { findSheetIndex, rectA1, SparklineAxisType, SparklineType, sheetAt } from "../formulon.js";
import { formatColor, parseColor } from "../styles/input.js";
import { assertInGrid, sheetName } from "./shared.js";
import { getSession, publicInfo, touch } from "./state.js";

const TYPE_NAMES = ["line", "column", "winLoss"] as const;
const TYPE_CODES: Record<SparklineTypeName, SparklineType> = {
  line: SparklineType.Line,
  column: SparklineType.Column,
  winLoss: SparklineType.Stacked,
};
const EMPTY_NAMES = ["gap", "zero", "span"] as const;
const EMPTY_CODES: Record<SparklineEmptyName, BlanksAs> = { gap: 1, zero: 0, span: 2 };
const COLOR_FIELDS = {
  series: "colorSeries",
  negative: "colorNegative",
  axis: "colorAxis",
  markers: "colorMarkers",
  first: "colorFirst",
  last: "colorLast",
  high: "colorHigh",
  low: "colorLow",
} as const;
const FLAGS = [
  "markers",
  "high",
  "low",
  "first",
  "last",
  "negative",
  "displayXAxis",
  "displayHidden",
] as const;

export type SparklineTypeName = (typeof TYPE_NAMES)[number];
export type SparklineEmptyName = (typeof EMPTY_NAMES)[number];
export type SparklineColorName = keyof typeof COLOR_FIELDS;
/** `individual` scales each sparkline alone, `group` shares one scale, a number fixes the bound. */
export type SparklineAxisBound = "individual" | "group" | number;

export type SparklineItemInput = { data: string; location: string };

export type SparklineGroupOptions = Partial<Record<(typeof FLAGS)[number], boolean>> & {
  type?: SparklineTypeName;
  colors?: Partial<Record<SparklineColorName, string>>;
  lineWeight?: number;
  emptyCells?: SparklineEmptyName;
  minAxis?: SparklineAxisBound;
  maxAxis?: SparklineAxisBound;
  dateAxis?: string;
};

type Session = ReturnType<typeof getSession>;
type GroupSpec = { -readonly [K in keyof SparklineGroupPatch]: SparklineGroupPatch[K] };

function quotedSheet(name: string): string {
  return `'${name.replaceAll("'", "''")}'`;
}

/** Resolves one sparkline: its data range (a bare one resolves on the host sheet) and its cell. */
function sparklineItem(item: SparklineItemInput) {
  const source = item.data.trim().replace(/^=/, "");
  parseRangeRef(source);
  const location = parseCellRef(item.location);
  if (location.sheetName !== undefined) {
    throw new Error(`sparkline location must be a cell on the selected sheet: ${item.location}`);
  }
  assertInGrid(location.row, location.col);
  return { source, location: cellToA1(location.row, location.col) };
}

/** The sheet-qualified A1 text of a range on a workbook sheet. */
function qualifiedRange(session: Session, range: SheetRect): string {
  return `${quotedSheet(sheetName(session, range.sheet))}!${rectA1(range.rect)}`;
}

function colorSpec(spec: string): Color {
  return { kind: 1, argb: parseColor(spec), theme: 0, tint: 0, indexed: 0 };
}

function describeColor(color: Color | null): string | { theme: number; tint: number } | undefined {
  if (color === null) {
    return undefined;
  }
  if (color.kind === 1) {
    return formatColor(color.argb);
  }
  return color.kind === 2 ? { theme: color.theme, tint: color.tint } : undefined;
}

function axisSpec(
  bound: SparklineAxisBound | undefined,
  label: string,
): { type: SparklineAxisType; manual: number | null } | undefined {
  if (bound === undefined) {
    return undefined;
  }
  if (typeof bound === "number") {
    if (!Number.isFinite(bound)) {
      throw new Error(`${label} must be finite: ${bound}`);
    }
    return { type: SparklineAxisType.Custom, manual: bound };
  }
  return {
    type: bound === "group" ? SparklineAxisType.Group : SparklineAxisType.Individual,
    manual: null,
  };
}

function groupSpec(options: SparklineGroupOptions) {
  const spec: GroupSpec = {};
  if (options.type !== undefined) {
    spec.type = TYPE_CODES[options.type];
  }
  for (const flag of FLAGS) {
    if (options[flag] !== undefined) {
      spec[flag] = options[flag];
    }
  }
  for (const [name, color] of Object.entries(options.colors ?? {})) {
    if (color !== undefined) {
      spec[COLOR_FIELDS[name as SparklineColorName]] = colorSpec(color);
    }
  }
  if (options.lineWeight !== undefined) {
    spec.lineWidthPt = options.lineWeight;
  }
  if (options.emptyCells !== undefined) {
    spec.emptyCellsAs = EMPTY_CODES[options.emptyCells];
  }
  const min = axisSpec(options.minAxis, "minAxis");
  if (min) {
    spec.minAxisType = min.type;
    spec.manualMin = min.manual;
  }
  const max = axisSpec(options.maxAxis, "maxAxis");
  if (max) {
    spec.maxAxisType = max.type;
    spec.manualMax = max.manual;
  }
  if (options.dateAxis !== undefined) {
    const dateRange = options.dateAxis.trim().replace(/^=/, "");
    spec.dateAxis = dateRange !== "";
    if (dateRange !== "") {
      parseRangeRef(dateRange);
    }
    spec.dateAxisRange = dateRange === "" ? null : dateRange;
  }
  return spec;
}

function describeAxis(type: SparklineAxisType, manual: number | null): SparklineAxisBound {
  if (type === SparklineAxisType.Custom && manual !== null) {
    return manual;
  }
  return type === SparklineAxisType.Group ? "group" : "individual";
}

function describeGroup(session: Session, group: SparklineGroup, index: number) {
  const colors: Partial<Record<SparklineColorName, ReturnType<typeof describeColor>>> = {};
  for (const [name, field] of Object.entries(COLOR_FIELDS)) {
    const color = describeColor(group[field]);
    if (color !== undefined) {
      colors[name as SparklineColorName] = color;
    }
  }
  return {
    index,
    type: TYPE_NAMES[group.type] ?? "line",
    sparklines: group.sparklines.map(({ source, location }) => ({
      data: source === null ? null : qualifiedRange(session, source),
      location: cellToA1(location.row, location.col),
    })),
    ...Object.fromEntries(FLAGS.filter((flag) => group[flag]).map((flag) => [flag, true])),
    ...(Object.keys(colors).length > 0 ? { colors } : {}),
    lineWeight: group.lineWidthPt,
    emptyCells: EMPTY_NAMES.find((name) => EMPTY_CODES[name] === group.emptyCellsAs) ?? "gap",
    minAxis: describeAxis(group.minAxisType, group.manualMin),
    maxAxis: describeAxis(group.maxAxisType, group.manualMax),
    ...(group.dateAxis && group.dateAxisRange
      ? { dateAxis: qualifiedRange(session, group.dateAxisRange) }
      : {}),
  };
}

function groups(session: Session, sheetIndex: number) {
  return sheetAt(session.workbook, sheetIndex).listSparklineGroups();
}

function groupAt(session: Session, sheetIndex: number, index: number) {
  const group = groups(session, sheetIndex)[index];
  if (!group) {
    throw new Error(`no sparkline group ${index} on sheet ${sheetIndex}`);
  }
  return describeGroup(session, group, index);
}

/** Lists a sheet's sparkline groups in document order with their sparklines. */
export function listSessionSparklines(id: string, sheet: number | string | undefined) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  const value = groups(session, sheetIndex).map((group, index) =>
    describeGroup(session, group, index),
  );
  touch(session, false);
  return {
    session: publicInfo(session),
    sheet: sheetIndex,
    sheetName: sheetName(session, sheetIndex),
    count: value.length,
    groups: value,
  };
}

/** Appends a sparkline group; every sparkline in it shares the group's type, colours and axes. */
export function addSessionSparklines(
  id: string,
  sheet: number | string | undefined,
  sparklines: SparklineItemInput[],
  options: SparklineGroupOptions,
) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  if (sparklines.length === 0) {
    throw new Error("a sparkline group needs at least one sparkline");
  }
  const spec = {
    ...groupSpec(options),
    sparklines: sparklines.map(sparklineItem),
  };
  touch(session, true);
  sheetAt(session.workbook, sheetIndex).addSparklineGroup(spec);
  const index = groups(session, sheetIndex).length - 1;
  return {
    session: publicInfo(session),
    sheet: sheetIndex,
    group: groupAt(session, sheetIndex, index),
  };
}

/** Rewrites the stated group settings; `sparklines`, when given, replaces the group's sparklines. */
export function updateSessionSparklines(
  id: string,
  sheet: number | string | undefined,
  index: number,
  sparklines: SparklineItemInput[] | undefined,
  options: SparklineGroupOptions,
) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  if (!groups(session, sheetIndex)[index]) {
    throw new Error(`no sparkline group ${index} on sheet ${sheetIndex}`);
  }
  const spec = groupSpec(options);
  const items = sparklines?.map(sparklineItem);
  if (items?.length === 0) {
    throw new Error("a sparkline group needs at least one sparkline");
  }
  if (Object.keys(spec).length === 0 && items === undefined) {
    throw new Error("sparkline update requires at least one setting or sparklines");
  }
  touch(session, true);
  sheetAt(session.workbook, sheetIndex).updateSparklineGroup(index, {
    ...spec,
    ...(items === undefined ? {} : { sparklines: items }),
  });
  return {
    session: publicInfo(session),
    sheet: sheetIndex,
    group: groupAt(session, sheetIndex, index),
  };
}

/** Removes a sparkline group; later groups shift down one index. */
export function removeSessionSparklines(
  id: string,
  sheet: number | string | undefined,
  index: number,
) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  const removed = groupAt(session, sheetIndex, index);
  touch(session, true);
  sheetAt(session.workbook, sheetIndex).removeSparklineGroup(index);
  touch(session, true);
  return { session: publicInfo(session), sheet: sheetIndex, removed };
}
