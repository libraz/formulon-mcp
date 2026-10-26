import type { ColorSpec, SparklineGroup, SparklineGroupInput } from "@libraz/formulon";
import { cellToA1, parseCellRef, parseRangeRef } from "../a1.js";
import {
  assertStatus,
  findSheetIndex,
  requireListResult,
  SparklineAxisType,
  SparklineEmptyCells,
  SparklineType,
} from "../formulon.js";
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
const EMPTY_CODES: Record<SparklineEmptyName, SparklineEmptyCells> = {
  gap: SparklineEmptyCells.Gap,
  zero: SparklineEmptyCells.Zero,
  span: SparklineEmptyCells.Span,
};
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

function quotedSheet(name: string): string {
  return `'${name.replaceAll("'", "''")}'`;
}

/** Resolves one sparkline: its data range (qualified with the host sheet when bare) and its cell. */
function sparklineItem(session: Session, sheetIndex: number, item: SparklineItemInput) {
  const data = item.data.trim().replace(/^=/, "");
  const range = parseRangeRef(data);
  const location = parseCellRef(item.location);
  if (location.sheetName !== undefined) {
    throw new Error(`sparkline location must be a cell on the selected sheet: ${item.location}`);
  }
  assertInGrid(location.row, location.col);
  return {
    formula:
      range.sheetName === undefined
        ? `${quotedSheet(sheetName(session, sheetIndex))}!${data}`
        : data,
    location: cellToA1(location.row, location.col),
  };
}

function colorSpec(spec: string): ColorSpec {
  return { kind: 1, rgb: parseColor(spec), theme: 0, tint: 0, indexed: 0 };
}

function describeColor(color: ColorSpec): string | { theme: number; tint: number } | undefined {
  if (color.kind === 1) {
    return formatColor(color.rgb);
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

function groupSpec(session: Session, sheetIndex: number, options: SparklineGroupOptions) {
  const spec: SparklineGroupInput = {};
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
    spec.lineWeight = options.lineWeight;
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
    spec.dateAxisRef =
      dateRange === "" || parseRangeRef(dateRange).sheetName !== undefined
        ? dateRange
        : `${quotedSheet(sheetName(session, sheetIndex))}!${dateRange}`;
  }
  return spec;
}

function describeAxis(type: SparklineAxisType, manual: number | null): SparklineAxisBound {
  if (type === SparklineAxisType.Custom && manual !== null) {
    return manual;
  }
  return type === SparklineAxisType.Group ? "group" : "individual";
}

function describeGroup(group: SparklineGroup, index: number) {
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
    sparklines: group.sparklines.map(({ formula, location }) => ({ data: formula, location })),
    ...Object.fromEntries(FLAGS.filter((flag) => group[flag]).map((flag) => [flag, true])),
    ...(Object.keys(colors).length > 0 ? { colors } : {}),
    lineWeight: group.lineWeight,
    emptyCells: EMPTY_NAMES[group.emptyCellsAs] ?? "gap",
    minAxis: describeAxis(group.minAxisType, group.manualMin),
    maxAxis: describeAxis(group.maxAxisType, group.manualMax),
    ...(group.dateAxis ? { dateAxis: group.dateAxisRef } : {}),
  };
}

function groups(session: Session, sheetIndex: number) {
  return requireListResult(
    session.workbook.listSparklineGroups(sheetIndex),
    "list sparkline groups",
  );
}

function groupAt(session: Session, sheetIndex: number, index: number) {
  const group = groups(session, sheetIndex)[index];
  if (!group) {
    throw new Error(`no sparkline group ${index} on sheet ${sheetIndex}`);
  }
  return describeGroup(group, index);
}

/** Lists a sheet's sparkline groups in document order with their sparklines. */
export function listSessionSparklines(id: string, sheet: number | string | undefined) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  const value = groups(session, sheetIndex).map(describeGroup);
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
  const spec: SparklineGroupInput = {
    ...groupSpec(session, sheetIndex, options),
    sparklines: sparklines.map((item) => sparklineItem(session, sheetIndex, item)),
  };
  touch(session, true);
  assertStatus(session.workbook.addSparklineGroup(sheetIndex, spec), "add sparkline group");
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
  const current = groups(session, sheetIndex)[index];
  if (!current) {
    throw new Error(`no sparkline group ${index} on sheet ${sheetIndex}`);
  }
  const spec = groupSpec(session, sheetIndex, options);
  const items = sparklines?.map((item) => sparklineItem(session, sheetIndex, item));
  if (items?.length === 0) {
    throw new Error("a sparkline group needs at least one sparkline");
  }
  if (Object.keys(spec).length === 0 && items === undefined) {
    throw new Error("sparkline update requires at least one setting or sparklines");
  }
  touch(session, true);
  if (Object.keys(spec).length > 0) {
    // An update rewrites every modelled field, so the unstated ones are carried over.
    const { sparklines: _kept, ...settings } = current;
    assertStatus(
      session.workbook.updateSparklineGroup(sheetIndex, index, { ...settings, ...spec }),
      "update sparkline group",
    );
  }
  if (items !== undefined) {
    assertStatus(
      session.workbook.setSparklineGroupSparklines(sheetIndex, index, items),
      "set sparklines",
    );
  }
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
  assertStatus(session.workbook.removeSparklineGroup(sheetIndex, index), "remove sparkline group");
  touch(session, true);
  return { session: publicInfo(session), sheet: sheetIndex, removed };
}
