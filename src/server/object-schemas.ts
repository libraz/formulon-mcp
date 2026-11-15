import {
  CfDataBarAxisPosition,
  CfDataBarDirection,
  CfRuleType,
  CfValueObjectType,
  ComparisonOperator,
  type ConditionalFormatInput,
  type DataValidationInput,
  DatePeriod,
  enumName,
  IconSetName,
  ValidationErrorStyle,
  ValidationType,
} from "@libraz/formulon";
import { z } from "zod";
import { StatusCode, statusError } from "../formulon.js";

type EnumTable = Readonly<Record<string, number>>;
/** Where an enum value sits in an input record: a key holding a table's member name, a nested record or a list of them. */
type EnumSpec = { readonly [key: string]: EnumTable | EnumSpec | readonly [EnumSpec] };

const names = (table: EnumTable) => Object.keys(table).join(", ");

const isTable = (spec: EnumTable | EnumSpec | readonly [EnumSpec]): spec is EnumTable =>
  !Array.isArray(spec) && Object.values(spec).every((value) => typeof value === "number");

/**
 * Replaces each enum member name in `value` with the ordinal the engine takes, at the places
 * `spec` names. A member is given by name only, so an ordinal or an unknown name is refused with
 * the list of names; every other key passes through for the engine to check.
 */
function enumOrdinals(value: unknown, spec: EnumSpec, path: string): unknown {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return value;
  }
  const out: Record<string, unknown> = { ...(value as Record<string, unknown>) };
  for (const [key, inner] of Object.entries(spec)) {
    const field = out[key];
    if (field === undefined || field === null) {
      continue;
    }
    const at = `${path}.${key}`;
    if (Array.isArray(inner)) {
      if (Array.isArray(field)) {
        out[key] = field.map((item, index) => enumOrdinals(item, inner[0], `${at}[${index}]`));
      }
    } else if (isTable(inner)) {
      if (typeof field !== "string" || !Object.hasOwn(inner, field)) {
        throw statusError(StatusCode.InvalidArgument, `${at} must be one of ${names(inner)}`);
      }
      out[key] = inner[field];
    } else {
      out[key] = enumOrdinals(field, inner as EnumSpec, at);
    }
  }
  return out;
}

/**
 * The inverse of {@link enumOrdinals}: replaces each ordinal in an engine record with its member
 * name at the places `spec` names, so a listed record is the shape the matching add takes. A null
 * field and an ordinal without a member pass through.
 */
function enumNames(value: unknown, spec: EnumSpec): unknown {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return value;
  }
  const out: Record<string, unknown> = { ...(value as Record<string, unknown>) };
  for (const [key, inner] of Object.entries(spec)) {
    const field = out[key];
    if (field === undefined || field === null) {
      continue;
    }
    if (Array.isArray(inner)) {
      if (Array.isArray(field)) {
        out[key] = field.map((item) => enumNames(item, inner[0]));
      }
    } else if (isTable(inner)) {
      if (typeof field === "number") {
        out[key] = enumName(inner, field) ?? field;
      }
    } else {
      out[key] = enumNames(field, inner as EnumSpec);
    }
  }
  return out;
}

const valueObject: EnumSpec = { type: CfValueObjectType };

const VALIDATION_ENUMS: EnumSpec = {
  type: ValidationType,
  operator: ComparisonOperator,
  errorStyle: ValidationErrorStyle,
};

const CF_ENUMS: EnumSpec = {
  type: CfRuleType,
  operator: ComparisonOperator,
  timePeriod: DatePeriod,
  colorScale: { thresholds: [valueObject] },
  dataBar: {
    min: valueObject,
    max: valueObject,
    axisPosition: CfDataBarAxisPosition,
    direction: CfDataBarDirection,
  },
  iconSet: {
    name: IconSetName,
    thresholds: [valueObject],
    floor: valueObject,
    icons: [{ iconSet: IconSetName }],
  },
};

/** The engine's `DataValidationInput` for a tool's validation record. */
export function validationInput(validation: Record<string, unknown>): DataValidationInput {
  return enumOrdinals(validation, VALIDATION_ENUMS, "validation") as DataValidationInput;
}

/** The engine's `ConditionalFormatInput` for a tool's rule record. */
export function conditionalFormatInput(rule: Record<string, unknown>): ConditionalFormatInput {
  return enumOrdinals(rule, CF_ENUMS, "rule") as ConditionalFormatInput;
}

/** A validation record as the tool's `validation` input takes it: enum fields by member name. */
export function validationRecord(validation: unknown): unknown {
  return enumNames(validation, VALIDATION_ENUMS);
}

/** A conditional-format record as the tool's `rule` input takes it: enum fields by member name. */
export function conditionalFormatRecord(rule: unknown): unknown {
  return enumNames(rule, CF_ENUMS);
}

const ranges = z
  .union([z.string(), z.array(z.union([z.string(), z.object({}).passthrough()]))])
  .describe(
    "The cells the record covers: A1 text ('A2:A20', several separated by spaces) or a list of A1 ranges / {firstRow, firstCol, lastRow, lastCol} rectangles.",
  );

/**
 * A data validation as the engine's `DataValidationInput`, with its enum fields given by member
 * name. The engine checks every field; an unknown key or a wrong value is the tool error.
 */
export const validationSchema = z
  .object({
    ranges,
    type: z.string().describe(`ValidationType member: ${names(ValidationType)}.`),
    operator: z
      .string()
      .nullish()
      .describe(
        `ComparisonOperator member for whole, decimal, date, time and textLength: ${names(ComparisonOperator)}.`,
      ),
    errorStyle: z
      .string()
      .nullish()
      .describe(`ValidationErrorStyle member: ${names(ValidationErrorStyle)}.`),
  })
  .passthrough()
  .describe(
    "DataValidationInput (see formulon_workbook_api describe Sheet.addValidation): allowBlank, showInputMessage, showErrorMessage, showInCellDropdown (each default true), formula1, formula2, errorTitle, errorMessage, promptTitle, promptMessage. A literal list formula uses Excel quotes, for example '\"Draft,Approved\"'.",
  );

/**
 * A conditional-format rule as the engine's `ConditionalFormatInput`, with its enum fields given
 * by member name. The engine checks every field of the rule kind `type` selects.
 */
export const conditionalFormatSchema = z
  .object({
    ranges,
    type: z.string().describe(`CfRuleType member: ${names(CfRuleType)}.`),
    operator: z
      .string()
      .nullish()
      .describe(`CellIs: ComparisonOperator member: ${names(ComparisonOperator)}.`),
    timePeriod: z
      .string()
      .nullish()
      .describe(`TimePeriod: DatePeriod member: ${names(DatePeriod)}.`),
  })
  .passthrough()
  .describe(
    `ConditionalFormatInput (see formulon_workbook_api describe Sheet.addCfRule). Common keys: priority, stopIfTrue, id, and format (a CellFormatPatch) or dxfIndex for the rules that format cells; formula1/formula2, text, rank/percent/bottom, aboveAverage/equalAverage/stdDev by kind. colorScale {thresholds, colors}, dataBar {min, max, fill, ...} and iconSet {name, thresholds, ...} carry thresholds {type: CfValueObjectType member (${names(CfValueObjectType)}), value}; colours are '#RRGGBB' or Color records; iconSet.name is an IconSetName member (${names(IconSetName)}).`,
  );
