import assert from "node:assert/strict";
import { test } from "vitest";
import { callTool, textPayload, withClient } from "./helpers/mcp.js";

/** Calls a tool and returns its parsed payload. */
const caller = (client) => async (name, args) => textPayload(await callTool(client, name, args));

/** `record` without the keys that identify it or place it. */
const without = (record, keys) =>
  Object.fromEntries(Object.entries(record).filter(([key]) => !keys.includes(key)));

test("a listed validation is added back by its member names", async () => {
  await withClient(async (client) => {
    const call = caller(client);
    const sessionId = "replay-validation";
    await call("formulon_open_workbook", { sessionId });
    await call("formulon_validation_operation", {
      sessionId,
      operation: "add",
      validation: {
        ranges: "A1:A5",
        type: "Whole",
        operator: "NotBetween",
        errorStyle: "Warning",
        formula1: "1",
        formula2: "10",
      },
    });
    const first = (await call("formulon_validation_operation", { sessionId, operation: "list" }))
      .result[0];
    assert.equal(first.type, "Whole");
    assert.equal(first.operator, "NotBetween");
    assert.equal(first.errorStyle, "Warning");

    await call("formulon_validation_operation", {
      sessionId,
      operation: "add",
      validation: { ...without(first, ["id"]), ranges: "C1:C5" },
    });
    const listed = (await call("formulon_validation_operation", { sessionId, operation: "list" }))
      .result;
    assert.equal(listed.length, 2);
    const identity = ["id", "ranges"];
    assert.deepEqual(without(listed[1], identity), without(listed[0], identity));
  });
});

test("a listed conditional-format rule is added back by its member names", async () => {
  await withClient(async (client) => {
    const call = caller(client);
    const sessionId = "replay-cf";
    await call("formulon_open_workbook", { sessionId });
    const rules = [
      { type: "CellIs", operator: "Between", formula1: "1", formula2: "5" },
      { type: "Expression", formula1: "=A1>0" },
      { type: "TimePeriod", timePeriod: "LastWeek" },
      {
        type: "ColorScale",
        colorScale: {
          thresholds: [{ type: "Min" }, { type: "Percentile", value: "50" }, { type: "Max" }],
          colors: ["#FF0000", "#FFFF00", "#00FF00"],
        },
      },
      {
        type: "DataBar",
        dataBar: {
          min: { type: "Number", value: "0" },
          max: { type: "Max" },
          fill: "#FF0064C8",
          axisPosition: "Middle",
          direction: "RightToLeft",
        },
      },
      {
        type: "IconSet",
        iconSet: {
          name: "ThreeArrows",
          floor: { type: "Number", value: "0" },
          thresholds: [
            { type: "Number", value: "50" },
            { type: "Percent", value: "80" },
          ],
        },
      },
    ];
    for (const [index, rule] of rules.entries()) {
      const range = `A${index * 3 + 1}:A${index * 3 + 3}`;
      await call("formulon_conditional_format_operation", {
        sessionId,
        operation: "add",
        rule: { ...rule, ranges: range, format: { font: { bold: true } } },
      });
    }
    const before = (
      await call("formulon_conditional_format_operation", { sessionId, operation: "list" })
    ).result;
    assert.equal(before.length, rules.length);
    assert.deepEqual(
      before.map((rule) => rule.type),
      rules.map((rule) => rule.type),
    );

    for (const [index, rule] of before.entries()) {
      await call("formulon_conditional_format_operation", {
        sessionId,
        operation: "add",
        rule: {
          ...without(rule, ["id", "priority"]),
          ranges: `C${index * 3 + 1}:C${index * 3 + 3}`,
        },
      });
    }
    const after = (
      await call("formulon_conditional_format_operation", { sessionId, operation: "list" })
    ).result;
    assert.equal(after.length, rules.length * 2);
    const identity = ["id", "priority", "ranges"];
    for (const [index, rule] of before.entries()) {
      assert.deepEqual(without(after[rules.length + index], identity), without(rule, identity));
    }
  });
});
