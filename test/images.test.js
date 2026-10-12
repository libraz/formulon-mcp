import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { test } from "vitest";
import { callTool, errorPayload, textPayload, withClient } from "./helpers/mcp.js";
import { solidPng } from "./helpers/png.js";

test("inserts, lists, reads, and removes pictures through a save round trip", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-images-"));
  try {
    const logo = path.join(dir, "logo.png");
    const png = solidPng(40, 20);
    await writeFile(logo, png);
    const saved = path.join(dir, "images.xlsx");

    await withClient(async (client) => {
      const sessionId = "images";
      textPayload(await callTool(client, "formulon_open_workbook", { sessionId }));
      textPayload(
        await callTool(client, "formulon_set_cells", {
          sessionId,
          mutations: [{ type: "formula", sheet: 0, row: 0, col: 0, formula: "=1+1" }],
        }),
      );

      const first = textPayload(
        await callTool(client, "formulon_image_operation", {
          sessionId,
          operation: "insert",
          path: logo,
          anchor: "B2",
          offsetX: { size: 4, unit: "px" },
          width: { size: 60, unit: "pt" },
          description: "Company logo",
        }),
      );
      assert.deepEqual(first.source, { path: logo, format: "png", pxWidth: 40, pxHeight: 20 });
      assert.equal(first.placed.from, "B2");
      assert.equal(first.placed.offset.x.px, 4);
      // Height follows the source aspect ratio when only the width is given.
      assert.deepEqual(first.placed.width, { pt: 60, px: 80 });
      assert.deepEqual(first.placed.height, { pt: 30, px: 40 });
      assert.equal(first.placed.description, "Company logo");

      const second = textPayload(
        await callTool(client, "formulon_image_operation", {
          sessionId,
          operation: "insert",
          path: logo,
          anchor: "E5",
        }),
      );
      assert.deepEqual(second.placed.width, { pt: 30, px: 40 });

      const got = await callTool(client, "formulon_image_operation", {
        sessionId,
        operation: "get",
        objectId: first.objectId,
        outputPath: path.join(dir, "copy.png"),
      });
      assert.equal(got.content[0].type, "image");
      assert.equal(got.content[0].mimeType, "image/png");
      assert.deepEqual(Buffer.from(got.content[0].data, "base64"), png);
      assert.deepEqual(await readFile(path.join(dir, "copy.png")), png);

      textPayload(
        await callTool(client, "formulon_image_operation", {
          sessionId,
          operation: "remove",
          objectId: second.objectId,
        }),
      );
      textPayload(
        await callTool(client, "formulon_save_session", { sessionId, outputPath: saved }),
      );
      textPayload(await callTool(client, "formulon_close_workbook", { sessionId }));

      textPayload(
        await callTool(client, "formulon_open_workbook", { sessionId: "reopened", path: saved }),
      );
      const listed = textPayload(
        await callTool(client, "formulon_image_operation", {
          sessionId: "reopened",
          operation: "list",
        }),
      );
      assert.equal(listed.count, 1);
      assert.equal(listed.objects[0].kind, "picture");
      assert.equal(listed.objects[0].from, "B2");
      assert.equal(listed.objects[0].description, "Company logo");
      assert.deepEqual(listed.objects[0].width, { pt: 60, px: 80 });
      const cell = textPayload(
        await callTool(client, "formulon_get_cell", { sessionId: "reopened", a1: "A1" }),
      );
      assert.equal(cell.formula, "=1+1");
    });
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("moves, resizes, and restacks a picture in place", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-images-"));
  try {
    const logo = path.join(dir, "logo.png");
    await writeFile(logo, solidPng(40, 20));
    await withClient(async (client) => {
      const sessionId = "image-move";
      textPayload(await callTool(client, "formulon_open_workbook", { sessionId }));
      const insert = (anchor, extra = {}) =>
        callTool(client, "formulon_image_operation", {
          sessionId,
          operation: "insert",
          path: logo,
          anchor,
          ...extra,
        }).then(textPayload);
      const first = await insert("B2", { anchorKind: "twoCell", description: "Logo" });
      const second = await insert("E5");

      const moved = textPayload(
        await callTool(client, "formulon_image_operation", {
          sessionId,
          operation: "move",
          objectId: first.objectId,
          anchor: "D4",
          width: { size: 120, unit: "pt" },
        }),
      );
      assert.equal(moved.placed.objectId, first.objectId);
      assert.equal(moved.placed.from, "D4");
      // An unstated kind, alt text and aspect ratio are kept.
      assert.equal(moved.placed.anchorKind, "twoCell");
      assert.equal(moved.placed.description, "Logo");
      assert.deepEqual(moved.placed.width, { pt: 120, px: 160 });
      assert.deepEqual(moved.placed.height, { pt: 60, px: 80 });

      const nudged = textPayload(
        await callTool(client, "formulon_image_operation", {
          sessionId,
          operation: "move",
          objectId: first.objectId,
          offsetY: { size: 10, unit: "pt" },
        }),
      );
      assert.equal(nudged.placed.from, "D4");
      assert.equal(nudged.placed.offset.y.pt, 10);
      assert.deepEqual(nudged.placed.width, { pt: 120, px: 160 });

      const reordered = textPayload(
        await callTool(client, "formulon_image_operation", {
          sessionId,
          operation: "reorder",
          objectId: second.objectId,
          index: 0,
        }),
      );
      assert.deepEqual(reordered.order, [second.objectId, first.objectId]);

      assert.match(
        errorPayload(
          await callTool(client, "formulon_image_operation", {
            sessionId,
            operation: "reorder",
            objectId: first.objectId,
          }),
        ),
        /index is required/,
      );
      assert.match(
        errorPayload(
          await callTool(client, "formulon_image_operation", {
            sessionId,
            operation: "move",
            objectId: 999,
            anchor: "A1",
          }),
        ),
        /no drawing object 999/,
      );
    });
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("rejects image operations missing their required arguments", async () => {
  await withClient(async (client) => {
    const sessionId = "image-errors";
    textPayload(await callTool(client, "formulon_open_workbook", { sessionId }));
    assert.match(
      errorPayload(
        await callTool(client, "formulon_image_operation", { sessionId, operation: "insert" }),
      ),
      /path and anchor are required/,
    );
    assert.match(
      errorPayload(
        await callTool(client, "formulon_image_operation", { sessionId, operation: "remove" }),
      ),
      /objectId is required/,
    );
    assert.match(
      errorPayload(
        await callTool(client, "formulon_image_operation", {
          sessionId,
          operation: "insert",
          path: "logo.png",
          anchor: "Sheet1!B2",
        }),
      ),
      /without a sheet name/,
    );
  });
});

// ---------------------------------------------------------------------------
// Moving and restacking pictures
// ---------------------------------------------------------------------------

const POINTS_PER_UNIT = { pt: 1, px: 0.75, mm: 72 / 25.4 };
const toPoints = ({ size, unit = "pt" }) => size * POINTS_PER_UNIT[unit];
const STORED_PT_TOLERANCE = 0.011;

/**
 * Pairwise matrix over the move arguments (anchor x offsetX x offsetY x width
 * x height x anchorKind x unit) and the picture's starting state; the first 13
 * rows cover every value pair.
 */
const MOVE_CASES = [
  ["absent", "absent", "present", "absent", "absent", "absent", "pt", "twoCell", "proportional"],
  [
    "present",
    "present",
    "absent",
    "present",
    "present",
    "oneCell",
    "mm",
    "oneCell",
    "proportional",
  ],
  ["present", "present", "absent", "present", "absent", "absent", "px", "twoCell", "stretched"],
  ["present", "absent", "absent", "absent", "present", "twoCell", "pt", "oneCell", "stretched"],
  ["absent", "present", "present", "absent", "absent", "twoCell", "px", "oneCell", "proportional"],
  ["absent", "present", "present", "present", "present", "oneCell", "mm", "twoCell", "stretched"],
  ["absent", "absent", "absent", "present", "absent", "oneCell", "pt", "twoCell", "stretched"],
  ["absent", "absent", "absent", "present", "present", "twoCell", "px", "twoCell", "stretched"],
  ["absent", "absent", "present", "absent", "present", "absent", "mm", "oneCell", "stretched"],
  ["present", "present", "present", "present", "absent", "absent", "pt", "twoCell", "proportional"],
  [
    "present",
    "absent",
    "present",
    "present",
    "present",
    "twoCell",
    "mm",
    "twoCell",
    "proportional",
  ],
  ["absent", "absent", "absent", "absent", "absent", "absent", "mm", "oneCell", "stretched"],
  ["absent", "present", "absent", "absent", "absent", "oneCell", "px", "twoCell", "stretched"],
  // Beyond the pairwise rows: lone width or height on a stretched picture in every unit,
  // and every argument at once.
  ["absent", "absent", "absent", "present", "absent", "absent", "px", "oneCell", "stretched"],
  ["absent", "absent", "absent", "present", "absent", "absent", "mm", "twoCell", "stretched"],
  ["absent", "absent", "absent", "absent", "present", "absent", "pt", "twoCell", "stretched"],
  ["absent", "absent", "absent", "absent", "present", "absent", "px", "oneCell", "stretched"],
  ["absent", "absent", "absent", "absent", "present", "absent", "mm", "oneCell", "proportional"],
  ["present", "present", "present", "present", "present", "twoCell", "px", "oneCell", "stretched"],
  ["present", "present", "present", "present", "present", "oneCell", "pt", "twoCell", "stretched"],
];

const MOVE_VALUES = {
  anchor: "F9",
  // Kept inside one default cell so the engine has no reason to re-express the anchor.
  offsetX: 2,
  offsetY: 1,
  width: 90,
  height: 45,
};

/** Reads the sheet's drawing objects through the tool. */
async function listPictures(client, sessionId) {
  return textPayload(
    await callTool(client, "formulon_image_operation", { sessionId, operation: "list" }),
  ).objects;
}

async function sessionIsDirty(client, sessionId) {
  const { sessions } = textPayload(await callTool(client, "formulon_list_sessions"));
  return sessions.find((session) => session.id === sessionId).dirty;
}

function assertClose(actual, expected, message) {
  assert.ok(
    Math.abs(actual - expected) <= STORED_PT_TOLERANCE,
    `${message}: expected ${expected}pt, got ${actual}pt`,
  );
}

/** The placement fields a move may change, plus the identity fields it must keep. */
function placementOf(object) {
  const { objectId, kind, name, description, anchorKind, from, offset, width, height, format } =
    object;
  return { objectId, kind, name, description, anchorKind, from, offset, width, height, format };
}

async function runMoveCase(client, dir, logo, row, index) {
  const [anchor, offsetX, offsetY, width, height, kindKey, unit, initialKind, shape] = row;
  const sessionId = `image-move-matrix-${index}`;
  const savedPath = path.join(dir, `moved-${index}.xlsx`);
  textPayload(await callTool(client, "formulon_open_workbook", { sessionId }));
  try {
    const inserted = textPayload(
      await callTool(client, "formulon_image_operation", {
        sessionId,
        operation: "insert",
        path: logo,
        anchor: "C3",
        offsetX: { size: 5, unit: "pt" },
        offsetY: { size: 3, unit: "pt" },
        anchorKind: initialKind,
        name: "Logo",
        description: "Company logo",
      }),
    );
    const { objectId } = inserted;
    if (shape === "stretched") {
      // A previous resize that does not keep the source aspect ratio.
      textPayload(
        await callTool(client, "formulon_image_operation", {
          sessionId,
          operation: "move",
          objectId,
          width: { size: 100, unit: "pt" },
          height: { size: 20, unit: "pt" },
        }),
      );
    }
    textPayload(
      await callTool(client, "formulon_save_session", { sessionId, outputPath: savedPath }),
    );
    assert.equal(await sessionIsDirty(client, sessionId), false);

    const before = (await listPictures(client, sessionId)).find(
      (item) => item.objectId === objectId,
    );
    assert.equal(before.anchorKind, initialKind);

    const args = {};
    if (anchor === "present") {
      args.anchor = MOVE_VALUES.anchor;
    }
    if (offsetX === "present") {
      args.offsetX = { size: MOVE_VALUES.offsetX, unit };
    }
    if (offsetY === "present") {
      args.offsetY = { size: MOVE_VALUES.offsetY, unit };
    }
    if (width === "present") {
      args.width = { size: MOVE_VALUES.width, unit };
    }
    if (height === "present") {
      args.height = { size: MOVE_VALUES.height, unit };
    }
    if (kindKey !== "absent") {
      args.anchorKind = kindKey;
    }
    const moved = textPayload(
      await callTool(client, "formulon_image_operation", {
        sessionId,
        operation: "move",
        objectId,
        ...args,
      }),
    );
    assert.equal(moved.objectId, objectId);
    assert.equal(await sessionIsDirty(client, sessionId), true);

    const after = (await listPictures(client, sessionId)).find(
      (item) => item.objectId === objectId,
    );
    assert.deepEqual(placementOf(moved.placed), placementOf(after));

    // Identity is kept whatever moved.
    assert.equal(after.objectId, objectId);
    assert.equal(after.name, "Logo");
    assert.equal(after.description, "Company logo");
    assert.equal(after.kind, "picture");
    assert.equal(after.format, "png");
    assert.equal(after.anchorKind, kindKey === "absent" ? initialKind : kindKey);

    // Position: stated fields change, omitted ones keep their value.
    assert.equal(after.from, anchor === "present" ? MOVE_VALUES.anchor : before.from);
    for (const [axis, key, state] of [
      ["x", "offsetX", offsetX],
      ["y", "offsetY", offsetY],
    ]) {
      if (state === "present") {
        assertClose(after.offset[axis].pt, toPoints({ size: MOVE_VALUES[key], unit }), key);
      } else {
        assert.deepEqual(after.offset[axis], before.offset[axis], `${key} is kept`);
      }
    }

    // Size: both given is exact, a lone side keeps the current aspect ratio, none keeps both.
    if (width === "present" && height === "present") {
      assertClose(after.width.pt, toPoints({ size: MOVE_VALUES.width, unit }), "width");
      assertClose(after.height.pt, toPoints({ size: MOVE_VALUES.height, unit }), "height");
    } else if (width === "present") {
      assertClose(after.width.pt, toPoints({ size: MOVE_VALUES.width, unit }), "width");
      assertClose(
        after.height.pt,
        (after.width.pt * before.height.pt) / before.width.pt,
        "height follows the current ratio",
      );
    } else if (height === "present") {
      assertClose(after.height.pt, toPoints({ size: MOVE_VALUES.height, unit }), "height");
      assertClose(
        after.width.pt,
        (after.height.pt * before.width.pt) / before.height.pt,
        "width follows the current ratio",
      );
    } else {
      assert.deepEqual(after.width, before.width);
      assert.deepEqual(after.height, before.height);
    }
    if (shape === "stretched" && (width === "present") !== (height === "present")) {
      // The ratio kept is the stretched one (5:1), not the source image's 2:1.
      assert.ok(Math.abs(after.width.pt / after.height.pt - 5) < 0.02);
    }

    // The moved picture survives a save and reopen unchanged.
    textPayload(
      await callTool(client, "formulon_save_session", { sessionId, outputPath: savedPath }),
    );
    assert.equal(await sessionIsDirty(client, sessionId), false);
    const reopened = `${sessionId}-reopened`;
    textPayload(
      await callTool(client, "formulon_open_workbook", { sessionId: reopened, path: savedPath }),
    );
    try {
      const reloaded = await listPictures(client, reopened);
      assert.equal(reloaded.length, 1);
      assert.deepEqual(placementOf(reloaded[0]), placementOf(after));
    } finally {
      await callTool(client, "formulon_close_workbook", { sessionId: reopened });
    }
  } finally {
    await callTool(client, "formulon_close_workbook", { sessionId });
  }
}

test("moves keep every unstated placement field across the argument matrix", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-image-matrix-"));
  try {
    const logo = path.join(dir, "logo.png");
    await writeFile(logo, solidPng(40, 20));
    await withClient(async (client) => {
      for (const [index, row] of MOVE_CASES.entries()) {
        try {
          await runMoveCase(client, dir, logo, row, index);
        } catch (error) {
          error.message = `[${row.join(" / ")}] ${error.message}`;
          throw error;
        }
      }
    });
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("a move with no placement arguments changes nothing but still counts as an edit", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-image-noop-"));
  try {
    const logo = path.join(dir, "logo.png");
    await writeFile(logo, solidPng(40, 20));
    await withClient(async (client) => {
      const sessionId = "image-noop";
      textPayload(await callTool(client, "formulon_open_workbook", { sessionId }));
      const { objectId } = textPayload(
        await callTool(client, "formulon_image_operation", {
          sessionId,
          operation: "insert",
          path: logo,
          anchor: "B2",
          anchorKind: "twoCell",
          name: "Keep",
          description: "Kept alt text",
        }),
      );
      const before = await listPictures(client, sessionId);
      textPayload(
        await callTool(client, "formulon_image_operation", {
          sessionId,
          operation: "move",
          objectId,
        }),
      );
      assert.deepEqual(await listPictures(client, sessionId), before);

      // A zero offset is a stated value; it resets the offset rather than keeping it.
      textPayload(
        await callTool(client, "formulon_image_operation", {
          sessionId,
          operation: "move",
          objectId,
          offsetX: { size: 20, unit: "pt" },
          offsetY: { size: 10, unit: "pt" },
        }),
      );
      const offset = await callTool(client, "formulon_image_operation", {
        sessionId,
        operation: "move",
        objectId,
        offsetX: { size: 0 },
      });
      const placed = textPayload(offset).placed;
      assert.equal(placed.offset.x.pt, 0);
      assert.equal(placed.offset.y.pt, 10);
    });
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("rejects a move whose anchor is not a cell on the selected sheet", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-image-anchor-"));
  try {
    const logo = path.join(dir, "logo.png");
    await writeFile(logo, solidPng(40, 20));
    await withClient(async (client) => {
      const sessionId = "image-anchor";
      textPayload(await callTool(client, "formulon_open_workbook", { sessionId }));
      const { objectId } = textPayload(
        await callTool(client, "formulon_image_operation", {
          sessionId,
          operation: "insert",
          path: logo,
          anchor: "B2",
        }),
      );
      const before = await listPictures(client, sessionId);
      const move = (args) =>
        callTool(client, "formulon_image_operation", {
          sessionId,
          operation: "move",
          objectId,
          ...args,
        });
      for (const [anchor, pattern] of [
        ["Sheet1!A1", /without a sheet name/],
        ["'Sheet1'!A1", /without a sheet name/],
        ["XFE1", /exceeds Excel sheet limits/],
        ["A1048577", /exceeds Excel sheet limits/],
        ["XFD1048577", /exceeds Excel sheet limits/],
        ["zz", /invalid A1 cell reference/],
        ["A:A", /invalid A1 cell reference/],
        ["1", /invalid A1 cell reference/],
        ["", /invalid A1 cell reference/],
      ]) {
        assert.match(errorPayload(await move({ anchor, width: { size: 99 } })), pattern, anchor);
      }
      // Nothing was applied by any refused move, size arguments included.
      assert.deepEqual(await listPictures(client, sessionId), before);

      // The last cell is a valid address, but a picture anchored there would run off the sheet.
      assert.match(errorPayload(await move({ anchor: "XFD1048576" })), /move image failed/);
      assert.deepEqual(await listPictures(client, sessionId), before);
      assert.equal(textPayload(await move({ anchor: "AAA1000" })).placed.from, "AAA1000");
      // Lower-case and dollar-marked cells are accepted spellings.
      assert.equal(textPayload(await move({ anchor: "$c$4" })).placed.from, "C4");
    });
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("a move finds objects on the selected sheet only", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-image-sheets-"));
  try {
    const logo = path.join(dir, "logo.png");
    await writeFile(logo, solidPng(40, 20));
    await withClient(async (client) => {
      const sessionId = "image-sheets";
      textPayload(await callTool(client, "formulon_open_workbook", { sessionId }));
      textPayload(
        await callTool(client, "formulon_sheet_operation", {
          sessionId,
          operation: "add",
          name: "Second",
        }),
      );
      const { objectId } = textPayload(
        await callTool(client, "formulon_image_operation", {
          sessionId,
          sheet: "Second",
          operation: "insert",
          path: logo,
          anchor: "B2",
        }),
      );
      for (const args of [
        { objectId: 999 },
        { objectId: objectId + 100 },
        { objectId, sheet: 0 },
        { objectId, sheet: "Sheet1" },
      ]) {
        assert.match(
          errorPayload(
            await callTool(client, "formulon_image_operation", {
              sessionId,
              operation: "move",
              anchor: "A1",
              ...args,
            }),
          ),
          /no drawing object \d+ on sheet 0/,
        );
      }
      const moved = textPayload(
        await callTool(client, "formulon_image_operation", {
          sessionId,
          sheet: "Second",
          operation: "move",
          objectId,
          anchor: "D5",
        }),
      );
      assert.equal(moved.sheet, 1);
      assert.equal(moved.placed.from, "D5");
    });
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("reorders pictures to the front, the back and the middle", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-image-order-"));
  try {
    const logo = path.join(dir, "logo.png");
    await writeFile(logo, solidPng(40, 20));
    await withClient(async (client) => {
      const sessionId = "image-order";
      textPayload(await callTool(client, "formulon_open_workbook", { sessionId }));
      const ids = [];
      for (const anchor of ["B2", "D4", "F6"]) {
        ids.push(
          textPayload(
            await callTool(client, "formulon_image_operation", {
              sessionId,
              operation: "insert",
              path: logo,
              anchor,
              name: `Pic ${anchor}`,
            }),
          ).objectId,
        );
      }
      const [a, b, c] = ids;
      const orderOf = async () =>
        (await listPictures(client, sessionId)).map((item) => item.objectId);
      assert.deepEqual(await orderOf(), [a, b, c]);

      const reorder = async (objectId, index) => {
        await callTool(client, "formulon_save_session", {
          sessionId,
          outputPath: path.join(dir, `order-${objectId}-${index}.xlsx`),
        });
        assert.equal(await sessionIsDirty(client, sessionId), false);
        const result = textPayload(
          await callTool(client, "formulon_image_operation", {
            sessionId,
            operation: "reorder",
            objectId,
            index,
          }),
        );
        assert.equal(await sessionIsDirty(client, sessionId), true);
        assert.equal(result.objectId, objectId);
        assert.equal(result.index, index);
        assert.deepEqual(result.order, await orderOf());
        return result.order;
      };

      assert.deepEqual(await reorder(c, 0), [c, a, b]);
      assert.deepEqual(await reorder(c, 2), [a, b, c]);
      assert.deepEqual(await reorder(a, 1), [b, a, c]);
      assert.deepEqual(await reorder(b, 2), [a, c, b]);
      assert.deepEqual(await reorder(b, 2), [a, c, b], "moving to the current slot is a no-op");
      assert.deepEqual(await reorder(a, 0), [a, c, b]);

      // Stacking order is stored in the file and survives a reopen; names stay with their ids.
      const saved = path.join(dir, "order.xlsx");
      textPayload(
        await callTool(client, "formulon_save_session", { sessionId, outputPath: saved }),
      );
      textPayload(
        await callTool(client, "formulon_open_workbook", {
          sessionId: "order-reopened",
          path: saved,
        }),
      );
      const reloaded = await listPictures(client, "order-reopened");
      assert.deepEqual(
        reloaded.map((item) => item.name),
        ["Pic B2", "Pic F6", "Pic D4"],
      );

      // Out-of-range positions and unknown ids fail without reordering anything.
      const before = await orderOf();
      for (const index of [3, 4, 1000]) {
        assert.match(
          errorPayload(
            await callTool(client, "formulon_image_operation", {
              sessionId,
              operation: "reorder",
              objectId: a,
              index,
            }),
          ),
          /reorder image failed/,
          `index ${index}`,
        );
      }
      assert.match(
        errorPayload(
          await callTool(client, "formulon_image_operation", {
            sessionId,
            operation: "reorder",
            objectId: 999,
            index: 0,
          }),
        ),
        /reorder image failed/,
      );
      assert.deepEqual(await orderOf(), before);

      // The schema refuses what cannot be a position or an id.
      for (const args of [{ index: -1 }, { index: 1.5 }, { objectId: 0, index: 0 }]) {
        assert.match(
          errorPayload(
            await callTool(client, "formulon_image_operation", {
              sessionId,
              operation: "reorder",
              objectId: a,
              ...args,
            }),
          ),
          /Input validation error/,
        );
      }
      assert.match(
        errorPayload(
          await callTool(client, "formulon_image_operation", {
            sessionId,
            operation: "reorder",
            objectId: a,
          }),
        ),
        /index is required/,
      );
    });
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("a move refuses a name or description instead of dropping it", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-image-rename-"));
  try {
    const logo = path.join(dir, "logo.png");
    await writeFile(logo, solidPng(40, 20));
    await withClient(async (client) => {
      const sessionId = "image-rename";
      textPayload(await callTool(client, "formulon_open_workbook", { sessionId }));
      const { objectId } = textPayload(
        await callTool(client, "formulon_image_operation", {
          sessionId,
          operation: "insert",
          path: logo,
          anchor: "B2",
          name: "Keep",
          description: "Kept alt text",
        }),
      );
      const before = await listPictures(client, sessionId);
      for (const extra of [
        { name: "New" },
        { description: "New alt" },
        { name: "", anchor: "D4" },
      ]) {
        assert.match(
          errorPayload(
            await callTool(client, "formulon_image_operation", {
              sessionId,
              operation: "move",
              objectId,
              ...extra,
            }),
          ),
          /name and description apply to insert only/,
        );
      }
      // Nothing moved: the refusal happens before the placement is touched.
      assert.deepEqual(await listPictures(client, sessionId), before);
    });
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
