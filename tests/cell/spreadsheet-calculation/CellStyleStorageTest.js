/*
 * (c) Copyright Ascensio System SIA 2010-2024
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

$(function () {
	let CellStyleStorage = AscCommonExcel.CellStyleStorage;
	let CRangeAttrArray = AscCommonExcel.CRangeAttrArray;

	// Worksheet stub that exposes only what CellStyleStorage touches.
	function FakeWs() {
		this.cellStylesByCol = [];
	}
	CellStyleStorage.installOnWorksheet(FakeWs);

	function newWs() {
		return new FakeWs();
	}

	function bbox(c1, r1, c2, r2) {
		return { c1: c1, r1: r1, c2: c2, r2: r2 };
	}

	QUnit.test("Test: getCellStyleStore lazy creation and idempotency", function (assert) {
		let ws = newWs();
		assert.strictEqual(ws.getCellStyleStore(0), null, "no store by default");
		assert.strictEqual(ws.getCellStyleStore(0, false), null);
		assert.strictEqual(ws.cellStylesByCol[0], undefined);

		let s = ws.getCellStyleStore(0, true);
		assert.ok(s instanceof CRangeAttrArray, "creates a CRangeAttrArray");
		assert.strictEqual(ws.getCellStyleStore(0, true), s, "subsequent create returns the same instance");
		assert.strictEqual(ws.getCellStyleStore(0), s, "no-create path returns the existing instance");

		assert.strictEqual(ws.getCellStyleStore(-1, true), null, "negative col is rejected");
	});

	QUnit.test("Test: getCellXf default is 0 for empty/missing", function (assert) {
		let ws = newWs();
		assert.strictEqual(ws.getCellXf(0, 0), 0);
		assert.strictEqual(ws.getCellXf(100, 5), 0);
		assert.strictEqual(ws.getCellXf(0, -1), 0);
	});

	QUnit.test("Test: setCellXf with integer index", function (assert) {
		let ws = newWs();
		ws.setCellXf(10, 3, 7);
		assert.strictEqual(ws.getCellXf(10, 3), 7);
		assert.strictEqual(ws.getCellXf(9, 3), 0);
		assert.strictEqual(ws.getCellXf(10, 2), 0);
	});

	QUnit.test("Test: setCellXf treats 0 / null / undefined as no direct style", function (assert) {
		let ws = newWs();
		ws.setCellXf(5, 1, 42);
		assert.strictEqual(ws.getCellXf(5, 1), 42);

		ws.setCellXf(5, 1, 0);
		assert.strictEqual(ws.getCellXf(5, 1), 0, "0 clears");

		ws.setCellXf(5, 1, 99);
		ws.setCellXf(5, 1, null);
		assert.strictEqual(ws.getCellXf(5, 1), 0, "null clears");

		ws.setCellXf(5, 1, 99);
		ws.setCellXf(5, 1, undefined);
		assert.strictEqual(ws.getCellXf(5, 1), 0, "undefined clears");

		// negative index is treated as no style
		ws.setCellXf(5, 1, -3);
		assert.strictEqual(ws.getCellXf(5, 1), 0);
	});

	QUnit.test("Test: setCellXf accepts an object with getIndexNumber()", function (assert) {
		let ws = newWs();
		ws.setCellXf(0, 0, { getIndexNumber: function () { return 12; } });
		assert.strictEqual(ws.getCellXf(0, 0), 12);

		// getIndexNumber returning 0 -> clears.
		ws.setCellXf(0, 0, { getIndexNumber: function () { return 0; } });
		assert.strictEqual(ws.getCellXf(0, 0), 0);
	});

	QUnit.test("Test: setCellXfRange and clearCellXfRange", function (assert) {
		let ws = newWs();
		ws.setCellXfRange(2, 1, 4, 3, 5);
		// rows 2..4, cols 1..3 = 5; everything else 0.
		for (let r = 0; r <= 6; r++) {
			for (let c = 0; c <= 5; c++) {
				let v = ws.getCellXf(r, c);
				if (r >= 2 && r <= 4 && c >= 1 && c <= 3) {
					assert.strictEqual(v, 5, "r=" + r + " c=" + c);
				} else {
					assert.strictEqual(v, 0, "r=" + r + " c=" + c);
				}
			}
		}

		ws.clearCellXfRange(3, 2, 3, 2);
		assert.strictEqual(ws.getCellXf(3, 2), 0);
		assert.strictEqual(ws.getCellXf(3, 1), 5);

		// setCellXfRange with 0 clears across cols without creating empty stores.
		ws.setCellXfRange(2, 5, 4, 6, 0);
		assert.strictEqual(ws.getCellStyleStore(5), null, "no store created when clearing through index 0");
		assert.strictEqual(ws.getCellStyleStore(6), null);

		// inverted ranges no-op
		ws.setCellXfRange(5, 2, 3, 1, 9);
		ws.setCellXfRange(2, 3, 4, 1, 9);
		assert.strictEqual(ws.getCellXf(4, 1), 5, "inverted range ignored");
	});

	QUnit.test("Test: copyCellXfRange across stub worksheets", function (assert) {
		let from = newWs();
		from.setCellXfRange(0, 0, 1, 1, 7); // 2x2 block of 7 at (rows 0..1, cols 0..1)
		from.setCellXf(2, 2, 9);

		let to = newWs();
		to.setCellXfRange(10, 10, 12, 12, 3); // pre-existing pattern in the destination window

		to.copyCellXfRange(from, bbox(0, 0, 2, 2), bbox(10, 10, 12, 12));

		// dst rows 10..11, cols 10..11 should be 7 (from src 0..1, 0..1).
		assert.strictEqual(to.getCellXf(10, 10), 7);
		assert.strictEqual(to.getCellXf(11, 11), 7);
		// dst (12,12) should be 9 (from src (2,2)).
		assert.strictEqual(to.getCellXf(12, 12), 9);
		// dst cells inside the window with no source value: cleared.
		assert.strictEqual(to.getCellXf(12, 10), 0);
		assert.strictEqual(to.getCellXf(10, 12), 0);
		// outside the destination window: untouched (was 0 to begin with).
		assert.strictEqual(to.getCellXf(13, 13), 0);
	});

	QUnit.test("Test: copyCellXfRange within same worksheet (overlap-safe)", function (assert) {
		let ws = newWs();
		ws.setCellXfRange(0, 0, 2, 2, 4); // 3x3 of 4
		ws.copyCellXfRange(ws, bbox(0, 0, 2, 2), bbox(2, 2, 4, 4));
		// dst block 2..4 rows / 2..4 cols all == 4
		for (let r = 2; r <= 4; r++) {
			for (let c = 2; c <= 4; c++) {
				assert.strictEqual(ws.getCellXf(r, c), 4, "r=" + r + " c=" + c);
			}
		}
		// src rows 0..1 / cols 0..1 (the part that was NOT overwritten) still 4
		assert.strictEqual(ws.getCellXf(0, 0), 4);
		assert.strictEqual(ws.getCellXf(1, 1), 4);
	});

	QUnit.test("Test: shiftCellXfs 'down' inserts empty rows in cols", function (assert) {
		let ws = newWs();
		ws.setCellXfRange(0, 0, 2, 2, 1); // 3x3 of 1
		ws.shiftCellXfs(bbox(0, 5, 2, 7), null, "down"); // insert 3 rows starting at row 5 in cols 0..2
		// Existing data at rows 0..2 untouched.
		assert.strictEqual(ws.getCellXf(2, 2), 1);
		// Rows 5..7 are empty (newly inserted).
		assert.strictEqual(ws.getCellXf(5, 0), 0);
		assert.strictEqual(ws.getCellXf(7, 2), 0);
	});

	QUnit.test("Test: shiftCellXfs 'up' deletes rows", function (assert) {
		let ws = newWs();
		// rows 0..5, cols 0..2 = 1 (setCellXfRange takes r1, c1, r2, c2)
		ws.setCellXfRange(0, 0, 5, 2, 1);
		ws.shiftCellXfs(bbox(0, 2, 2, 3), null, "up"); // delete rows 2..3 in cols 0..2
		// After shift: rows 0..3 = 1 (rows 4..5 became 2..3); rows 4..5 cleared.
		assert.strictEqual(ws.getCellXf(0, 0), 1);
		assert.strictEqual(ws.getCellXf(3, 0), 1);
		assert.strictEqual(ws.getCellXf(4, 0), 0);
	});

	QUnit.test("Test: shiftCellXfs 'right' shifts columns and clears the gap", function (assert) {
		let ws = newWs();
		// Put a value in cols 1, 2, 3 at row 5
		ws.setCellXf(5, 1, 11);
		ws.setCellXf(5, 2, 22);
		ws.setCellXf(5, 3, 33);
		// Insert 2 columns at col 1 (so old col 1->3, old col 2->4, old col 3->5).
		ws.shiftCellXfs(bbox(1, 5, 2, 5), null, "right");
		assert.strictEqual(ws.getCellXf(5, 1), 0, "gap left at original c1");
		assert.strictEqual(ws.getCellXf(5, 2), 0, "gap left at c1+1");
		assert.strictEqual(ws.getCellXf(5, 3), 11);
		assert.strictEqual(ws.getCellXf(5, 4), 22);
		assert.strictEqual(ws.getCellXf(5, 5), 33);
	});

	QUnit.test("Test: shiftCellXfs 'left' pulls columns into the gap", function (assert) {
		let ws = newWs();
		// values at cols 1..5 row 7
		for (let c = 1; c <= 5; c++) {
			ws.setCellXf(7, c, c * 10);
		}
		// shift left over cols [2..3] at row 7
		ws.shiftCellXfs(bbox(2, 7, 3, 7), null, "left");
		// cols 1, 2, 3 should now hold 10 (unchanged), 40 (was col 4), 50 (was col 5)
		assert.strictEqual(ws.getCellXf(7, 1), 10);
		assert.strictEqual(ws.getCellXf(7, 2), 40);
		assert.strictEqual(ws.getCellXf(7, 3), 50);
		// Cols 4..5 cleared at row 7.
		assert.strictEqual(ws.getCellXf(7, 4), 0);
		assert.strictEqual(ws.getCellXf(7, 5), 0);
	});

	QUnit.test("Test: iterCellXfs emits runs in column-major order", function (assert) {
		let ws = newWs();
		ws.setCellXfRange(2, 1, 5, 1, 7); // col 1 rows 2..5
		ws.setCellXfRange(0, 2, 1, 2, 8); // col 2 rows 0..1
		ws.setCellXf(7, 4, 9);             // col 4 row 7

		let seen = [];
		ws.iterCellXfs(bbox(0, 0, 5, 10), function (lo, hi, col, v) {
			seen.push([lo, hi, col, v]);
		});
		assert.deepEqual(seen, [
			[2, 5, 1, 7],
			[0, 1, 2, 8],
			[7, 7, 4, 9]
		]);

		// Callback returning false aborts.
		let count = 0;
		ws.iterCellXfs(bbox(0, 0, 5, 10), function () {
			count++;
			return false;
		});
		assert.strictEqual(count, 1);

		// bbox clamps emitted (lo, hi).
		let clipped = [];
		ws.iterCellXfs(bbox(1, 3, 4, 6), function (lo, hi, col, v) {
			clipped.push([lo, hi, col, v]);
		});
		assert.deepEqual(clipped, [
			[3, 5, 1, 7]
		]);
	});

	QUnit.test("Test: stub Worksheet methods are pure additions, no side effects on cellsByCol", function (assert) {
		// Cross-check: the API never reads or writes anything outside cellStylesByCol on the stub.
		let ws = newWs();
		ws.cellsByColShouldBeUntouched = "sentinel";
		ws.setCellXfRange(0, 0, 99, 99, 1);
		ws.shiftCellXfs(bbox(0, 0, 5, 5), null, "down");
		ws.clearCellXfRange(0, 0, 99, 99);
		assert.strictEqual(ws.cellsByColShouldBeUntouched, "sentinel");
	});

	// Also verify that the methods are installed on the real Worksheet class
	// (so Stage 3 shadow code can use them through `this.setCellXf(...)` etc.).
	QUnit.test("Test: real Worksheet has the new style API", function (assert) {
		let Worksheet = AscCommonExcel.Worksheet;
		assert.ok(typeof Worksheet === "function", "AscCommonExcel.Worksheet exists");
		assert.ok(typeof Worksheet.prototype.getCellStyleStore === "function");
		assert.ok(typeof Worksheet.prototype.getCellXf === "function");
		assert.ok(typeof Worksheet.prototype.setCellXf === "function");
		assert.ok(typeof Worksheet.prototype.setCellXfRange === "function");
		assert.ok(typeof Worksheet.prototype.clearCellXfRange === "function");
		assert.ok(typeof Worksheet.prototype.copyCellXfRange === "function");
		assert.ok(typeof Worksheet.prototype.shiftCellXfs === "function");
		assert.ok(typeof Worksheet.prototype.iterCellXfs === "function");
	});
});
