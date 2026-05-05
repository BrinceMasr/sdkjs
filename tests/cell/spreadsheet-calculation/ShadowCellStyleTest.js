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

	QUnit.module('ShadowCellStyle');

	let api = new Asc.spreadsheet_api({ 'id-view': 'editor_sdk' });
	api.FontLoader = { LoadDocumentFonts: function () {} };
	api.initCollaborativeEditing({});
	let wb = new AscCommonExcel.Workbook(new AscCommonExcel.asc_CHandlersList(), api, true);
	let ws = new AscCommonExcel.Worksheet(wb, 0);

	function makeFill(color) {
		// Build a CellXfs that g_StyleCache can dedupe via addXf.
		let fill = new AscCommonExcel.Fill();
		fill.fromColor(new AscCommonExcel.RgbColor(color));
		let xfs = new AscCommonExcel.CellXfs();
		xfs.fill = fill;
		return xfs;
	}

	QUnit.test('setStyleInternal mirrors xfIndex into cellStylesByCol', function (assert) {
		let cell = new AscCommonExcel.Cell(ws);
		cell.setRowCol(3, 7);
		cell.setStyleInternal(makeFill(0xFF0000));

		assert.ok(cell.xfs, 'cell.xfs assigned');
		assert.ok(cell.xfs.getIndexNumber() > 0, 'xfIndex is positive');
		assert.strictEqual(ws.getCellXf(3, 7), cell.xfs.getIndexNumber(),
			'cellStylesByCol holds the same xfIndex');

		// Re-setting null clears the mirror at that row.
		cell.setStyleInternal(null);
		assert.strictEqual(ws.getCellXf(3, 7), 0, 'null clears the mirror');
	});

	QUnit.test('setStyleInternal does NOT mirror when _isTransient = true', function (assert) {
		// Use a fresh column to avoid collisions with previous tests.
		let cell = new AscCommonExcel.Cell(ws);
		cell._isTransient = true;
		cell.setRowCol(0, 100);
		cell.setStyleInternal(makeFill(0x00FF00));

		assert.ok(cell.xfs, 'in-memory xfs is still assigned on the cell');
		assert.strictEqual(ws.getCellXf(0, 100), 0,
			'cellStylesByCol untouched for transient cell');
		assert.strictEqual(ws.getCellStyleStore(100), null,
			'no per-column store materialized for transient cell');
	});

	QUnit.test('setStyleInternal does NOT mirror when nRow/nCol unset', function (assert) {
		let cell = new AscCommonExcel.Cell(ws);
		// nRow / nCol stay -1 (unbound to a real position).
		cell.setStyleInternal(makeFill(0x0000FF));
		assert.ok(cell.xfs, 'in-memory xfs assigned');
		// No explicit storage column to check, just ensure no exception and
		// that getCellXf for col -1 stays at 0.
		assert.strictEqual(ws.getCellXf(0, -1), 0);
	});

	QUnit.test('clearDataKeepXf re-stamps the mirror after the direct xfs assignment', function (assert) {
		let cell = new AscCommonExcel.Cell(ws);
		cell.setRowCol(11, 11);
		cell.setStyleInternal(makeFill(0x123456));
		let idx = cell.xfs.getIndexNumber();
		assert.strictEqual(ws.getCellXf(11, 11), idx, 'mirror present before clearDataKeepXf');

		// Sneakily wipe the storage so we can verify clearDataKeepXf re-stamps it.
		ws.clearCellXfRange(11, 11, 11, 11);
		assert.strictEqual(ws.getCellXf(11, 11), 0, 'mirror manually cleared');

		// clearDataKeepXf preserves cell.xfs and must explicitly mirror the
		// preserved value back into cellStylesByCol.
		cell.clearDataKeepXf(null);
		assert.strictEqual(cell.xfs && cell.xfs.getIndexNumber(), idx,
			'cell.xfs preserved');
		assert.strictEqual(ws.getCellXf(11, 11), idx,
			'cellStylesByCol re-stamped after clearDataKeepXf');
	});

	QUnit.test('Range.setStyle mirrors via setStyleInternal for a regular range', function (assert) {
		let r1 = 50, c1 = 50, r2 = 52, c2 = 52;
		// Pick a fresh corner of the sheet so previous tests don't interfere.
		let range = ws.getRange3(r1, c1, r2, c2);
		range.setStyle(makeFill(0x778899));

		// Pick one cell in the range and verify the mirror landed there with
		// a positive xfIndex.
		let stamped = ws.getCellXf(r1, c1);
		assert.ok(stamped > 0, 'Range.setStyle mirrored an xfIndex into cellStylesByCol');
		assert.strictEqual(ws.getCellXf(r2, c2), stamped, 'opposite corner shares the same xfIndex');
		assert.strictEqual(ws.getCellXf(r1 + 1, c1 + 1), stamped, 'mid cell shares the same xfIndex');
		assert.strictEqual(ws.getCellXf(r1 - 1, c1), 0, 'cells outside the range untouched');
		assert.strictEqual(ws.getCellXf(r2 + 1, c2), 0, 'cells outside the range untouched');
	});

	QUnit.test('_initCell inheritance write does NOT pollute cellStylesByCol', function (assert) {
		// Stamp a row style on a fresh row, then trigger _initCell via _getCell
		// at a column with no direct cell style. The cell will inherit the row
		// style in-memory but cellStylesByCol must remain empty for that (row, col).
		let r = 200, c = 5;
		ws._getRow(r, function (row) {
			row.setStyle(makeFill(0xABCDEF));
		});
		// row.setStyle goes through Row.setStyle (not Cell.setStyleInternal)
		// so it should NOT touch cellStylesByCol.
		assert.strictEqual(ws.getCellXf(r, c), 0,
			'Row.setStyle alone does not populate cellStylesByCol');

		// Now bring in an empty cell at (r, c). _initCell will inherit row.xfs.
		ws._getCell(r, c, function (cell) {
			assert.ok(cell.xfs, 'cell inherited row xfs in-memory');
		});
		// But the storage at (r, c) should still be empty — only direct cell
		// styles belong in cellStylesByCol.
		assert.strictEqual(ws.getCellXf(r, c), 0,
			'_initCell inheritance is NOT mirrored into cellStylesByCol');
	});

	QUnit.test('mirrorCellStyle is a safe no-op for malformed input', function (assert) {
		let mirror = AscCommonExcel.CellStyleStorage.mirrorCellStyle;
		mirror(null);
		mirror(undefined);
		mirror({}); // no nRow/nCol/ws
		mirror({ nRow: -1, nCol: 0, ws: ws, _isTransient: false });
		mirror({ nRow: 0, nCol: -1, ws: ws, _isTransient: false });
		mirror({ nRow: 0, nCol: 0, ws: null, _isTransient: false });
		assert.ok(true, 'no exceptions thrown for malformed mirror inputs');
	});

	QUnit.test('Cell._isTransient defaults to false on every fresh Cell', function (assert) {
		let cell = new AscCommonExcel.Cell(ws);
		assert.strictEqual(cell._isTransient, false);
	});

	QUnit.test('Row._tempCell is marked transient', function (assert) {
		let row = new AscCommonExcel.Row(ws);
		assert.strictEqual(row._tempCell._isTransient, true,
			'Row._tempCell is reused across rows and must not pollute cellStylesByCol');
	});

	// ---- Stage 4: writer helpers ----

	QUnit.test('Stage 4: getWriterCellXfIndex prefers cellStylesByCol when populated', function (assert) {
		let cell = new AscCommonExcel.Cell(ws);
		cell.setRowCol(80, 80);
		cell.setStyleInternal(makeFill(0xCAFE00));
		let idx = cell.xfs.getIndexNumber();
		// Both stores are equivalent (shadow mode is in sync).
		assert.strictEqual(ws.getCellXf(80, 80), idx, 'shadow synced');

		let helper = AscCommonExcel.CellStyleStorage;
		assert.strictEqual(helper.getWriterCellXfIndex(ws, 80, 80, cell), idx,
			'helper returns the shadowed index');
		let xfs = helper.getWriterCellXfs(ws, 80, 80, cell);
		assert.ok(xfs && xfs.getIndexNumber() === idx,
			'helper resolves CellXfs from cellStylesByCol');
	});

	QUnit.test('Stage 4: getWriterCellXfs falls back to cell.xfs when cellStylesByCol is empty', function (assert) {
		// Simulate a "loaded from disk, never mutated" cell: in-memory xfs set
		// directly, no shadow mirror to cellStylesByCol.
		let cell = new AscCommonExcel.Cell(ws);
		cell.setRowCol(81, 81);
		// Pre-cache an xfs through g_StyleCache without going through setStyleInternal.
		let xfs = AscCommonExcel.g_StyleCache.addXf(makeFill(0xBADC0DE));
		cell.xfs = xfs;
		let idx = xfs.getIndexNumber();
		assert.strictEqual(ws.getCellXf(81, 81), 0,
			'cellStylesByCol intentionally empty (no setStyleInternal)');

		let helper = AscCommonExcel.CellStyleStorage;
		assert.strictEqual(helper.getWriterCellXfIndex(ws, 81, 81, cell), idx,
			'helper falls back to cell.xfs when shadow has no entry');
		let resolved = helper.getWriterCellXfs(ws, 81, 81, cell);
		assert.strictEqual(resolved && resolved.getIndexNumber(), idx,
			'helper returns the fallback CellXfs');
	});

	QUnit.test('Stage 4: getWriterCellXfIndex returns 0 when neither store has a style', function (assert) {
		let cell = new AscCommonExcel.Cell(ws);
		cell.setRowCol(82, 82);
		// Both stores empty.
		let helper = AscCommonExcel.CellStyleStorage;
		assert.strictEqual(helper.getWriterCellXfIndex(ws, 82, 82, cell), 0);
		assert.strictEqual(helper.getWriterCellXfs(ws, 82, 82, cell), null);
	});

	QUnit.test('Stage 4: getWriterCellXfs is robust to missing ws / negative coords / null cell', function (assert) {
		let helper = AscCommonExcel.CellStyleStorage;
		assert.strictEqual(helper.getWriterCellXfs(null, 0, 0, null), null);
		assert.strictEqual(helper.getWriterCellXfs(ws, -1, 0, null), null);
		assert.strictEqual(helper.getWriterCellXfs(ws, 0, -1, null), null);
		// With a cell having xfs but invalid coords, the helper still falls
		// back to cell.xfs (writer's last-resort safety).
		let cell = new AscCommonExcel.Cell(ws);
		cell.xfs = AscCommonExcel.g_StyleCache.addXf(makeFill(0x0F0F0F));
		let resolved = helper.getWriterCellXfs(ws, -1, -1, cell);
		assert.ok(resolved, 'fallback engages even when coords are invalid');
		assert.strictEqual(resolved.getIndexNumber(), cell.xfs.getIndexNumber());
	});

	QUnit.test('Stage 4: when cellStylesByCol disagrees with cell.xfs, cellStylesByCol wins', function (assert) {
		let cell = new AscCommonExcel.Cell(ws);
		cell.setRowCol(83, 83);
		cell.setStyleInternal(makeFill(0x111111));
		let firstIdx = cell.xfs.getIndexNumber();

		// Manually override the shadow store with a different value to
		// simulate a divergence; helper must follow the new primary source.
		let secondXfs = AscCommonExcel.g_StyleCache.addXf(makeFill(0x222222));
		ws.setCellXf(83, 83, secondXfs);
		assert.notStrictEqual(secondXfs.getIndexNumber(), firstIdx,
			'two distinct xfIndex values for the test');

		let helper = AscCommonExcel.CellStyleStorage;
		assert.strictEqual(helper.getWriterCellXfIndex(ws, 83, 83, cell),
			secondXfs.getIndexNumber(),
			'cellStylesByCol is the new primary read source');
	});
});
