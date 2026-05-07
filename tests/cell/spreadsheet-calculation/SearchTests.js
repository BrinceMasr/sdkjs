/*
 * (c) Copyright Ascensio System SIA 2010-2026
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

	// Environment stubs

	Asc.spreadsheet_api.prototype._init = function () {
		this._loadModules();
	};
	Asc.spreadsheet_api.prototype._loadFonts = function (fonts, callback) {
		callback();
	};
	Asc.spreadsheet_api.prototype.onEndLoadFile = function () {
		openDocument();
	};
	AscCommonExcel.WorkbookView.prototype._calcMaxDigitWidth = function () {};
	AscCommonExcel.WorkbookView.prototype._init = function () {};
	AscCommonExcel.WorkbookView.prototype._isLockedUserProtectedRange = function (callback) {
		callback(true);
	};
	AscCommonExcel.WorkbookView.prototype._onWSSelectionChanged = function () {};
	AscCommonExcel.WorkbookView.prototype.showWorksheet = function () {};
	AscCommonExcel.WorkbookView.prototype.recalculateDrawingObjects = function () {};
	AscCommonExcel.WorkbookView.prototype.restoreFocus = function () {};
	AscCommonExcel.WorksheetView.prototype._init = function () {};
	AscCommonExcel.WorksheetView.prototype.updateRanges = function () {};
	AscCommonExcel.WorksheetView.prototype._autoFitColumnsWidth = function () {};
	AscCommonExcel.WorksheetView.prototype.cleanSelection = function () {};
	AscCommonExcel.WorksheetView.prototype._drawSelection = function () {};
	AscCommonExcel.WorksheetView.prototype._scrollToRange = function () {};
	AscCommonExcel.WorksheetView.prototype.draw = function () {};
	AscCommonExcel.WorksheetView.prototype._prepareDrawingObjects = function () {};
	AscCommonExcel.WorksheetView.prototype._initCellsArea = function () {};
	AscCommonExcel.WorksheetView.prototype.getZoom = function () {};
	AscCommonExcel.WorksheetView.prototype._onUpdateFormatTable = function () {};
	AscCommonExcel.WorksheetView.prototype.setSelection = function () {};
	AscCommonExcel.WorksheetView.prototype._reinitializeScroll = function () {};
	AscCommonExcel.WorksheetView.prototype._getPPIX = function () {};
	AscCommonExcel.WorksheetView.prototype._getPPIY = function () {};
	AscCommonExcel.WorksheetView.prototype._prepareCellTextMetricsCache = function () {};
	AscCommonExcel.WorksheetView.prototype._isLockedCells = function (range, subType, callback) {
		callback(true);
		return true;
	};
	AscCommonExcel.WorksheetView.prototype._isLockedAll = function (callback) {
		callback(true);
	};
	AscCommonExcel.WorksheetView.prototype._isLockedFrozenPane = function (callback) {
		callback(true);
	};
	AscCommonExcel.WorksheetView.prototype._updateVisibleColsCount = function () {};
	AscCommonExcel.WorksheetView.prototype._calcActiveCellOffset = function () {};
	AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function () {};
	Asc.ReadDefTableStyles = function () {};

	// Document initialization

	function openDocument() {
		AscCommon.g_oTableId.init();
		api._onEndLoadSdk();
		api.isOpenOOXInBrowser = false;
		api.OpenDocumentFromBin(null, AscCommon.getEmpty());
	}

	var api = new Asc.spreadsheet_api({'id-view': 'editor_sdk'});
	api.FontLoader = {LoadDocumentFonts: function () {}};
	window['Asc']['editor'] = api;
	AscCommon.g_oTableId.init();
	api._onEndLoadSdk();
	api.isOpenOOXInBrowser = false;
	api.OpenDocumentFromBin(null, AscCommon.getEmpty());
	api.initCollaborativeEditing({});
	api.wb = new AscCommonExcel.WorkbookView(
		api.wbModel, api.controller, api.handlers, api.HtmlElement,
		api.topLineEditorElement, api, api.collaborativeEditing, api.fontRenderingMode
	);

	var wb = api.wbModel;
	wb.handlers.add('getSelectionState', function () { return null; });
	wb.handlers.add('getLockDefNameManagerStatus', function () { return true; });
	wb.handlers.add('asc_onConfirmAction', function (test1, callback) { callback(true); });

	var wsView = api.wb.getWorksheet(0);
	wsView.handlers = api.handlers;
	wsView.objectRender = new AscFormat.DrawingObjects();

	api.wb.model.handlers.add('changeDocument', function (prop, arg1, arg2) {
		api.wb.SearchEngine && api.wb.SearchEngine.changeDocument(prop, arg1, arg2);
	});

	var ws = api.wbModel.aWorksheets[0];

	// Helpers

	function makeFindOptions(findWhat) {
		var opts = new Asc.asc_CFindOptions();
		opts.findWhat = findWhat;
		opts.scanByRows = true;
		opts.scanForward = true;
		opts.isMatchCase = false;
		opts.isWholeCell = false;
		opts.lookIn = Asc.c_oAscFindLookIn.Value;
		opts.scanOnOnlySheet = Asc.c_oAscSearchBy.Sheet;
		opts.activeCell = {row: 0, col: 0};
		return opts;
	}

	function setCellValue(row, col, text) {
		ws._getCell(row, col, function (cell) {
			cell.setValue(text);
		});
	}

	function clearArea(r1, c1, r2, c2) {
		ws.getRange3(r1, c1, r2, c2)._foreachNoEmpty(function (cell) {
			cell.clearData();
			cell.saveContent(true);
		});
	}

	function resetSearch() {
		if (api.wb.SearchEngine) {
			api.wb.SearchEngine.Clear();
		}
		api.selectSearchingResults = false;
	}

	function getElements(engine) {
		return Object.keys(engine.Elements)
			.filter(function (id) { return engine.Elements[id]; })
			.map(function (id) { return engine.Elements[id]; })
			.sort(function (a, b) { return a.row !== b.row ? a.row - b.row : a.col - b.col; });
	}

	function setup() {
		resetSearch();
		clearArea(0, 0, 9, 9);
	}

	function setupWithSelectSearching() {
		resetSearch();
		clearArea(0, 0, 9, 9);
		api.asc_selectSearchingResults(true);
	}

	QUnit.test('Initial search populates engine with matching cells', function (assert) {
		setup();
		setCellValue(0, 0, 'hello');
		setCellValue(1, 0, 'world');

		var opts = makeFindOptions('hello');
		var engine = api.wb.Search(opts);

		assert.ok(engine, 'Search returns a SearchEngine');
		assert.strictEqual(engine.Count, 1, 'Exactly one cell matches "hello"');

		var elems = getElements(engine);
		assert.strictEqual(elems.length, 1, 'One element in Elements');
		assert.strictEqual(elems[0].row, 0, 'Found cell is at row 0 (A1)');
		assert.strictEqual(elems[0].col, 0, 'Found cell is at col 0 (column A)');
		assert.strictEqual(elems[0].text, 'hello', 'Found cell value is "hello"');
	});

	QUnit.test('Search returns cached engine when params are unchanged', function (assert) {
		setup();
		setCellValue(0, 0, 'hello');

		var opts = makeFindOptions('hello');
		var engine1 = api.wb.Search(opts);
		assert.strictEqual(engine1.Count, 1, 'First search: 1 match');

		var elems1 = getElements(engine1);
		assert.strictEqual(elems1[0].row, 0, 'Initial element is at row 0 (A1)');

		setCellValue(1, 0, 'hello');

		var opts2 = makeFindOptions('hello');
		var engine2 = api.wb.Search(opts2);

		assert.strictEqual(engine1, engine2, 'Same engine object is returned from cache');
		assert.strictEqual(engine2.Count, 1, 'Cache is used — Count unchanged despite new match at A2');

		var elems2 = getElements(engine2);
		assert.strictEqual(elems2.length, 1, 'Still only one element — cache not updated');
		assert.strictEqual(elems2[0].row, 0, 'Cached element is still at row 0 (A1)');
	});

	QUnit.test('Search re-runs when findWhat changes', function (assert) {
		setup();
		setCellValue(0, 0, 'hello');
		setCellValue(1, 0, 'world');

		var opts1 = makeFindOptions('hello');
		api.wb.Search(opts1);

		var opts2 = makeFindOptions('world');
		var engine = api.wb.Search(opts2);

		assert.strictEqual(engine.Count, 1, 'Re-search on different term finds "world"');

		var elems = getElements(engine);
		assert.strictEqual(elems.length, 1, 'One element in Elements');
		assert.strictEqual(elems[0].row, 1, 'Found "world" at row 1 (A2)');
		assert.strictEqual(elems[0].col, 0, 'Found "world" at col 0 (column A)');
		assert.strictEqual(elems[0].text, 'world', 'Found cell value is "world"');
	});

	QUnit.test('isNeedRecalc forces re-search even with unchanged params', function (assert) {
		setup();
		setCellValue(0, 0, 'hello');

		var opts = makeFindOptions('hello');
		var engine1 = api.wb.Search(opts);
		assert.strictEqual(engine1.Count, 1, 'First search: 1 match');

		setCellValue(1, 0, 'hello');

		var opts2 = makeFindOptions('hello');
		opts2.isNeedRecalc = true;
		var engine2 = api.wb.Search(opts2);

		assert.strictEqual(engine2.Count, 2, 'isNeedRecalc bypasses cache — finds both matches');

		var elems = getElements(engine2);
		assert.strictEqual(elems.length, 2, 'Two elements in Elements');
		assert.strictEqual(elems[0].row, 0, 'First element at row 0 (A1)');
		assert.strictEqual(elems[0].text, 'hello', 'A1 value is "hello"');
		assert.strictEqual(elems[1].row, 1, 'Second element at row 1 (A2)');
		assert.strictEqual(elems[1].text, 'hello', 'A2 value is "hello"');
	});

	QUnit.test('modifiedDocument=true with lastSearchElem set forces re-search (existing behaviour)', function (assert) {
		setup();
		setCellValue(0, 0, 'abc');
		setCellValue(1, 0, 'abc');

		var opts = makeFindOptions('abc');
		var engine = api.wb.Search(opts);
		assert.strictEqual(engine.Count, 2, 'Initial search: 2 matches');

		engine.modifiedDocument = true;
		engine.Count = 1;

		var opts2 = makeFindOptions('abc');
		opts2.lastSearchElem = [0, 'Sheet1', null, 'A1', 'abc', null];
		var engine2 = api.wb.Search(opts2);

		assert.strictEqual(engine2.Count, 2, 'Re-search after modification restores both results');

		var elems = getElements(engine2);
		assert.strictEqual(elems.length, 2, 'Two elements in Elements after re-search');
		assert.strictEqual(elems[0].row, 0, 'First element at row 0 (A1)');
		assert.strictEqual(elems[0].col, 0, 'First element at col 0');
		assert.strictEqual(elems[0].text, 'abc', 'A1 value is "abc"');
		assert.strictEqual(elems[1].row, 1, 'Second element at row 1 (A2)');
		assert.strictEqual(elems[1].col, 0, 'Second element at col 0');
		assert.strictEqual(elems[1].text, 'abc', 'A2 value is "abc"');
	});

	QUnit.test('modifiedDocument=true without lastSearchElem forces re-search (fix for bug 81461)', function (assert) {
		setup();
		setCellValue(0, 0, 'abc');

		var opts = makeFindOptions('abc');
		var engine = api.wb.Search(opts);
		assert.strictEqual(engine.Count, 1, 'Initial search: 1 match');

		engine.Count = 0;
		engine.Elements = {};
		engine.mapFindCells = {};
		engine.modifiedDocument = true;

		var opts2 = makeFindOptions('abc');
		var engine2 = api.wb.Search(opts2);

		assert.strictEqual(engine2.Count, 1, 'Re-search runs when modifiedDocument=true even without lastSearchElem');

		var elems = getElements(engine2);
		assert.strictEqual(elems.length, 1, 'One element in Elements after re-search');
		assert.strictEqual(elems[0].row, 0, 'Re-found element is at row 0 (A1)');
		assert.strictEqual(elems[0].col, 0, 'Re-found element is at col 0');
		assert.strictEqual(elems[0].text, 'abc', 'Re-found cell value is "abc"');
	});

	QUnit.test('modifiedDocument=false uses cached results', function (assert) {
		setup();
		setCellValue(0, 0, 'xyz');

		var opts = makeFindOptions('xyz');
		var engine = api.wb.Search(opts);
		assert.strictEqual(engine.Count, 1, 'Initial: 1 match');

		var elemsInitial = getElements(engine);
		assert.strictEqual(elemsInitial[0].row, 0, 'Initial element is at row 0 (A1)');

		setCellValue(1, 0, 'xyz');

		var opts2 = makeFindOptions('xyz');
		var engine2 = api.wb.Search(opts2);

		assert.strictEqual(engine2.Count, 1, 'Cache used when modifiedDocument is false');

		var elems = getElements(engine2);
		assert.strictEqual(elems.length, 1, 'Only the cached element is present');
		assert.strictEqual(elems[0].row, 0, 'Cached element is still at row 0 (A1), A2 not added');
	});

	QUnit.test('Search after delete + undo finds results (regression: bug 81461)', function (assert) {
		setupWithSelectSearching();
		setCellValue(0, 0, '1');
		setCellValue(1, 0, '2');

		var opts = makeFindOptions('1');
		var engine = api.wb.Search(opts);
		assert.strictEqual(engine.Count, 1, 'Initial search finds "1" in A1');
		assert.strictEqual(getElements(engine)[0].row, 0, 'Initial element at row 0 (A1)');

		ws._removeRows(0, 0);
		assert.strictEqual(engine.Count, 0, 'After row deletion: "1" is no longer in results');
		assert.strictEqual(getElements(engine).length, 0, 'No elements remain after deletion');

		AscCommon.History.Undo();

		assert.ok(api.wb.SearchEngine.modifiedDocument, 'modifiedDocument is true after undo restored the matching cell');
		assert.strictEqual(engine.Count, 0, 'Count is still 0 before re-search (cache not updated by Undo)');

		var opts2 = makeFindOptions('1');
		var engine2 = api.wb.Search(opts2);

		assert.strictEqual(engine2.Count, 1, 'After fix: re-search after delete+undo finds "1" again');

		var elems = getElements(engine2);
		assert.strictEqual(elems.length, 1, 'One element in Elements');
		assert.strictEqual(elems[0].row, 0, 'Re-found element is back at row 0 (A1)');
		assert.strictEqual(elems[0].col, 0, 'Re-found element is at col 0');
		assert.strictEqual(elems[0].text, '1', 'Re-found cell value is "1"');
	});

	QUnit.test('Search after delete + undo finds results when searching from non-zero active cell', function (assert) {
		setupWithSelectSearching();
		setCellValue(0, 0, '1');
		setCellValue(2, 0, '1');

		var opts = makeFindOptions('1');
		opts.activeCell = {row: 1, col: 0};
		api.wb.Search(opts);

		ws._removeRows(0, 0);
		AscCommon.History.Undo();

		var opts2 = makeFindOptions('1');
		opts2.activeCell = {row: 1, col: 0};
		var engine = api.wb.Search(opts2);

		assert.strictEqual(engine.Count, 2, 'Both matches restored after delete+undo');

		var elems = getElements(engine);
		assert.strictEqual(elems.length, 2, 'Two elements in Elements');
		assert.strictEqual(elems[0].row, 0, 'First element at row 0 (A1)');
		assert.strictEqual(elems[0].text, '1', 'A1 value is "1"');
		assert.strictEqual(elems[1].row, 2, 'Second element at row 2 (A3)');
		assert.strictEqual(elems[1].text, '1', 'A3 value is "1"');
	});

	QUnit.test('Element below deleted row has its row index decremented', function (assert) {
		setupWithSelectSearching();
		setCellValue(0, 0, 'x');
		setCellValue(1, 0, 'find-me');

		var opts = makeFindOptions('find-me');
		var engine = api.wb.Search(opts);
		assert.strictEqual(engine.Count, 1, 'Initial search: 1 match at A2 (row 1)');

		var elemsBefore = getElements(engine);
		assert.strictEqual(elemsBefore[0].row, 1, 'Element is at row 1 (A2) before deletion');
		assert.strictEqual(elemsBefore[0].col, 0, 'Element is at col 0');
		assert.strictEqual(elemsBefore[0].text, 'find-me', 'Element value is "find-me"');

		ws._removeRows(0, 0);

		var elemsAfter = getElements(engine);
		assert.strictEqual(elemsAfter[0].row, 0, 'After deleting row 0, element row index decremented to 0');
		assert.strictEqual(elemsAfter[0].col, 0, 'Col is unchanged');
		assert.strictEqual(elemsAfter[0].text, 'find-me', 'Text is unchanged');
		assert.strictEqual(engine.Count, 1, 'Count unchanged — element was below the deleted row');
	});

	QUnit.test('Element in deleted row is removed from engine (selectSearchingResults=true)', function (assert) {
		setupWithSelectSearching();
		setCellValue(0, 0, 'target');

		var opts = makeFindOptions('target');
		var engine = api.wb.Search(opts);
		assert.strictEqual(engine.Count, 1, 'Initial: element at row 0');

		var elemsBefore = getElements(engine);
		assert.strictEqual(elemsBefore[0].row, 0, 'Element is at row 0 (A1) before deletion');
		assert.strictEqual(elemsBefore[0].text, 'target', 'Element value is "target"');

		ws._removeRows(0, 0);

		assert.strictEqual(engine.Count, 0, 'Element removed from engine when its row is deleted');
		assert.strictEqual(getElements(engine).length, 0, 'Elements is empty after deletion');
	});

	QUnit.test('isMatchCase=false finds cell regardless of case', function (assert) {
		setup();
		setCellValue(0, 0, 'Hello');

		var opts = makeFindOptions('hello');
		var engine = api.wb.Search(opts);

		assert.strictEqual(engine.Count, 1, 'Case-insensitive: "hello" finds "Hello"');
		assert.strictEqual(getElements(engine)[0].text, 'Hello', 'Found cell value is "Hello"');
	});

	QUnit.test('isMatchCase=true does not find cell with different case', function (assert) {
		setup();
		setCellValue(0, 0, 'Hello');
		setCellValue(1, 0, 'hello');

		var opts = makeFindOptions('hello');
		opts.isMatchCase = true;
		var engine = api.wb.Search(opts);

		assert.strictEqual(engine.Count, 1, 'Case-sensitive: finds only exact "hello"');
		assert.strictEqual(getElements(engine)[0].row, 1, 'Found cell at row 1 (lowercase "hello")');
		assert.strictEqual(getElements(engine)[0].text, 'hello', 'Value is "hello"');
	});

	QUnit.test('isWholeCell=false finds partial match', function (assert) {
		setup();
		setCellValue(0, 0, 'hello world');

		var opts = makeFindOptions('hello');
		var engine = api.wb.Search(opts);

		assert.strictEqual(engine.Count, 1, 'Partial match finds "hello" inside "hello world"');
		assert.strictEqual(getElements(engine)[0].text, 'hello world', 'Found cell contains full text');
	});

	QUnit.test('isWholeCell=true does not find partial match, finds exact match', function (assert) {
		setup();
		setCellValue(0, 0, 'hello world');
		setCellValue(1, 0, 'hello');

		var opts1 = makeFindOptions('hello');
		opts1.isWholeCell = true;
		var engine1 = api.wb.Search(opts1);

		assert.strictEqual(engine1.Count, 1, 'Whole cell: "hello" not found in "hello world", but found in "hello"');
		assert.strictEqual(getElements(engine1)[0].row, 1, 'Found exact "hello" at row 1');

		var opts2 = makeFindOptions('hello world');
		opts2.isWholeCell = true;
		var engine2 = api.wb.Search(opts2);

		assert.strictEqual(engine2.Count, 1, 'Whole cell: exact "hello world" is found');
		assert.strictEqual(getElements(engine2)[0].row, 0, 'Found "hello world" at row 0');
	});

	QUnit.test('GetSearchElementId returns nearest element then advances on subsequent calls', function (assert) {
		setup();
		setCellValue(0, 0, 'nav');
		setCellValue(2, 0, 'nav');

		var opts = makeFindOptions('nav');
		var engine = api.wb.Search(opts);
		assert.strictEqual(engine.Count, 2, '2 matches');

		var id1 = api.wb.GetSearchElementId(true);
		assert.notStrictEqual(id1, null, 'First call returns an id');
		assert.strictEqual(engine.Elements[id1].row, 0, 'First result is nearest to activeCell — row 0 (A1)');

		engine.SetCurrent(id1);

		var id2 = api.wb.GetSearchElementId(true);
		assert.notStrictEqual(id2, null, 'Second call returns an id');
		assert.strictEqual(engine.Elements[id2].row, 2, 'Second result advances to row 2 (A3)');
	});

	QUnit.test('GetSearchElementId wraps around from last element to first', function (assert) {
		setup();
		setCellValue(0, 0, 'w');
		setCellValue(1, 0, 'w');
		setCellValue(2, 0, 'w');

		var opts = makeFindOptions('w');
		var engine = api.wb.Search(opts);
		assert.strictEqual(engine.Count, 3, '3 matches');

		var id1 = api.wb.GetSearchElementId(true);
		engine.SetCurrent(id1);
		var id2 = api.wb.GetSearchElementId(true);
		engine.SetCurrent(id2);
		var id3 = api.wb.GetSearchElementId(true);
		engine.SetCurrent(id3);
		var id4 = api.wb.GetSearchElementId(true);

		assert.strictEqual(engine.Elements[id1].row, 0, 'First at row 0');
		assert.strictEqual(engine.Elements[id2].row, 1, 'Second at row 1');
		assert.strictEqual(engine.Elements[id3].row, 2, 'Third at row 2');
		assert.strictEqual(engine.Elements[id4].row, 0, 'Wrapped back to row 0');
		assert.strictEqual(id4, id1, 'Wrap-around returns same id as first element');
	});

	QUnit.test('GetSearchElementId false navigates backward', function (assert) {
		setup();
		setCellValue(0, 0, 'bk');
		setCellValue(1, 0, 'bk');
		setCellValue(2, 0, 'bk');

		var opts = makeFindOptions('bk');
		var engine = api.wb.Search(opts);

		var id1 = api.wb.GetSearchElementId(true);
		engine.SetCurrent(id1);
		var id2 = api.wb.GetSearchElementId(true);
		engine.SetCurrent(id2);

		var idBack = api.wb.GetSearchElementId(false);

		assert.strictEqual(engine.Elements[id1].row, 0, 'Forward first: row 0');
		assert.strictEqual(engine.Elements[id2].row, 1, 'Forward second: row 1');
		assert.strictEqual(engine.Elements[idBack].row, 0, 'Backward from row 1 returns to row 0');
		assert.strictEqual(idBack, id1, 'Backward returns same id as first element');
	});

	QUnit.test('Multi-row delete removes elements in range and shifts elements below', function (assert) {
		setupWithSelectSearching();
		setCellValue(0, 0, 'target');
		setCellValue(1, 0, 'target');
		setCellValue(3, 0, 'target');

		var opts = makeFindOptions('target');
		var engine = api.wb.Search(opts);
		assert.strictEqual(engine.Count, 3, 'Initial: 3 matches at rows 0, 1, 3');

		ws._removeRows(0, 1);

		assert.strictEqual(engine.Count, 1, 'After deleting rows 0-1: 1 match remains');
		var elems = getElements(engine);
		assert.strictEqual(elems[0].row, 1, 'Element from row 3 shifted to row 1 (3 - 2 deleted rows)');
		assert.strictEqual(elems[0].text, 'target', 'Surviving element value is "target"');
	});

	QUnit.test('selectSearchingResults=false: element in deleted row survives in engine', function (assert) {
		setup();
		setCellValue(0, 0, 'survive');

		var opts = makeFindOptions('survive');
		var engine = api.wb.Search(opts);
		assert.strictEqual(engine.Count, 1, 'Initial: 1 match at row 0');

		ws._removeRows(0, 0);

		assert.strictEqual(engine.Count, 1, 'Element survives: changeCellValue is skipped when selectSearchingResults=false');
		assert.strictEqual(getElements(engine)[0].row, 0, 'Element row index unchanged (not adjusted by changeRemoveRows either)');
	});

	QUnit.test('Editing matched cell value removes it from engine', function (assert) {
		setupWithSelectSearching();
		setCellValue(0, 0, 'match');

		var opts = makeFindOptions('match');
		var engine = api.wb.Search(opts);
		assert.strictEqual(engine.Count, 1, 'Initial: 1 match');

		setCellValue(0, 0, 'other');

		assert.strictEqual(engine.Count, 0, 'After editing: element removed from engine');
		assert.strictEqual(getElements(engine).length, 0, 'Elements is empty');
		assert.strictEqual(engine.modifiedDocument, null, 'modifiedDocument not set — removal alone does not trigger it');
	});

	QUnit.test('New cell matching search term sets modifiedDocument', function (assert) {
		setupWithSelectSearching();

		var opts = makeFindOptions('appear');
		var engine = api.wb.Search(opts);
		assert.strictEqual(engine.Count, 0, 'Initial: no matches');
		assert.strictEqual(engine.modifiedDocument, null, 'modifiedDocument is null after empty search');

		setCellValue(0, 0, 'appear');

		assert.strictEqual(engine.modifiedDocument, true, 'modifiedDocument=true when a new matching cell appears');
		assert.strictEqual(engine.Count, 0, 'Count is still 0 — new cell is not added to Elements automatically');
	});

	QUnit.test('Search with no results returns empty engine', function (assert) {
		setup();
		setCellValue(0, 0, 'hello');

		var opts = makeFindOptions('xyz_not_present');
		var engine = api.wb.Search(opts);

		assert.ok(engine, 'Engine is returned even with no results');
		assert.strictEqual(engine.Count, 0, 'Count is 0');
		assert.strictEqual(getElements(engine).length, 0, 'Elements is empty');
	});

});
