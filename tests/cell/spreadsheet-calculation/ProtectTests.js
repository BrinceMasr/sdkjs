/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation, together with the
 * additional terms provided in the LICENSE file.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. For
 * details, see the GNU AGPL at: https://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA by email at info@onlyoffice.com
 * or by postal mail at 20A-6 Ernesta Birznieka-Upisha Street, Riga,
 * LV-1050, Latvia, European Union.
 *
 * The interactive user interfaces in modified versions of the Program
 * are required to display Appropriate Legal Notices in accordance with
 * Section 5 of the GNU AGPL version 3.
 *
 * No trademark rights are granted under this License.
 *
 * All non-code elements of the Product, including illustrations,
 * icon sets, and technical writing content, are licensed under the
 * Creative Commons Attribution-ShareAlike 4.0 International License:
 * https://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 * This license applies only to such non-code elements and does not
 * modify or replace the licensing terms applicable to the Program's
 * source code, which remains licensed under the GNU Affero General
 * Public License v3.
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */

QUnit.config.autostart = false;
$(function () {

	Asc.spreadsheet_api.prototype._init = function () {
		this._loadModules();
	};
	Asc.spreadsheet_api.prototype._loadFonts = function (fonts, callback) {
		callback();
	};
	Asc.spreadsheet_api.prototype.onEndLoadFile = function (fonts, callback) {
		openDocument();
	};
	AscCommonExcel.WorkbookView.prototype._calcMaxDigitWidth = function () {
	};
	AscCommonExcel.WorkbookView.prototype._init = function () {
	};
	AscCommonExcel.WorkbookView.prototype._isLockedUserProtectedRange = function (callback) {
		callback(true);
	};
	AscCommonExcel.WorkbookView.prototype._onWSSelectionChanged = function () {
	};
	AscCommonExcel.WorkbookView.prototype.showWorksheet = function () {
	};
	AscCommonExcel.WorkbookView.prototype.recalculateDrawingObjects = function () {
	};
	AscCommonExcel.WorkbookView.prototype.restoreFocus = function () {
	};
	AscCommonExcel.WorksheetView.prototype._init = function () {
	};
	AscCommonExcel.WorksheetView.prototype.updateRanges = function () {
	};
	AscCommonExcel.WorksheetView.prototype._autoFitColumnsWidth = function () {
	};
	AscCommonExcel.WorksheetView.prototype.cleanSelection = function () {
	};
	AscCommonExcel.WorksheetView.prototype._drawSelection = function () {
	};
	AscCommonExcel.WorksheetView.prototype._scrollToRange = function () {
	};
	AscCommonExcel.WorksheetView.prototype.draw = function () {
	};
	AscCommonExcel.WorksheetView.prototype._prepareDrawingObjects = function () {
	};
	AscCommonExcel.WorksheetView.prototype._initCellsArea = function () {
	};
	AscCommonExcel.WorksheetView.prototype.getZoom = function () {
	};
	AscCommonExcel.WorksheetView.prototype._prepareCellTextMetricsCache = function () {
	};

	AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function () {
	};
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
	AscCommonExcel.WorksheetView.prototype._updateVisibleColsCount = function () {
	};
	AscCommonExcel.WorksheetView.prototype._calcActiveCellOffset = function () {
	};

	var api = new Asc.spreadsheet_api({
		'id-view': 'editor_sdk'
	});
	api.FontLoader = {
		LoadDocumentFonts: function () {
			setTimeout(startTests, 0)
		}
	};
	window["Asc"]["editor"] = api;

	var wb, ws, wsview;

	function openDocument() {
		AscCommon.g_oTableId.init();
		api._onEndLoadSdk();
		api.isOpenOOXInBrowser = false;
		api.OpenDocumentFromBin(null, AscCommon.getEmpty());
		api.initCollaborativeEditing({});
		api.wb = new AscCommonExcel.WorkbookView(api.wbModel, api.controller, api.handlers, api.HtmlElement,
			api.topLineEditorElement, api, api.collaborativeEditing, api.fontRenderingMode);

		wb = api.wbModel;
		wb.handlers.add("getSelectionState", function () {
			return null;
		});

		wsview = api.wb.getWorksheet();
		wsview.objectRender = {};
		wsview.objectRender.updateDrawingObject = function () {
		};
		wsview.objectRender.updateSizeDrawingObjects = function () {
		};
		wsview.objectRender.selectedGraphicObjectsExists = function () {
		};
		wsview.handlers = {};
		wsview.handlers.trigger = function () {
		};
		ws = api.wbModel.aWorksheets[0];
	}

	function testValidTitle() {
		QUnit.test("Test: check valid title name", function (assert) {
			let checkRes = api.asc_checkProtectedRangeName("test");
			assert.strictEqual(checkRes, Asc.c_oAscDefinedNameReason.OK, "check valid name_1");
			checkRes = api.asc_checkProtectedRangeName("test1");
			assert.strictEqual(checkRes, Asc.c_oAscDefinedNameReason.OK, "check valid name_2");
			checkRes = api.asc_checkProtectedRangeName("test_1");
			assert.strictEqual(checkRes, Asc.c_oAscDefinedNameReason.OK, "check valid name_3");
			checkRes = api.asc_checkProtectedRangeName("test_ 1");
			assert.strictEqual(checkRes, Asc.c_oAscDefinedNameReason.OK, "check valid name_4");
			checkRes = api.asc_checkProtectedRangeName("test _ 1");
			assert.strictEqual(checkRes, Asc.c_oAscDefinedNameReason.OK, "check valid name_4");

			checkRes = api.asc_checkProtectedRangeName("test!");
			assert.strictEqual(checkRes, Asc.c_oAscDefinedNameReason.WrongName, "check valid name_5");
			checkRes = api.asc_checkProtectedRangeName("1test");
			assert.strictEqual(checkRes, Asc.c_oAscDefinedNameReason.WrongName, "check valid name_6");

			checkRes = api.asc_checkProtectedRangeName("t test");
			assert.strictEqual(checkRes, Asc.c_oAscDefinedNameReason.OK, "check valid name_8");
		});
	}

	QUnit.module("ProtectTests");

	function startTests() {
		QUnit.start();

		testValidTitle();
	}
});
