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

"use strict";

(function (window)
{
	const charCellEditorWidth = 5 * 96 / 25.4; // Shaper returns character width of 5mm for font size 10
	const cellEditorWidth = 100 * 96 / 25.4;

	AscCommonExcel.WorkbookView.prototype._onUpdateCursor = function () {};
	AscCommonExcel.WorkbookView.prototype._canResize = function () {};
	AscCommonExcel.WorkbookView.prototype._onWSSelectionChanged = function () {};
	AscCommonExcel.WorkbookView.prototype.sendCursor = function () {};
	AscCommonExcel.WorkbookView.prototype.showWorksheet = function (index, bLockDraw)
	{
		index = AscFormat.isRealNumber(index) ? index : this.model.getActive();
		this.model.setActive(index);
		this.wsActive = index;
	};

	AscCommonExcel.WorksheetView.prototype._init = function () {this._initWorksheetDefaultWidth();};
	AscCommonExcel.WorksheetView.prototype.updateRanges = function () {};
	AscCommonExcel.WorksheetView.prototype._autoFitColumnsWidth = function () {};
	AscCommonExcel.WorksheetView.prototype.setSelection = function () {};
	AscCommonExcel.WorksheetView.prototype.cleanSelection = function () {};
	AscCommonExcel.WorksheetView.prototype.draw = function () {editor.wb._updateSelectionInfo()};
	AscCommonExcel.WorksheetView.prototype._onUpdateFormatTable = function () {};
	AscCommonExcel.WorksheetView.prototype._drawSelection = function () {};
	AscCommonExcel.WorksheetView.prototype._calcRangeOffset = function () {return {};};
	AscCommonExcel.WorksheetView.prototype._calcVisibleRows = function () {};
	AscCommonExcel.WorksheetView.prototype._calcVisibleColumns = function () {};
	AscCommonExcel.WorksheetView.prototype._normalizeViewRange = function () {};
	AscCommonExcel.WorksheetView.prototype._fixVisibleRange = function () {};
	AscCommonExcel.WorksheetView.prototype.getCursorTypeFromXY = function () {return {};};
	AscCommonExcel.WorksheetView.prototype._calcActiveCellOffset = function () {return {};};
	AscCommonExcel.WorksheetView.prototype._initCellsArea = function () {};
	AscCommonExcel.WorksheetView.prototype._prepareDrawingObjects = function ()
	{
		this.objectRender = new AscFormat.DrawingObjects();
		this.objectRender.init(this);
		this.objectRender.showDrawingObjects = function () {};
		this.objectRender.getContextWidth = function () {return 100};
		this.objectRender.getContextHeight = function () {return 100};
		this.objectRender.controller.updateOverlay = function () {};
		this.objectRender.controller.updateSelectionState = function () {};
	};

	AscCommonExcel.asc_CEventsController.prototype._createScrollBars = function () {};
	const OldEventControllerInit = AscCommonExcel.asc_CEventsController.prototype.init;
	AscCommonExcel.asc_CEventsController.prototype.init = function (view, widgetElem, canvasElem, handlers)
	{
		return OldEventControllerInit.call(this, view, document.createElement('div'), document.createElement('canvas'), handlers);
	};

	AscCommon.InitBrowserInputContext = function () {};

	Asc.DrawingContext.prototype.setFont = function () {};
	Asc.DrawingContext.prototype.setFont = function () {};
	Asc.DrawingContext.prototype.fillText = function () {};
	Asc.DrawingContext.prototype.getFontMetrics = function ()
	{
		return {ascender: 15, descender: 4, lineGap: 1, nat_scale: 1000, nat_y1: 1000, nat_y2: -1000};
	};
	Asc.DrawingContext.prototype.measureChar = function ()
	{
		return Asc.TextMetrics(charCellEditorWidth, 9, 10, 1, 1, 10, charCellEditorWidth);
	}

	AscFonts.CFontManager.prototype.MeasureChar = function ()
	{
		return {fAdvanceX: charCellEditorWidth, oBBox: {fMaxX: 0, fMinX: 0}};
	};

	AscCommonExcel.CellEditor.prototype._init = function ()
	{
		this.drawingCtx = new Asc.DrawingContext({
			canvas: document.createElement('canvas'), units: 0/*px*/, fmgrGraphics: this.fmgrGraphics, font: this.m_oFont
		});
		this.overlayCtx = new Asc.DrawingContext({
			canvas: document.createElement('canvas'), units: 0/*px*/, fmgrGraphics: this.fmgrGraphics, font: this.m_oFont
		});
		this.textRender = new AscCommonExcel.CellTextRender(this.drawingCtx);
	}
	AscCommonExcel.CellEditor.prototype.updateWizardMode = function () {};
	AscCommonExcel.CellEditor.prototype._cleanText = function () {};
	AscCommonExcel.CellEditor.prototype._cleanSelection = function () {};
	AscCommonExcel.CellEditor.prototype._adjustCanvas = function () {};
	AscCommonExcel.CellEditor.prototype._showCanvas = function () {};
	AscCommonExcel.CellEditor.prototype._calculateCanvasSize = function () {};
	AscCommonExcel.CellEditor.prototype._renderText = function () {};
	AscCommonExcel.CellEditor.prototype._updateCursorPosition = function () {};
	AscCommonExcel.CellEditor.prototype._updateCursor = function () {};
	AscCommonExcel.CellEditor.prototype._hideCanvas = function () {};
	const OldCellEditorOpen = AscCommonExcel.CellEditor.prototype.open;
	AscCommonExcel.CellEditor.prototype.open = function (options)
	{
		options.getSides = function ()
		{
			return {l: [0], r: [cellEditorWidth], b: [10], cellX: 0, cellY: 0, ri: 0, bi: 0};
		}
		OldCellEditorOpen.call(this, options);
	};

	Asc.spreadsheet_api.prototype._init = function () {this._loadModules();};
	Asc.spreadsheet_api.prototype._loadFonts = function (fonts, callback) {callback();};
	Asc.spreadsheet_api.prototype.onEndLoadFile = function (fonts, callback) {OpenDocument();};
	AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function () {this.ImageLoader = AscCommon.g_image_loader;};
	AscCommon.baseEditorsApi.prototype.getShortcut = function (e)
	{
		if (AscFormat.isRealNumber(e.nShortcutType))
		{
			return e.nShortcutType;
		}
	};

	let editor;
	function InitEditor(Callback)
	{
		editor = new Asc.spreadsheet_api({
			'id-view': 'editor_sdk'
		});
		editor.HtmlElement = document.createElement('div');
		editor.topLineEditorElement = document.createElement('input');
		editor.FontLoader = {
			LoadDocumentFonts: function ()
			{
				setTimeout(Callback, 0)
			}
		};
		let docInfo = new Asc.asc_CDocInfo();
		docInfo.asc_putTitle("TeSt.xlsx");
		editor.DocInfo = docInfo;

		window["Asc"]["editor"] = editor;
	}

	function OpenDocument()
	{
		AscCommon.g_oTableId.init();
		AscCommon.g_clipboardBase.Init(editor);
		editor._onEndLoadSdk();
		AscFormat.initStyleManager();
		editor.isOpenOOXInBrowser = false;
		editor.OpenDocumentFromBin(null, AscCommon.getEmpty());
		editor.initCollaborativeEditing({});
		editor.wb = new AscCommonExcel.WorkbookView(editor.wbModel, editor.controller, editor.handlers, editor.HtmlElement,
			editor.topLineEditorElement, editor, editor.collaborativeEditing, editor.fontRenderingMode);
		const wb = editor.wbModel;

		const wsView = editor.wb.getWorksheet(0);
		wsView.handlers = editor.handlers;
		wb.DrawingDocument.SelectEnabled = function () {};
		wb.DrawingDocument.UpdateTarget = function () {};
		wb.DrawingDocument.CheckTargetShow = function () {};
		wb.DrawingDocument.TargetStart = function () {};

		wb.setActive(0);
		editor.wb.wsActive = 0;
	}

	AscTestShortcut.InitEditor = InitEditor;
})(window);
