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
	const drawingDocument = {
		CanvasHit : null,
		CanvasHitContext : null,
		GoToPage: function (slideNumber)
		{
			AscTest.DrawingDocument.m_oLogicDocument.Set_CurPage(slideNumber);
			editor.WordControl.m_oDrawingDocument.SlideCurrent = slideNumber;
			this.Thumbnails && this.Thumbnails.SelectPage(slideNumber);
		},
		ConvertCoordsToCursorWR: function () {return {X:0,Y:0}},
		Notes_GetWidth: function () {return 100;},
		Notes_OnRecalculate: function () {return 100;},
		scrollToY: function () {},
		OnStartRecalculate: function () {},
		OnRecalculateSlide: function () {},
		OnEndRecalculate: function () {},
		UpdateTargetTransform : function(){},
		SetTargetColor : function(){},
		ClearCachePages : function(){},
		FirePaint : function(){},
		SetTargetSize : function(){},
		UpdateTarget : function(){},
		Set_RulerState_Paragraph : function(){},
		SelectEnabled : function(){},
		SelectShow : function(){},
		TargetStart : function(){},
		SetCursorType : function(){},
		TargetShow : function(){},
		TargetEnd : function(){},
		Update_MathTrack : function(){},
		CheckTableStyles : function(){},
		OnUpdateOverlay : function(){},
		AddPageSelection : function(){},
		IsTrackText  : function(){},
		Set_RulerState_Table  : function(){},
		SelectClear : function() {},
		GetDotsPerMM : function(value) {return 72;},
		GetMMPerDot : function(value){return value / this.GetDotsPerMM(1);},
		m_oNotesApi: {},
		clear: function () {},
		GetSlidesCount: function () {
			return editor.getCountSlides();
		},
		getGraphicController: function() {
			return AscTest.DrawingDocument.m_oLogicDocument.GetCurrentController();
		},
		getMaxScrolledY: function() {
			return 10;
		},
		getCurScrolledY : function() {
			return 0;
		}
	};

	drawingDocument.CanvasHit = document.createElement('canvas');
	drawingDocument.CanvasHitContext = drawingDocument.CanvasHit.getContext('2d');

	window['asc_docs_api'] = AscCommon.baseEditorsApi;

	const editor = new AscCommon.baseEditorsApi({});
	editor.WordControl = drawingDocument;
	editor.WordControl.m_oDrawingDocument = drawingDocument;
	editor.WordControl.m_oDrawingDocument.m_oWordControl = drawingDocument;
	editor.WordControl.m_oApi = editor;

	editor.textArtPreviewManager = drawingDocument;
	editor.externalChartCollector = {
		onUpdateExternalList: function () {},
		checkChart          : function () {}
	};

	editor.thumbnailsPosition = AscCommon.thumbnailsPositionMap.left;
	editor.asc_hideComments = function () {};
	editor.isSlideShow = function () {return false};
	editor.sync_HideComment = function () {};
	editor.sync_CanUndoCallback = function () {};
	editor.sync_CanRedoCallback = function () {};
	editor.CheckChangedDocument = function () {};
	editor.sync_BeginCatchSelectedElements = function(){};
	editor.ClearPropObjCallback = function(){};
	editor.sync_slidePropCallback = function(){};
	editor.sync_PrLineSpacingCallBack = function(){};
	editor.sync_EndCatchSelectedElements = function(){};
	editor.sync_CanAddHyperlinkCallback = function(){};
	editor.sync_shapePropCallback = function(){};
	editor.sync_VerticalTextAlign = function(){};
	editor.sync_Vert = function(){};
	editor.sync_animPropCallback = function(){};
	editor.sync_ImgPropCallback = function(){};
	editor.sync_TblPropCallback = function(){};
	editor.sync_EndCatchSelectedElements = function(){};
	editor.Update_ParaTab = function(){};
	editor.sync_ParaStyleName = function(){};
	editor.UpdateParagraphProp = function(){};
	editor.UpdateTextPr = function(){};
	editor.sync_MathPropCallback = function(){};
	editor.sync_HyperlinkPropCallback = function(){};
	editor.DemonstrationReporterEnd = function () {};
	editor.private_GetLogicDocument = function(){return this.WordControl.m_oLogicDocument;};
	editor.asc_getKeyboardLanguage = function(){return -1;};
	editor.isMasterMode = function(){return false;};
	editor.getCountSlides = function(){
		let oPresentation = this.private_GetLogicDocument();
		if(!oPresentation) return 0;
		return oPresentation.GetSlidesCount();
	};
	editor.initCollaborativeEditing = AscCommon.SlideEditorApi.prototype.initCollaborativeEditing.bind(editor);
	editor.getThumbnailsPosition  = AscCommon.SlideEditorApi.prototype.getThumbnailsPosition.bind(editor);
	editor.asc_PasteData = AscCommon.SlideEditorApi.prototype.asc_PasteData.bind(editor);
	editor.pre_Paste = function(_fonts, _images, callback)
	{
		callback(true);
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscTest.DrawingDocument = editor.WordControl.m_oDrawingDocument;
	AscTest.Editor = editor;

	window.editor = editor;
	Asc.editor = editor;
	editor.initCollaborativeEditing();

})(window);
