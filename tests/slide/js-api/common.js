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

$(function()
{
    AscTest.Editor.GetPresentation = AscCommon.SlideEditorApi.prototype.GetPresentation.bind(AscTest.Editor);

    AscTest.Editor.private_checkPlaceholders = function(){};
    AscTest.Editor.private_CreateApiDocContent = AscCommon.SlideEditorApi.prototype.private_CreateApiDocContent.bind(AscTest.Editor);
    AscTest.Editor.private_CreateApiParagraph = AscCommon.SlideEditorApi.prototype.private_CreateApiParagraph.bind(AscTest.Editor);

	AscTest.JsApi = {};

	AscTest.JsApi.GetPresentation = AscCommon.SlideEditorApi.prototype.GetPresentation.bind(AscTest.Editor);
	AscTest.JsApi.CreateSlide = AscCommon.SlideEditorApi.prototype.CreateSlide.bind(AscTest.Editor);
	AscTest.JsApi.CreateMaster = AscCommon.SlideEditorApi.prototype.CreateMaster.bind(AscTest.Editor);
	AscTest.JsApi.CreateLayout = AscCommon.SlideEditorApi.prototype.CreateLayout.bind(AscTest.Editor);
	AscTest.JsApi.CreatePlaceholder = AscCommon.SlideEditorApi.prototype.CreatePlaceholder.bind(AscTest.Editor);
	AscTest.JsApi.CreateTheme = AscCommon.SlideEditorApi.prototype.CreateTheme.bind(AscTest.Editor);

	AscTest.JsApi.CreateImage = AscCommon.SlideEditorApi.prototype.CreateImage.bind(AscTest.Editor);
	AscTest.JsApi.CreateShape = AscCommon.SlideEditorApi.prototype.CreateShape.bind(AscTest.Editor);
	AscTest.JsApi.CreateChart = AscCommon.SlideEditorApi.prototype.CreateChart.bind(AscTest.Editor);
	AscTest.JsApi.CreateGroup = AscCommon.SlideEditorApi.prototype.CreateGroup.bind(AscTest.Editor);
	AscTest.JsApi.CreateTable = AscCommon.SlideEditorApi.prototype.CreateTable.bind(AscTest.Editor);
	AscTest.JsApi.CreateParagraph = AscCommon.SlideEditorApi.prototype.CreateParagraph.bind(AscTest.Editor);
	AscTest.JsApi.CreateWordArt = AscCommon.SlideEditorApi.prototype.CreateWordArt.bind(AscTest.Editor);

	AscTest.JsApi.CreateSolidFill = AscCommon.SlideEditorApi.prototype.CreateSolidFill.bind(AscTest.Editor);
	AscTest.JsApi.CreateLinearGradientFill = AscCommon.SlideEditorApi.prototype.CreateLinearGradientFill.bind(AscTest.Editor);
	AscTest.JsApi.CreateRadialGradientFill = AscCommon.SlideEditorApi.prototype.CreateRadialGradientFill.bind(AscTest.Editor);
	AscTest.JsApi.CreatePatternFill = AscCommon.SlideEditorApi.prototype.CreatePatternFill.bind(AscTest.Editor);
	AscTest.JsApi.CreateBlipFill = AscCommon.SlideEditorApi.prototype.CreateBlipFill.bind(AscTest.Editor);
	AscTest.JsApi.CreateNoFill = AscCommon.SlideEditorApi.prototype.CreateNoFill.bind(AscTest.Editor);
	AscTest.JsApi.CreateStroke = AscCommon.SlideEditorApi.prototype.CreateStroke.bind(AscTest.Editor);
	AscTest.JsApi.CreateGradientStop = AscCommon.SlideEditorApi.prototype.CreateGradientStop.bind(AscTest.Editor);

	AscTest.JsApi.CreateRGBColor = AscCommon.SlideEditorApi.prototype.CreateRGBColor.bind(AscTest.Editor);
	AscTest.JsApi.CreateSchemeColor = AscCommon.SlideEditorApi.prototype.CreateSchemeColor.bind(AscTest.Editor);
	AscTest.JsApi.CreatePresetColor = AscCommon.SlideEditorApi.prototype.CreatePresetColor.bind(AscTest.Editor);

	Asc.editor.getLogicDocument = function(){return Asc.editor.WordControl.m_oLogicDocument}
	// QUnit.testStart(function()
	// {
	// 	AscTest.CreateLogicDocument();
	// 	AscCommon.History.Clear();
	// 	AscTest.ClearDocument();
	// });
});
