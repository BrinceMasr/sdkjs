/*
 * (c) Copyright Ascensio System SIA 2010-2025
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

$(function()
{
    AscTest.Editor.GetPresentation = AscBuilder.Slide.Api.GetPresentation.bind(AscTest.Editor);

    AscTest.Editor.private_checkPlaceholders = function(){};
    AscTest.Editor.private_CreateApiDocContent = AscBuilder.Slide.Api.private_CreateApiDocContent.bind(AscTest.Editor);
    AscTest.Editor.private_CreateApiParagraph = AscBuilder.Slide.Api.private_CreateApiParagraph.bind(AscTest.Editor);

	AscTest.JsApi = {};

	AscTest.JsApi.GetPresentation = AscBuilder.Slide.Api.GetPresentation.bind(AscTest.Editor);
	AscTest.JsApi.CreateSlide = AscBuilder.Slide.Api.CreateSlide.bind(AscTest.Editor);
	AscTest.JsApi.CreateMaster = AscBuilder.Slide.Api.CreateMaster.bind(AscTest.Editor);
	AscTest.JsApi.CreateLayout = AscBuilder.Slide.Api.CreateLayout.bind(AscTest.Editor);
	AscTest.JsApi.CreatePlaceholder = AscBuilder.Slide.Api.CreatePlaceholder.bind(AscTest.Editor);
	AscTest.JsApi.CreateTheme = AscBuilder.Slide.Api.CreateTheme.bind(AscTest.Editor);

	AscTest.JsApi.CreateImage = AscBuilder.Slide.Api.CreateImage.bind(AscTest.Editor);
	AscTest.JsApi.CreateShape = AscBuilder.Slide.Api.CreateShape.bind(AscTest.Editor);
	AscTest.JsApi.CreateChart = AscBuilder.Slide.Api.CreateChart.bind(AscTest.Editor);
	AscTest.JsApi.CreateGroup = AscBuilder.Slide.Api.CreateGroup.bind(AscTest.Editor);
	AscTest.JsApi.CreateTable = AscBuilder.Slide.Api.CreateTable.bind(AscTest.Editor);
	AscTest.JsApi.CreateParagraph = AscBuilder.Slide.Api.CreateParagraph.bind(AscTest.Editor);
	AscTest.JsApi.CreateWordArt = AscBuilder.Slide.Api.CreateWordArt.bind(AscTest.Editor);

	AscTest.JsApi.CreateSolidFill = AscBuilder.Slide.Api.CreateSolidFill.bind(AscTest.Editor);
	AscTest.JsApi.CreateLinearGradientFill = AscBuilder.Slide.Api.CreateLinearGradientFill.bind(AscTest.Editor);
	AscTest.JsApi.CreateRadialGradientFill = AscBuilder.Slide.Api.CreateRadialGradientFill.bind(AscTest.Editor);
	AscTest.JsApi.CreatePatternFill = AscBuilder.Slide.Api.CreatePatternFill.bind(AscTest.Editor);
	AscTest.JsApi.CreateBlipFill = AscBuilder.Slide.Api.CreateBlipFill.bind(AscTest.Editor);
	AscTest.JsApi.CreateNoFill = AscBuilder.Slide.Api.CreateNoFill.bind(AscTest.Editor);
	AscTest.JsApi.CreateStroke = AscBuilder.Slide.Api.CreateStroke.bind(AscTest.Editor);
	AscTest.JsApi.CreateGradientStop = AscBuilder.Slide.Api.CreateGradientStop.bind(AscTest.Editor);

	AscTest.JsApi.CreateRGBColor = AscBuilder.Word.Api.RGB.bind(AscTest.Editor);
	AscTest.JsApi.CreateSchemeColor = AscBuilder.Slide.Api.CreateSchemeColor.bind(AscTest.Editor);
	AscTest.JsApi.CreatePresetColor = AscBuilder.Slide.Api.CreatePresetColor.bind(AscTest.Editor);

	Asc.editor.getLogicDocument = function(){return Asc.editor.WordControl.m_oLogicDocument}
	// QUnit.testStart(function()
	// {
	// 	AscTest.CreateLogicDocument();
	// 	AscCommon.History.Clear();
	// 	AscTest.ClearDocument();
	// });
});
