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
	AscTest.Editor.GetDocument = AscBuilder.Word.Api.GetDocument.bind(AscTest.Editor);
	AscTest.JsApi = {};
	
	AscTest.JsApi.GetDocument = AscBuilder.Word.Api.GetDocument.bind(AscTest.Editor);
	AscTest.JsApi.ReplaceTextSmart = AscBuilder.Word.Api.ReplaceTextSmart.bind(AscTest.Editor);
	AscTest.JsApi.CreateRun = AscBuilder.Word.Api.CreateRun.bind(AscTest.Editor);
	AscTest.JsApi.CreateParagraph = AscBuilder.Word.Api.CreateParagraph.bind(AscTest.Editor);
	AscTest.JsApi.CreateInlineLvlSdt = AscBuilder.Word.Api.CreateInlineLvlSdt.bind(AscTest.Editor);
	AscTest.JsApi.CreateBlockLvlSdt  = AscBuilder.Word.Api.CreateBlockLvlSdt.bind(AscTest.Editor);
	AscTest.JsApi.CreateTable = AscBuilder.Word.Api.CreateTable.bind(AscTest.Editor);
	AscTest.JsApi.CreateShape = AscBuilder.Word.Api.CreateShape.bind(AscTest.Editor);
	AscTest.JsApi.CreateSolidFill = AscBuilder.Word.Api.CreateSolidFill.bind(AscTest.Editor);
	AscTest.JsApi.CreateStroke = AscBuilder.Word.Api.CreateStroke.bind(AscTest.Editor);
	AscTest.JsApi.CreateNoFill = AscBuilder.Word.Api.CreateNoFill.bind(AscTest.Editor);
	AscTest.JsApi.HexColor = AscBuilder.Word.Api.HexColor.bind(AscTest.Editor);
	AscTest.JsApi.ThemeColor = AscBuilder.Word.Api.ThemeColor.bind(AscTest.Editor);
	AscTest.JsApi.AutoColor = AscBuilder.Word.Api.AutoColor.bind(AscTest.Editor);
	AscTest.JsApi.RGBA = AscBuilder.Word.Api.RGBA.bind(AscTest.Editor);
	AscTest.JsApi.RGB = AscBuilder.Word.Api.RGB.bind(AscTest.Editor);
	AscTest.JsApi.FromJSON = AscBuilder.Word.Api.FromJSON.bind(AscTest.Editor);
	
	AscTest.JsApi.CreateDocContent = function()
	{
		let docContent = new AscWord.CDocumentContent();
		return new AscBuilder.ApiDocumentContent(docContent);
	};
	
	if (!QUnit.assert.equalShd)
	{
		QUnit.assert.equalBorder = function(actualColor, expectedColor, message)
		{
			if (expectedColor.IsThemeColor())
			{
				QUnit.assert.strictEqual(actualColor.IsThemeColor(), true, message + "check theme color type");
				QUnit.assert.strictEqual(actualColor.GetThemeName(), expectedColor.GetThemeName(), message + "check theme color");
			}
			else if (expectedColor.IsAutoColor())
			{
				QUnit.assert.strictEqual(actualColor.IsAutoColor(), true, message + "check auto color");
			}
			else
			{
				let actualRGB   = actualColor.GetRGB();
				let expectedRGB = expectedColor.GetRGB();
				QUnit.assert.strictEqual(actualRGB["r"], expectedRGB["r"], message + "check r component");
				QUnit.assert.strictEqual(actualRGB["g"], expectedRGB["g"], message + "check g component");
				QUnit.assert.strictEqual(actualRGB["b"], expectedRGB["b"], message + "check b component");
			}
		}
		
		QUnit.assert.equalShd = function(actualShd, expectedShd, message)
		{
			if (!actualShd || !expectedShd)
				return QUnit.assert.strictEqual(actualShd, expectedShd, message);
			
			message = message ? message + " - " : "";
			QUnit.assert.strictEqual(actualShd["Type"], expectedShd["Type"], message + "check type");
			
			if ("nil" === actualShd["Type"] && "nil" === expectedShd["Type"])
				return;
			
			QUnit.assert.equalBorder(actualShd["Color"], expectedShd["Color"], message);
		};
		
		QUnit.assert.equalBorder = function(actualBrd, expectedBrd, message)
		{
			if (!actualBrd || !expectedBrd)
				return QUnit.assert.strictEqual(actualBrd, expectedBrd, message);
			
			message = message ? message + " - " : "";
			
			QUnit.assert.strictEqual(actualBrd["Type"], expectedBrd["Type"], (message ? message : "") + " check type");
			
			if ("none" === actualBrd["Type"])
				return;
			
			QUnit.assert.strictEqual(actualBrd["Size"], expectedBrd["Size"], (message ? message : "") + " check type");
			QUnit.assert.strictEqual(actualBrd["Space"], expectedBrd["Space"], (message ? message : "") + " check type");
			
			QUnit.assert.equalBorder(actualBrd["Color"], expectedBrd["Color"], message);
		};
	}

	QUnit.testStart(function()
	{
		AscTest.CreateLogicDocument();
		AscTest.SetTrackRevisions(false);
		AscCommon.History.Clear();
		AscTest.ClearDocument();
	});
});
