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

$(function () {
	QUnit.module('Test the ApiTableCell methods');
	
	QUnit.test('SetColor, GetColor', function(assert)
	{
		const table = AscTest.JsApi.CreateTable(2, 2);
		const cell  = table.GetCell(0, 0);
		
		assert.strictEqual(cell.GetBackgroundColor(), null, 'Color check for a newly created table cell');
		
		cell.SetBackgroundColor(255, 127, 0);
		assert.equalRgb(cell.GetBackgroundColor(), {
			r : 255,
			g : 127,
			b : 0
		}, 'Color check after setting color with RGB components');
		
		const hexColor = AscTest.JsApi.HexColor('bada55');
		cell.SetBackgroundColor(hexColor);
		assert.strictEqual(cell.GetBackgroundColor().GetHex(), '#BADA55', 'Color check after setting color with ApiColor (hex)');
		
		const themeColor = AscTest.JsApi.ThemeColor('accent2');
		cell.SetBackgroundColor(themeColor);
		assert.strictEqual(cell.GetBackgroundColor().IsThemeColor(), true, 'Color check after setting color with ApiColor (theme)');
		
		cell.SetBackgroundColor(0, 0, 0, true);
		assert.strictEqual(cell.GetBackgroundColor(), null, 'Color check after resetting color');
	});

	QUnit.test('AddText', function (assert)
	{
		const table = AscTest.JsApi.CreateTable(2, 2);
		const cell  = table.GetCell(0, 0);

		cell.AddText("Hello");
		assert.strictEqual(cell.GetText(), "Hello\t", "Check AddText appends text to cell");

		cell.AddText(" World");
		assert.strictEqual(cell.GetText(), "Hello World\t", "Check AddText appends to existing paragraph");
	});

	QUnit.test('GetText', function (assert)
	{
		const table = AscTest.JsApi.CreateTable(2, 2);
		const cell  = table.GetCell(0, 0);

		const run = cell.GetContent().GetElement(0).AddText("line1");
		run.AddLineBreak();
		run.AddText("line2");

		assert.strictEqual(cell.GetText(), "line1\rline2\t", "Check GetText returns cell text");
		assert.strictEqual(cell.GetText({"NewLineSeparator": "_nl_"}), "line1_nl_line2\t", "Check GetText with custom newline separator");
	});

	QUnit.test('SetText', function (assert)
	{
		const table = AscTest.JsApi.CreateTable(2, 2);
		const cell  = table.GetCell(0, 0);

		cell.AddText("Original");
		cell.SetText("Replaced");
		assert.strictEqual(cell.GetText(), "Replaced\t", "Check SetText replaces cell content");

		cell.SetText("");
		assert.strictEqual(cell.GetText(), "\t", "Check SetText with empty string");
	});
});
