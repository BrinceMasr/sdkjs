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

$(function ()
{
	QUnit.module('Test the ApiDocument methods');

	QUnit.test('SetFormsHighlight, GetFormsHighlight', function (assert)
	{
		const doc = AscTest.JsApi.GetDocument();

		assert.equalRgb(doc.GetFormsHighlight().GetRGB(), { r: 201, g: 200, b: 255 }, 'Check forms highlight returns default color');

		doc.SetFormsHighlight(255, 122, 100);
		assert.equalRgb(doc.GetFormsHighlight().GetRGB(), { r: 255, g: 122, b: 100 }, 'Check forms highlight after setting with RGB components');

		const hexColor = AscTest.JsApi.HexColor('a1b2c3');
		doc.SetFormsHighlight(hexColor);
		assert.equalRgb(doc.GetFormsHighlight().GetRGB(), { r: 161, g: 178, b: 195 }, 'Check forms highlight after setting with ApiColor');

		doc.SetFormsHighlight(0, 0, 0, true);
		assert.strictEqual(doc.GetFormsHighlight(), null, 'Check forms highlight is null after disabling it');
	});

	QUnit.test('SetControlsHighlight, GetControlsHighlight', function (assert)
	{
		const doc = AscTest.JsApi.GetDocument();

		assert.strictEqual(doc.GetControlsHighlight(), null, 'Check that controls highlight is null by default');

		doc.SetControlsHighlight(255, 122, 100);
		assert.equalRgb(doc.GetControlsHighlight().GetRGB(), { r: 255, g: 122, b: 100 }, 'Check controls highlight after setting with RGB components');

		const hexColor = AscTest.JsApi.HexColor('a1b2c3');
		doc.SetControlsHighlight(hexColor);
		assert.equalRgb(doc.GetControlsHighlight().GetRGB(), { r: 161, g: 178, b: 195 }, 'Check controls highlight after setting with ApiColor');

		doc.SetControlsHighlight(0, 0, 0, true);
		assert.strictEqual(doc.GetControlsHighlight(), null, 'Check that controls highlight is null after disabling it');
	});

	QUnit.test('InsertContent', function (assert)
	{
		const doc = AscTest.JsApi.GetDocument();
		doc.Push(AscTest.JsApi.CreateParagraph());

		const para = AscTest.JsApi.CreateParagraph();
		para.AddText("Hello");
		assert.strictEqual(doc.InsertContent([para]), true, 'InsertContent returns true for a single paragraph');
		assert.strictEqual(doc.GetElement(0).GetText(), "Hello\r\n", 'Inserted paragraph appears as first element');

		AscTest.ClearDocument();
		doc.Push(AscTest.JsApi.CreateParagraph());

		const para2 = AscTest.JsApi.CreateParagraph();
		para2.AddText("Block");
		const run = AscTest.JsApi.CreateRun();
		run.AddText("Inline");
		const countBefore = doc.GetElementsCount();
		assert.strictEqual(doc.InsertContent([para2, "plain text", run]), true, 'InsertContent returns true for mixed content');
		assert.ok(doc.GetElementsCount() >= countBefore + 2, 'Two new blocks added: one paragraph and one grouped inline paragraph');
		assert.strictEqual(doc.GetElement(0).GetText(), "Block\r\n", 'First inserted block is the ApiParagraph');
		assert.strictEqual(doc.GetElement(1).GetText(), "plain textInline\r\n", 'String and run are grouped into one paragraph');

		AscTest.ClearDocument();
		doc.Push(AscTest.JsApi.CreateParagraph());

		const table = AscTest.JsApi.CreateTable(2, 2);
		table.Cells[0][0].AddText("1");
		table.Cells[0][1].AddText("2");
		table.Cells[1][0].AddText("3");
		table.Cells[1][1].AddText("4");
		assert.strictEqual(doc.InsertContent([table]), true, 'InsertContent returns true when inserting a table');
		assert.strictEqual(doc.GetText({
			TableCellSeparator : "_c_",
			TableRowSeparator : "_r_",
			ParaSeparator : "_p_"
		}), "1_c_2_r_3_c_4_r__p_", 'Check document text');
	});
});
