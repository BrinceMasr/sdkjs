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
	QUnit.module("ApiDocumentContent");
	
	
	QUnit.test("GetText", function (assert)
	{
		let docContent = AscTest.JsApi.CreateDocContent();
		let p = docContent.GetElement(0);
		let run = p.AddText("123");
		run.AddTabStop();
		run.AddText("456");
		run.AddLineBreak();
		run.AddText("789");
		
		let table = AscTest.JsApi.CreateTable(2,2);
		table.GetRow(0).GetCell(0).GetContent().GetElement(0).AddText("A");
		table.GetRow(0).GetCell(1).GetContent().GetElement(0).AddText("B");
		table.GetRow(1).GetCell(0).GetContent().GetElement(0).AddText("C");
		table.GetRow(1).GetCell(1).GetContent().GetElement(0).AddText("D");
		
		docContent.Push(table);
		
		assert.strictEqual(docContent.GetText(), "123\t456\r789\r\nA\tB\r\nC\tD\r\n\r\n", "Check GetText");
		assert.strictEqual(docContent.GetText({
			"TabSymbol" : "_t_",
			"NewLineSeparator" : "_nl_",
			"TableCellSeparator" : "_c_",
			"TableRowSeparator" : "_r_",
			"ParaSeparator" : "_p_"

		}), "123_t_456_nl_789_p_A_c_B_r_C_c_D_r__p_", "Check GetText");
	});

	QUnit.test("SetText", function (assert)
	{
		let docContent = AscTest.JsApi.CreateDocContent();
		docContent.GetElement(0).AddText("Original text");

		docContent.SetText("Replaced text");
		assert.strictEqual(docContent.GetText(), "Replaced text\r\n", "Check SetText replaces all content");

		docContent.SetText("");
		assert.strictEqual(docContent.GetText(), "\r\n", "Check SetText with empty string");
	});

	QUnit.test("AddText", function (assert)
	{
		let docContent = AscTest.JsApi.CreateDocContent();
		docContent.AddText("Appended");
		assert.strictEqual(docContent.GetText(), "Appended\r\n", "Check AddText appends text to last paragraph");

		docContent.AddText(" More");
		assert.strictEqual(docContent.GetText(), "Appended More\r\n", "Check AddText appends to existing last paragraph");
	});
});
