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

$(function ()
{
	function checkBounds(actual, expected, message)
	{
		let maxDifference = 0.01;
		
		QUnit.assert.close(actual.Left, expected.Left, maxDifference, "Left: " + message);
		QUnit.assert.close(actual.Top, expected.Top, maxDifference, "Top: " + message);
		QUnit.assert.close(actual.Right, expected.Right, maxDifference, "Right: " + message);
		QUnit.assert.close(actual.Bottom, expected.Bottom, maxDifference, "Bottom: " + message);
	}
	
	const logicDocument = AscTest.CreateLogicDocument()
	
	function setupDocument()
	{
		AscTest.ClearDocument();
		let sectPr = AscTest.GetFinalSection();
		sectPr.SetPageSize(400, 400);
		sectPr.SetPageMargins(50, 50, 50, 50);
	}
	
	
	QUnit.module("Test various situations with table header", {
		beforeEach : function()
		{
			setupDocument();
			AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Current);
		}
	});
	
	QUnit.test("Test page break", function (assert)
	{
		let paragraph = AscTest.CreateParagraph();
		logicDocument.PushToContent(paragraph);
		paragraph.SetParagraphSpacing({Before: 150});
		
		AscTest.Recalculate();
		assert.deepEqual(paragraph.GetPageBounds(0), new AscWord.CDocumentBounds(50, 50, 350, 200 + AscTest.FontHeight), "Check page bounds of the first paragraph");
		
		let table = AscTest.CreateTable(4, 3, [100, 100, 100]);
		logicDocument.PushToContent(table);

		AscTest.RemoveTableBorders(table);
		AscTest.RemoveTableMargins(table);

		table.GetRow(0).SetHeight(50, Asc.linerule_AtLeast);
		table.GetRow(1).SetHeight(50, Asc.linerule_AtLeast);
		table.GetRow(2).SetHeight(50, Asc.linerule_AtLeast);
		table.GetRow(3).SetHeight(50, Asc.linerule_AtLeast);
		
		let cellContent = table.GetRow(0).GetCell(0).GetContent();
		cellContent.PushToContent(AscTest.CreateParagraph());
		cellContent.PushToContent(AscTest.CreateParagraph());
		

		// Test a normal table divided into two pages
		AscTest.Recalculate();
		assert.strictEqual(table.GetPagesCount(), 2, "Test a normal table divided into two pages");
		checkBounds(table.getRowBounds(1, 0), new AscWord.CDocumentBounds(50, 280, 350, 330), "Check row bounds of the last row on the first page");
		checkBounds(table.getRowBounds(2, 1), new AscWord.CDocumentBounds(50, 50, 350, 100), "Check row bounds of the first row on the second page");
		
		// Test table with 1 heading row divided into two pages
		table.GetRow(0).SetHeader(true);
		AscTest.Recalculate();
		assert.strictEqual(table.GetPagesCount(), 2, "Test table with 1 heading row divided into two pages");
		checkBounds(table.getRowBounds(2, 1), new AscWord.CDocumentBounds(50, 110, 350, 160), "Check row bounds of the first row on the second page");
		
		AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Word12);
		
		// #62031 (next several cases)
		paragraph.SetParagraphSpacing({Before: 200});
		AscTest.Recalculate();
		assert.strictEqual(table.GetPagesCount(), 2, "The case when the table is split by the row following the header row (2010 compatibility)");
		assert.strictEqual(table.IsEmptyPage(0), false, "First page should be empty");
		assert.strictEqual(table.IsEmptyPage(1), false, "Check second page");
		
		paragraph.SetParagraphSpacing({Before: 225});
		AscTest.Recalculate();
		assert.strictEqual(table.GetPagesCount(), 2, "The case when the table is split by the row the header row (but row it's self can be split across multiple pages) (2010 compatibility)");
		assert.strictEqual(table.IsEmptyPage(0), true, "First page should be empty");
		assert.strictEqual(table.IsEmptyPage(1), false, "Check second page");
		
		AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Word15);
		// Case when the table splits on the row following the header row
		paragraph.SetParagraphSpacing({Before: 200});
		AscTest.Recalculate();
		assert.strictEqual(table.GetPagesCount(), 2, "The case when the table is split by the row following the header row (2015 compatibility)");
		assert.strictEqual(table.IsEmptyPage(0), true, "First page should be empty");
		assert.strictEqual(table.IsEmptyPage(1), false, "Check second page");
		
		// Case when the table splits on the header row
		paragraph.SetParagraphSpacing({Before: 225});
		AscTest.Recalculate();
		assert.strictEqual(table.GetPagesCount(), 2, "The case when the table is split by the row the header row (but row it's self can be split across multiple pages) (2015 compatibility)");
		assert.strictEqual(table.IsEmptyPage(0), true, "First page should be empty");
		assert.strictEqual(table.IsEmptyPage(1), false, "Check second page");
	});
	
});
