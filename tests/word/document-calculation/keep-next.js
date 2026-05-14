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
	const logicDocument = AscTest.CreateLogicDocument()
	
	function setupDocument()
	{
		AscTest.ClearDocument();
		let sectPr = AscTest.GetFinalSection();
		sectPr.SetPageSize(400, 400);
		sectPr.SetPageMargins(50, 50, 50, 50);
	}
	
	
	QUnit.module("Test various situations keepNext property", {
		beforeEach : function()
		{
			setupDocument();
		}
	});
	
	QUnit.test("Test a simple situation with three paragraphs: keepNext = [false, true, false]", function (assert)
	{
		let p1 = AscTest.CreateParagraph();
		let p2 = AscTest.CreateParagraph();
		let p3 = AscTest.CreateParagraph();
		
		logicDocument.PushToContent(p1);
		logicDocument.PushToContent(p2);
		logicDocument.PushToContent(p3);
		
		AscTest.Recalculate();
		assert.strictEqual(logicDocument.GetPagesCount(), 1, "Should be 1 page");
		
		p1.SetParagraphSpacing({After: 250});

		AscTest.Recalculate();
		assert.strictEqual(p2.GetPagesCount(), 1, "p2 have 1 page");
		assert.strictEqual(p3.GetPagesCount(), 2, "p3 have 2 pages");
		assert.strictEqual(p3.IsEmptyPage(0) && !p3.IsEmptyPage(1), true, "p3 the first page is empty and the second is not");
		
		p2.SetParagraphKeepNext(true);
		AscTest.Recalculate();
		assert.strictEqual(p2.GetPagesCount(), 2, "p2 have 2 pages");
		assert.strictEqual(p2.IsEmptyPage(0) && !p2.IsEmptyPage(1), true, "p2 the first page is empty and the second is not");
		assert.strictEqual(p3.GetPagesCount(), 1, "p3 have 1 page");
		assert.strictEqual(p3.GetAbsoluteStartPage(), 1, "check p3 start page number");
	});
	
	QUnit.test("Test the case when a paragraph with the KeepNext property is followed by a table", function (assert)
	{
		let p1 = AscTest.CreateParagraph();
		let p2 = AscTest.CreateParagraph();
		
		let table = AscTest.CreateTable(3, 3);
		table.GetRow(0).SetHeight(50, Asc.linerule_AtLeast);
		table.GetRow(1).SetHeight(50, Asc.linerule_AtLeast);
		table.GetRow(2).SetHeight(50, Asc.linerule_AtLeast);
		
		logicDocument.PushToContent(p1);
		logicDocument.PushToContent(p2);
		logicDocument.PushToContent(table);
		
		AscTest.Recalculate();
		assert.strictEqual(logicDocument.GetPagesCount(), 1, "Should be 1 page");
		
		p1.SetParagraphSpacing({After: 250});
		
		AscTest.Recalculate();
		assert.strictEqual(p2.GetPagesCount(), 1, "p2 have 1 page");
		assert.strictEqual(table.GetPagesCount(), 2, "table have 2 pages");
		assert.strictEqual(table.IsEmptyPage(0) && !table.IsEmptyPage(1), true, "table's the first page is empty and the second is not");

		p2.SetParagraphKeepNext(true);
		AscTest.Recalculate();
		assert.strictEqual(p2.GetPagesCount(), 2, "p2 have 2 pages");
		assert.strictEqual(p2.IsEmptyPage(0) && !p2.IsEmptyPage(1), true, "p2 the first page is empty and the second is not");
		assert.strictEqual(table.GetPagesCount(), 1, "table have 1 page");
		assert.strictEqual(table.GetAbsoluteStartPage(), 1, "check table start page number");
	});
	
});
