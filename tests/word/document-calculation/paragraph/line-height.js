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

$(function () {
	
	const charWidth = AscTest.CharWidth * AscTest.FontSize;
	
	const L_FIELD = 20 * charWidth;
	const R_FIELD = 30 * charWidth;
	const PAGE_W  = 80 * charWidth;
	
	let logicDocument = AscTest.CreateLogicDocument();
	function initDocument()
	{
		AscTest.ClearDocument();
		logicDocument.AddToContent(0, AscTest.CreateParagraph());
		
		let sectPr = AscTest.GetFinalSection();
		sectPr.SetPageSize(PAGE_W, 1000);
		sectPr.SetPageMargins(L_FIELD, 50, R_FIELD, 50);
	}
	
	function checkLineHeight(assert, para, lines)
	{
		for (let i = 0, lineCount = Math.min(lines.length, para.GetLinesCount()); i < lineCount; ++i)
		{
			let _line = para.getLine(i);
			assert.close(_line.Bottom - _line.Top, lines[i], 0.001, "Check height for " + i + " line");
		}
	}
	
	QUnit.module("Test paragraph line height", {
		beforeEach : function()
		{
			initDocument();
			AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Current);
		}
	});
	
	// bug #74622
	QUnit.test("Test non-zero Y positioning for the whole line of text", function(assert)
	{
		let p1 = logicDocument.GetElement(0);
		
		p1.SetParagraphSpacing({After : 0, Before : 0, LineRule : linerule_Auto, Line : 1});
		p1.SetParagraphIndent({FirstLine : 0, Left : 0, Right : 0});
		
		AscTest.AddTextToParagraph(p1, "First line ________________ ");
		let run = AscTest.AddTextToParagraph(p1, "Second line");
		AscTest.Recalculate();
		
		assert.strictEqual(p1.getTextInLineRange(0, 0), "First line ________________ ", "Check the text in the first line");
		assert.strictEqual(p1.getTextInLineRange(1, 0), "Second line\r\n", "Check the text in the second line");
		
		checkLineHeight(assert, p1, [20, 20]);
		
		run.Set_Position(3);
		AscTest.Recalculate();
		checkLineHeight(assert, p1, [20, 20]);
		
		run.Set_Position(10);
		AscTest.Recalculate();
		checkLineHeight(assert, p1, [20, 25]);
		
		run.Set_Position(-10);
		AscTest.Recalculate();
		checkLineHeight(assert, p1, [20, 20]);
		
		run.Set_Position(-20);
		AscTest.Recalculate();
		checkLineHeight(assert, p1, [20, 25]);
	});
});
