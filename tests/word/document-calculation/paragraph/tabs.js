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
	// For correct calculation of a tab position we need to create logic document
	// since position depends on overall document margins
	let logicDocument = AscTest.CreateLogicDocument();
	let charWidth     = AscTest.CharWidth * AscTest.FontSize;
	
	let sectPr = AscTest.GetFinalSection();
	sectPr.SetPageSize(100 * charWidth, 1000);
	sectPr.SetPageMargins(10 * charWidth, 50, 15 * charWidth, 50);
	
	let p = AscTest.CreateParagraph();
	logicDocument.AddToContent(0, p);
	
	let r = AscTest.CreateRun();
	p.AddToContent(0, r);
	
	function setTabs(tabs)
	{
		let paraTabs = new CParaTabs();
		tabs.forEach(t => paraTabs.Add(new CParaTab(t.value, t.pos, t.leader)));
		p.SetParagraphTabs(paraTabs);
	}
	
	QUnit.module("Paragraph tabs calculation");
	
	QUnit.test("Special case for left tab which exceed right edge", function (assert)
	{
		// Check situation when left tab lies between right edge of a paragraph and right field of the document
		r.AddText("Before\tafter");
		p.SetParagraphIndent({Right : 20 * charWidth});
		setTabs([{value : tab_Left, pos : 70 * charWidth}]);
		
		AscTest.Recalculate();
		assert.strictEqual(p.GetLinesCount(), 1, "Check number of lines");
		assert.strictEqual(p.GetTextOnLine(0), "Before after", "Text on line 0 'Before after'");
	});
});
