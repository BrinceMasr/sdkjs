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
	
	let logicDocument    = AscTest.CreateLogicDocument();
	let styleManager     = logicDocument.GetStyles();
	let numberingManager = logicDocument.GetNumbering();
	
	function AddParagraph(style, text)
	{
		let p = AscTest.CreateParagraph();
		logicDocument.PushToContent(p);
		p.SetParagraphStyleById(style.GetId());
		
		let run = new AscWord.CRun();
		p.AddToContent(0, run);
		run.AddText(text);
		return p;
	}
	
	let styleCounter = 0;
	function CreateStyle()
	{
		let style = new AscWord.CStyle("style" + (++styleCounter), null, null, styletype_Paragraph);
		styleManager.Add(style);
		return style;
	}
	
	function CreateNum()
	{
		let num = numberingManager.CreateNum();
		numberingManager.AddNum(num);
		let numLvl = num.GetLvl(0).Copy();

		let paraPr = numLvl.GetParaPr();
		paraPr.Ind.Left      = 15;
		paraPr.Ind.FirstLine = 5;
		numLvl.SetParaPr(paraPr);
		
		num.SetLvl(numLvl, 0);
		
		return num;
	}
	
	QUnit.module("Style applicator");
	
	QUnit.test("Style application and change from current selection", function (assert)
	{
		AscTest.ClearDocument();
		
		let style = CreateStyle();
		p1 = AddParagraph(style, "First");
		p2 = AddParagraph(style, "Second");
		
		let num = CreateNum();
		
		function TestParagraph(p, left, right, first, align)
		{
			let compiledPr = p.GetCompiledParaPr(false);
			assert.strictEqual(compiledPr.Ind.Left, left, "Check Left indent");
			assert.strictEqual(compiledPr.Ind.Right, right, "Check Right indent");
			assert.strictEqual(compiledPr.Ind.FirstLine, first, "Check FirstLine");
			assert.strictEqual(compiledPr.Jc, align, "Check align");
		}
		
		assert.ok(true, "Create empty style and set it to two paragraphs. Check default values");
		
		TestParagraph(p1, 0, 0, 0, AscCommon.align_Left);
		TestParagraph(p2, 0, 0, 0, AscCommon.align_Left);
		
		
		assert.ok(true, "Apply direct properties and check values");
		p1.SetParagraphIndent({Left : 10, Right : 10, FirstLine : 20});
		p1.SetParagraphAlign(AscCommon.align_Center);
		
		TestParagraph(p1, 10, 10, 20, AscCommon.align_Center);

		assert.ok(true, "Apply direct numPr and check indents");
		p1.SetNumPr(num.GetId(), 0);
		p2.SetNumPr(num.GetId(), 0);
		
		TestParagraph(p1, 10, 10, 20, AscCommon.align_Center);
		TestParagraph(p2, 15, 0, 5, AscCommon.align_Left);
		
		// Updating a style in the UI currently works by getting the current formatting by selection, then applying it to the current style
		// and then applying this style to the selected part of the document
		let uiStyle = logicDocument.GetStyleFromFormatting();
		uiStyle.put_Name(style.GetName());
		AscTest.MoveCursorToParagraph(p1, true);
		logicDocument.Add_NewStyle(uiStyle);
		
		assert.ok(true, "Update style by paraPr of the first paragraph and check paragraph compiled properties after that");
		TestParagraph(p1, 10, 10, 20, AscCommon.align_Center);
		TestParagraph(p2, 10, 10, 20, AscCommon.align_Center);
	});
});
