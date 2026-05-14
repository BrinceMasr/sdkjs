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
	
	function CreateContentControl()
	{
		let cc = new AscWord.CBlockLevelSdt();
		cc.SetPlaceholder(c_oAscDefaultPlaceholderName.Text);
		cc.ReplacePlaceHolderWithContent();
		cc.SetShowingPlcHdr(false);
		return cc;
	}

	function CreateParagraphWithText(text)
	{
		let p = AscTest.CreateParagraph();
		let run = new AscWord.CRun();
		p.AddToContent(0, run);
		run.AddText(text);
		return p;
	}
	
	QUnit.module("Test the positioning of the cursor and selection for inline-level content controls");
	
	QUnit.test("Test remove/delete after/before content control", function (assert)
	{
		AscTest.SetTrackRevisions(false);
		AscTest.ClearDocument();

		function CreateFilledContentControl(texts)
		{
			let cc = CreateContentControl();
			let docContent = cc.GetContent();
			docContent.ClearContent(false);
			
			for (let iText = 0, nTexts = texts.length; iText < nTexts; ++iText)
			{
				let p = CreateParagraphWithText(texts[iText]);
				docContent.AddToContent(iText, p);
			}
			
			return cc;
		}
		logicDocument.AddToContent(0, AscTest.CreateParagraph());
		
		let cc1 = CreateFilledContentControl(["Text1", "Text2"]);
		let cc2 = CreateFilledContentControl(["Text3", "Text4"]);
		let p = CreateParagraphWithText("123");
		let lastPara = cc1.GetLastParagraph();
		let firstPara = cc2.GetFirstParagraph();
		
		logicDocument.AddToContent(0, cc1);
		logicDocument.AddToContent(1, p);
		logicDocument.AddToContent(2, cc2);
		
		assert.strictEqual(logicDocument.GetElementsCount(), 4, "Check number of elements in logic document");

		AscTest.MoveCursorToParagraph(p, true);
		AscTest.PressKey(AscTest.Key.backspace);
		assert.ok(true, "Move to the start of the middle paragraph and click backspace button");
		assert.strictEqual(logicDocument.GetElementsCount(), 4, "Check number of elements in logic document");
		assert.strictEqual(p.IsUseInDocument(), true, "Check if paragraph is present in the document");
		assert.strictEqual(lastPara.IsThisElementCurrent() && lastPara.IsCursorAtEnd(), true, "Check cursor position at the end of the first content control");
		
		AscTest.MoveCursorToParagraph(p, false);
		AscTest.PressKey(AscTest.Key.delete);
		assert.ok(true, "Move to the end of the middle paragraph and click delete button");
		assert.strictEqual(logicDocument.GetElementsCount(), 4, "Check number of elements in logic document");
		assert.strictEqual(p.IsUseInDocument(), true, "Check if paragraph is present in the document");
		assert.strictEqual(firstPara.IsThisElementCurrent() && firstPara.IsCursorAtBegin(), true, "Check cursor position at the start of the second content control");
		
		AscTest.ClearParagraph(p);
		
		AscTest.MoveCursorToParagraph(p, true);
		AscTest.PressKey(AscTest.Key.backspace);
		assert.ok(true, "Move to the start of the middle paragraph and click backspace button");
		assert.strictEqual(logicDocument.GetElementsCount(), 3, "Check number of elements in logic document");
		assert.strictEqual(p.IsUseInDocument(), false, "Check if paragraph is present in the document");
		assert.strictEqual(lastPara.IsThisElementCurrent() && lastPara.IsCursorAtEnd(), true, "Check cursor position at the end of the first content control");
		
		logicDocument.AddToContent(1, p);
		AscTest.MoveCursorToParagraph(p, false);
		AscTest.PressKey(AscTest.Key.delete);
		assert.ok(true, "Move to the end of the middle paragraph and click delete button");
		assert.strictEqual(logicDocument.GetElementsCount(), 3, "Check number of elements in logic document");
		assert.strictEqual(p.IsUseInDocument(), false, "Check if paragraph is present in the document");
		assert.strictEqual(firstPara.IsThisElementCurrent() && firstPara.IsCursorAtBegin(), true, "Check cursor position at the start of the second content control");
		
		logicDocument.AddToContent(1, p);

		AscTest.SetTrackRevisions(true);
		AscTest.MoveCursorToParagraph(p, true);
		AscTest.PressKey(AscTest.Key.backspace);
		
		assert.strictEqual(logicDocument.IsTrackRevisions(), true, "Turn on track revisions");
		assert.ok(true, "Move to the start of the middle paragraph and click backspace button");
		assert.strictEqual(logicDocument.GetElementsCount(), 4, "Check number of elements in logic document");
		assert.strictEqual(p.IsUseInDocument(), true, "Check if paragraph is present in the document");
		assert.strictEqual(lastPara.IsThisElementCurrent() && lastPara.IsCursorAtEnd(), true, "Check cursor position at the end of the first content control");
		assert.strictEqual(lastPara.GetReviewType(), reviewtype_Remove, "Check that the last paragraph in first cc has become deleted on review");
		
		AscTest.MoveCursorToParagraph(p, false);
		AscTest.PressKey(AscTest.Key.delete);
		assert.ok(true, "Move to the end of the middle paragraph and click delete button");
		assert.strictEqual(logicDocument.GetElementsCount(), 4, "Check number of elements in logic document");
		assert.strictEqual(p.IsUseInDocument(), true, "Check if paragraph is present in the document");
		assert.strictEqual(firstPara.IsThisElementCurrent() && firstPara.IsCursorAtBegin(), true, "Check cursor position at the start of the second content control");
		assert.strictEqual(p.GetReviewType(), reviewtype_Remove, "Check that middle paragraph has become deleted on review");
		
		p.SetReviewType(reviewtype_Add);
		assert.strictEqual(p.GetReviewType(), reviewtype_Add, "Change review type of the middle paragraph to added on review");
		AscTest.MoveCursorToParagraph(p, false);
		AscTest.PressKey(AscTest.Key.delete);
		assert.ok(true, "Move to the end of the middle paragraph and click delete button");
		assert.strictEqual(logicDocument.GetElementsCount(), 3, "Check number of elements in logic document");
		assert.strictEqual(p.IsUseInDocument(), false, "Check if paragraph is present in the document");
		assert.strictEqual(firstPara.IsThisElementCurrent() && firstPara.IsCursorAtBegin(), true, "Check cursor position at the start of the second content control");
	});
	
});
