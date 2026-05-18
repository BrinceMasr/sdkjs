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
		let cc = new AscWord.CInlineLevelSdt();
		cc.SetPlaceholder(c_oAscDefaultPlaceholderName.Text);
		cc.ReplaceContentWithPlaceHolder();
		return cc;
	}
	
	function CreateCheckBoxContentControl()
	{
		let cc = new AscWord.CInlineLevelSdt();
		cc.ApplyCheckBoxPr(new AscWord.CSdtCheckBoxPr(), new AscWord.CTextPr());
		return cc;
	}
	
	function CreatePictureContentControl()
	{
		let cc = new AscWord.CInlineLevelSdt();
		cc.ApplyPicturePr(true);
		return cc;
	}
	
	function AddParagraph(text)
	{
		let p = AscTest.CreateParagraph();
		logicDocument.PushToContent(p);
		let run = new AscWord.CRun();
		p.AddToContent(0, run);
		if (text)
			run.AddText(text);
		
		return p;
	}
	
	QUnit.module("Test the positioning of the cursor and selection for inline-level content controls");
	
	QUnit.test("Test behaviour of controls filled with placeholder", function (assert)
	{
		function TestDeletionEmptyContentControl(isFromStart)
		{
			let key = isFromStart ? AscTest.Key.delete : AscTest.Key.backspace;
			let msg = isFromStart ? "delete" : "backspace";
			
			AscTest.ClearDocument();
			let p = AddParagraph("");
			let cc = CreateContentControl();
			assert.strictEqual(cc.IsUseInDocument(), false, "Create content control and check if it is being used in the document");
			
			p.AddToContent(0, new AscWord.CRun());
			p.AddToContent(1, cc);
			p.AddToContent(2, new AscWord.CRun());
			
			assert.strictEqual(AscTest.GetParagraphText(p), "Your text here", "Check text of the paragraph after adding inline content control");
			assert.strictEqual(cc.IsUseInDocument(), true, "Check if content control is being used in the document");
			
			AscTest.MoveCursorToParagraph(p, isFromStart);
			assert.strictEqual(cc.IsThisElementCurrent(), false, "Move cursor to the " + (isFromStart ? "start" : "end") + " of paragraph and check if cursor outside the content control");
			
			AscTest.PressKey(key);
			assert.ok(true, "Press " + msg + " and check content control state");
			assert.strictEqual(cc.IsPlaceHolder(), true, "Check if content control is filled with placeholder");
			assert.strictEqual(cc.IsSelectedOnlyThis(), true, "Check if content control is being selected");
			
			AscTest.PressKey(key);
			assert.ok(true, "Press " + msg + " second time and check content control state");
			assert.strictEqual(cc.IsPlaceHolder(), false, "Check if content control is not filled with placeholder");
			assert.strictEqual(cc.GetInnerText(), "", "Check if text of content control is empty");
			assert.strictEqual(cc.IsSelectedOnlyThis(), true, "Check if content control is being selected");
			
			AscTest.PressKey(key);
			assert.ok(true, "Press " + msg + " for the third time and check content control state");
			assert.strictEqual(cc.IsUseInDocument(), false, "Check if content control is not being used in the document");
			assert.strictEqual(AscTest.GetParagraphText(p), "", "Check text of the paragraph after removing content control");
		}
		
		// Test removeing a control filled with a placeholder using triple backspace/delete
		TestDeletionEmptyContentControl(true);
		TestDeletionEmptyContentControl(false);
	});
	
	QUnit.test("Test deletion checkbox content control", function (assert)
	{
		function TestCheckBoxDeletion(isFromStart)
		{
			let key = isFromStart ? AscTest.Key.delete : AscTest.Key.backspace;
			
			AscTest.ClearDocument();
			let p = AddParagraph("");
			let cc = CreateCheckBoxContentControl();
			assert.strictEqual(cc.IsUseInDocument(), false, "Create content control and check if it is being used in the document");
			
			p.AddToContent(0, new AscWord.CRun());
			p.AddToContent(1, cc);
			p.AddToContent(2, new AscWord.CRun());
			
			AscTest.MoveCursorToParagraph(p, isFromStart);
			AscTest.PressKey(key);
			
			assert.strictEqual(cc.IsSelectedOnlyThis(), true, "Check if content control is being selected");
			
			AscTest.PressKey(key);
			assert.strictEqual(cc.IsUseInDocument(), false, "Check if content control is not being used in the document");
			assert.strictEqual(AscTest.GetParagraphText(p), "", "Check text of the paragraph after removing content control");
		}
		
		// Test removing a checkbox using double backspace/delete
		TestCheckBoxDeletion(true);
		TestCheckBoxDeletion(false);
		
		function TestPictureDeletion(isFromStart)
		{
			let key = isFromStart ? AscTest.Key.delete : AscTest.Key.backspace;
			
			AscTest.ClearDocument();
			let p = AddParagraph("");
			let cc = CreatePictureContentControl();
			assert.strictEqual(cc.IsUseInDocument(), false, "Create content control and check if it is being used in the document");
			
			p.AddToContent(0, new AscWord.CRun());
			p.AddToContent(1, cc);
			p.AddToContent(2, new AscWord.CRun());
			
			AscTest.MoveCursorToParagraph(p, isFromStart);
			AscTest.PressKey(key);
			
			assert.strictEqual(cc.IsSelectedOnlyThis(), true, "Check if content control is being selected");
			
			AscTest.PressKey(key);
			assert.strictEqual(cc.IsUseInDocument(), false, "Check if content control is not being used in the document");
			assert.strictEqual(AscTest.GetParagraphText(p), "", "Check text of the paragraph after removing content control");
		}
		
		TestPictureDeletion(true);
		TestPictureDeletion(false);
	});

});
