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
	
	let logicDocument = AscTest.CreateLogicDocument();
	let styleManager  = logicDocument.GetStyleManager();
	
	QUnit.module("Test adding and working with cross-references");
	
	QUnit.test("Test adding cross-ref to a block-level sdt", function (assert)
	{
		AscTest.ClearDocument();
		
		let cc = AscTest.CreateBlockLvlSdt();
		logicDocument.PushToContent(cc);
		
		let docContent = cc.GetContent();
		docContent.ClearContent(false);
		let headingParagraph = AscTest.CreateParagraph();
		docContent.AddToContent(0, headingParagraph);
		
		let styleId = styleManager.GetDefaultHeading(0);
		headingParagraph.SetParagraphStyle(styleManager.Get(styleId).GetName());
		AscTest.AddTextToParagraph(headingParagraph, "HeadingText");
		
		assert.strictEqual(AscTest.GetParagraphText(headingParagraph), "HeadingText", "Check paragraph text");
		
		let p = AscTest.CreateParagraph();
		logicDocument.PushToContent(p);
		AscTest.MoveCursorToParagraph(p);
		
		logicDocument.AddRefToParagraph(headingParagraph, 0, true, false, undefined);
		assert.strictEqual(AscTest.GetParagraphText(p), "HeadingText", "Check text after adding ref to a block-level sdt");
		assert.strictEqual(headingParagraph.GetBookmarkRefToParagraph(), "_Ref1", "Check bookmark name");
		
		// Check bug 69293. Adding a cross-ref to a locked block-level sdt
		
		AscTest.ClearDocument();
		
		cc = AscTest.CreateBlockLvlSdt();
		logicDocument.PushToContent(cc);
		
		docContent = cc.GetContent();
		docContent.ClearContent(false);
		headingParagraph = AscTest.CreateParagraph();
		docContent.AddToContent(0, headingParagraph);
		
		headingParagraph.SetParagraphStyle(styleManager.Get(styleId).GetName());
		AscTest.AddTextToParagraph(headingParagraph, "HeadingText");
		
		cc.SetContentControlLock(Asc.c_oAscSdtLockType.SdtContentLocked);
		
		p = AscTest.CreateParagraph();
		logicDocument.PushToContent(p);
		AscTest.MoveCursorToParagraph(p);
		
		logicDocument.AddRefToParagraph(headingParagraph, 0, true, false, undefined);
		assert.strictEqual(AscTest.GetParagraphText(p), "HeadingText", "Check text after adding ref to a locked block-level sdt");
		assert.strictEqual(headingParagraph.GetBookmarkRefToParagraph(), "_Ref1", "Check bookmark name");
	});
});
