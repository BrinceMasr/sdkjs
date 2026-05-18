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

	AscTest.CreateLogicDocument();

	const charWidth = AscTest.CharWidth * AscTest.FontSize;

	let dc = new AscWord.CDocumentContent();
	dc.ClearContent(false);

	let para = AscTest.CreateParagraph();
	dc.AddToContent(0, para);

	let run = new AscWord.CRun();
	para.AddToContent(0, run);

	function Recalculate(width)
	{
		if (!width)
			width = 100 * charWidth;

		dc.Reset(0, 0, width, 10000);
		dc.Recalculate_Page(0, true);
	}

	function SetText(text)
	{
		run.ClearContent();
		run.AddText(text);
	}

	function GetText()
	{
		return AscTest.GetParagraphText(para);
	}

	function TestText(assert, text)
	{
		SetText(text);
		Recalculate();

		assert.strictEqual(GetText(), text, "Paragraph text: " + text);
	}

	function TestCodePointType(assert, types)
	{
		let count = run.GetElementsCount();

		assert.strictEqual(count, types.length, "Check run element count");

		if (count !== types.length)
		{
			assert.true(false, "Bad elements and types length");
			return;
		}

		for (let index = 0; index < count; ++index)
		{
			let item = run.GetElement(index);
			assert.strictEqual(item.GetCodePointType(), types[index], "Check " + String.fromCodePoint(item.GetCodePoint()) + " code point type");
		}
	}

	function TestCursorMove(assert, count)
	{
		para.MoveCursorToStartPos();

		for (let index = 0; index < count; ++index)
		{
			assert.strictEqual(para.IsCursorAtEnd(), false, "Check cursor move right " + index);
			para.MoveCursorRight();
		}

		assert.strictEqual(para.IsCursorAtEnd(), true, "Check cursor at the end");

		para.MoveCursorToEndPos();
		assert.strictEqual(para.IsCursorAtEnd(), true, "Check cursor at the end");

		for (let index = 0; index < count; ++index)
		{
			assert.strictEqual(para.IsCursorAtBegin(), false, "Check cursor move left " + index);
			para.MoveCursorLeft();
		}

		assert.strictEqual(para.IsCursorAtBegin(), true, "Check cursor at the start");
	}

	function TestDelete(assert, text, countRemove, countDelete)
	{
		TestText(assert, text);
		para.MoveCursorToStartPos();

		for (let index = 0; index < countDelete; ++index)
		{
			assert.strictEqual(para.IsEmpty(), false, "Check delete " + index);
			para.Remove(1);
		}

		assert.strictEqual(para.IsEmpty(), true, "Check end of delete");

		TestText(assert, text);
		para.MoveCursorToEndPos();

		for (let index = 0; index < countRemove; ++index)
		{
			assert.strictEqual(para.IsEmpty(), false, "Check remove " + index);
			para.Remove(-1);
		}

		assert.strictEqual(para.IsEmpty(), true, "Check end of remove");

	}

	QUnit.module("Text shaper");

	QUnit.test("Test: \"code point types\"", function (assert)
	{
		function Test(text, codePointTypes, moveCount, removeCount, deleteCount)
		{
			TestText(assert, text);
			TestCodePointType(assert, codePointTypes);
			TestCursorMove(assert, moveCount);
			TestDelete(assert, text, removeCount, deleteCount);
		}

		Test("abc",
			[AscWord.CODEPOINT_TYPE.BASE, AscWord.CODEPOINT_TYPE.BASE, AscWord.CODEPOINT_TYPE.BASE],
			3,
			3,
			3);

		Test("ffi",
			[AscWord.CODEPOINT_TYPE.LIGATURE, AscWord.CODEPOINT_TYPE.LIGATURE_CONTINUE, AscWord.CODEPOINT_TYPE.LIGATURE_CONTINUE],
			3,
			3,
			3);

		Test("xyz",
			[AscWord.CODEPOINT_TYPE.BASE, AscWord.CODEPOINT_TYPE.COMBINING_MARK, AscWord.CODEPOINT_TYPE.COMBINING_MARK],
			1,
			3,
			1);

		// Check the diacritic mark that is not composed by the shaper
		Test("á",
			[AscWord.CODEPOINT_TYPE.BASE, AscWord.CODEPOINT_TYPE.BASE],
			1,
			2,
			1);
	});
});
