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
	AscCommon.CFirstLetterExceptions.prototype.GetMaxLen = function()
	{
		return 7;
	}
	
	let logicDocument = AscTest.CreateLogicDocument();
	let p;
	
	QUnit.module("Test 'as you type' autocorrections");
	
	QUnit.testStart(function()
	{
		AscTest.ClearDocument();
	});
	
	function enterAndCheckText(assert, text, checkText)
	{
		p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);
		AscTest.MoveCursorToParagraph(p, true);
		AscTest.EnterText(text);
		AscTest.PressKey(AscTest.Key.space);
		assert.strictEqual(AscTest.GetParagraphText(p), checkText + " ", text + " -> " + checkText);
	}
	
	QUnit.test("Test: capitalize first letter of the sentence", function (assert)
	{
		logicDocument.SetAutoCorrectFirstLetterOfSentences(true);
		enterAndCheckText(assert, "hello", "Hello");
		enterAndCheckText(assert, "привет", "Привет");
		enterAndCheckText(assert, "გამარჯობა", "გამარჯობა"); // bug 69089
		enterAndCheckText(assert, "Hello world! hello", "Hello world! Hello");
		enterAndCheckText(assert, "Hello world! hello world", "Hello world! hello world");
	});
	
	QUnit.test("Test: capitalize first letter of the sentence in table and capitalize first letter of table cell", function (assert)
	{
		let table = AscTest.CreateTable(3, 3);
		logicDocument.AddToContent(0, table);
		let dc = table.GetRow(0).GetCell(0).GetContent();
		let p;
		
		function enterAndCheckTextToCell(assert, text, checkText)
		{
			dc.ClearContent(false);
			p = AscTest.CreateParagraph();
			dc.AddToContent(0, p);
			AscTest.MoveCursorToParagraph(p, true);
			AscTest.EnterText(text);
			AscTest.PressKey(AscTest.Key.space);
			assert.strictEqual(AscTest.GetParagraphText(p), checkText + " ", text + " -> " + checkText);
		}
		
		logicDocument.SetAutoCorrectFirstLetterOfCells(false);
		logicDocument.SetAutoCorrectFirstLetterOfSentences(false);
		enterAndCheckTextToCell(assert, "hello", "hello");
		enterAndCheckTextToCell(assert, "привет", "привет");
		enterAndCheckTextToCell(assert, "Hello world! hello", "Hello world! hello");
		enterAndCheckTextToCell(assert, "Hello world! hello world", "Hello world! hello world");
		
		logicDocument.SetAutoCorrectFirstLetterOfCells(true);
		logicDocument.SetAutoCorrectFirstLetterOfSentences(true);
		enterAndCheckTextToCell(assert, "hello", "Hello");
		enterAndCheckTextToCell(assert, "привет", "Привет");
		enterAndCheckTextToCell(assert, "Hello world! hello", "Hello world! Hello");
		enterAndCheckTextToCell(assert, "Hello world! hello world", "Hello world! hello world");
		
		logicDocument.SetAutoCorrectFirstLetterOfCells(false);
		logicDocument.SetAutoCorrectFirstLetterOfSentences(true);
		enterAndCheckTextToCell(assert, "hello", "hello");
		enterAndCheckTextToCell(assert, "привет", "привет");
		enterAndCheckTextToCell(assert, "Hello world! hello", "Hello world! Hello");
		enterAndCheckTextToCell(assert, "Hello world! hello world", "Hello world! hello world");
		
		logicDocument.SetAutoCorrectFirstLetterOfCells(true);
		logicDocument.SetAutoCorrectFirstLetterOfSentences(false);
		enterAndCheckTextToCell(assert, "hello", "Hello");
		enterAndCheckTextToCell(assert, "привет", "Привет");
		enterAndCheckTextToCell(assert, "Hello world! hello", "Hello world! hello");
		enterAndCheckTextToCell(assert, "Hello world! hello world", "Hello world! hello world");
	});
});
