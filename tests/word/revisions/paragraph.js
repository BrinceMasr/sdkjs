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
	let logicDocument = AscTest.CreateLogicDocument();
	logicDocument.RemoveFromContent(0, logicDocument.GetElementsCount(), false);
	
	QUnit.module("Test the revisions in a paragraph");
	
	function fillDocument(runData)
	{
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.PushToContent(p);
		
		runData.forEach(function(d)
		{
			let run = AscTest.CreateRun();
			run.AddText(d.text);
			if (undefined !== d.reviewType)
				run.SetReviewType(d.reviewType);
			else
				run.SetReviewType(reviewtype_Common);
			
			p.AddToContentToEnd(run);
		});
		
		return p;
	}
	
	function fillDocument_1234test()
	{
		return fillDocument([
			{text : "1234", reviewType : reviewtype_Add},
			{text : "test"},
		]);
	}
	
	QUnit.test("Remove/replace text in a single run", function (assert)
	{
		AscTest.SetTrackRevisions(true);
		let p = fillDocument_1234test();
		assert.strictEqual(AscTest.GetParagraphText(p), "1234test", 'Check paragraph text');
		
		// TODO: We only checked appearance of the text, but didn't check review mode of the added text
		AscTest.SelectParagraphRange(p, 1, 3);
		AscTest.EnterText("QQQ");
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p),
			[
				[reviewtype_Add, "1QQQ4"],
				[reviewtype_Common, "test"],
			],
			"Select text. Enter text over selection"
		);
		
		p = fillDocument_1234test();
		AscTest.SelectParagraphRange(p, 1, 3);
		AscTest.PressKey(AscTest.Key.delete);
		AscTest.EnterText("QQQ");
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p),
			[
				[reviewtype_Add, "1QQQ4"],
				[reviewtype_Common, "test"],
			],
			"Select text. Press delete button. Enter text over selection"
		);
		
		p = fillDocument_1234test();
		AscTest.SelectParagraphRange(p, 1, 3);
		AscTest.PressKey(AscTest.Key.backspace);
		AscTest.EnterText("QQQ");
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p),
			[
				[reviewtype_Add, "1QQQ4"],
				[reviewtype_Common, "test"],
			],
			"Select text. Press backspace button. Enter text"
		);
	});
	
	QUnit.test("Remove/replace text in several runs", function (assert)
	{
		AscTest.SetTrackRevisions(true);
		let p = fillDocument_1234test();
		assert.strictEqual(AscTest.GetParagraphText(p), "1234test", 'Check paragraph text');

		// TODO: We only checked appearance of the text, but didn't check review mode of the added text
		AscTest.SelectParagraphRange(p, 1, 6);
		AscTest.EnterText("ABC");
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p),
			[
				[reviewtype_Add, "1ABC"],
				[reviewtype_Remove, "te"],
				[reviewtype_Common, "st"],
			],
			"Select text and enter text"
		);

		p = fillDocument_1234test();
		AscTest.SelectParagraphRange(p, 1, 6);
		AscTest.PressKey(AscTest.Key.delete);
		AscTest.EnterText("ABC");
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p),
			[
				[reviewtype_Add, "1"],
				[reviewtype_Remove, "te"],
				[reviewtype_Add, "ABC"],
				[reviewtype_Common, "st"],
			],
			"Select text. Press delete button. Enter text"
		);

		p = fillDocument_1234test();
		AscTest.SelectParagraphRange(p, 1, 6);
		AscTest.PressKey(AscTest.Key.backspace);
		AscTest.EnterText("ABC");
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p),
			[
				[reviewtype_Add, "1ABC"],
				[reviewtype_Remove, "te"],
				[reviewtype_Common, "st"],
			],
			"Select text. Press backspace button. Enter text"
		);

		p = fillDocument([
			{text : "Before"},
			{text : "1234", reviewType : reviewtype_Add},
			{text : "after"},
		]);
		AscTest.SelectParagraphRange(p, 8, 12);
		AscTest.EnterText("777");
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p),
			[
				[reviewtype_Common, "Before"],
				[reviewtype_Add, "12777"],
				[reviewtype_Remove, "af"],
				[reviewtype_Common, "ter"],
			],
			"Select text and enter text"
		);

		p = fillDocument([
			{text : "Before"},
			{text : "1234", reviewType : reviewtype_Add},
			{text : "after"},
		]);
		AscTest.SelectParagraphRange(p, 8, 12);
		AscTest.PressKey(AscTest.Key.delete);
		AscTest.EnterText("777");
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p),
			[
				[reviewtype_Common, "Before"],
				[reviewtype_Add, "12"],
				[reviewtype_Remove, "af"],
				[reviewtype_Add, "777"],
				[reviewtype_Common, "ter"],
			],
			"Select text. Press delete button. Enter text"
		);
		
		p = fillDocument([
			{text : "Before"},
			{text : "1234", reviewType : reviewtype_Add},
			{text : "after"},
		]);
		AscTest.SelectParagraphRange(p, 8, 12);
		AscTest.PressKey(AscTest.Key.backspace);
		AscTest.EnterText("777");
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p),
			[
				[reviewtype_Common, "Before"],
				[reviewtype_Add, "12777"],
				[reviewtype_Remove, "af"],
				[reviewtype_Common, "ter"],
			],
			"Select text. Press backspace button. Enter text"
		);
	});
});
