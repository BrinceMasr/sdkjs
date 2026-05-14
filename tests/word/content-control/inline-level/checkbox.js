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
	
	QUnit.test("Test various actions with checkbox content control", function (assert)
	{
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);
		AscTest.MoveCursorToParagraph(p);
		
		let checkboxPr = new AscWord.CSdtCheckBoxPr();
		checkboxPr.SetCheckedSymbol("T".codePointAt(0));
		checkboxPr.SetUncheckedSymbol("F".codePointAt(0));
		let checkbox = logicDocument.AddContentControlCheckBox(checkboxPr);

		assert.strictEqual(checkbox.IsUseInDocument(), true, "Check if checkbox is added to the document");
		assert.strictEqual(checkbox.IsCheckBoxChecked(), false);
		assert.strictEqual(AscTest.GetParagraphText(p), "F");
		
		checkbox.ToggleCheckBox();
		assert.strictEqual(checkbox.IsCheckBoxChecked(), true);
		assert.strictEqual(AscTest.GetParagraphText(p), "T", "Check toggle checkbox in normal mode");
		
		AscTest.SetTrackRevisions(true);
		checkbox.ToggleCheckBox();
		assert.strictEqual(checkbox.IsCheckBoxChecked(), false);
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p),
			[
				[reviewtype_Add, "F"],
				[reviewtype_Remove, "T"],
			],
			"Check toggle checkbox in review"
		);
		
		checkbox.ToggleCheckBox();
		assert.strictEqual(checkbox.IsCheckBoxChecked(), true);
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p),
			[
				[reviewtype_Common, "T"],
			],
			"Check toggle checkbox in review"
		);
		AscTest.SetTrackRevisions(false);
		
		checkbox.ToggleCheckBox();
		assert.strictEqual(checkbox.IsCheckBoxChecked(), false);
		assert.strictEqual(AscTest.GetParagraphText(p), "F", "Check toggle checkbox in normal mode");
		
		AscTest.SetTrackRevisions(true);
		checkbox.ToggleCheckBox();
		assert.strictEqual(checkbox.IsCheckBoxChecked(), true);
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p),
			[
				[reviewtype_Add, "T"],
				[reviewtype_Remove, "F"],
			],
			"Check toggle checkbox in review"
		);
		
		checkbox.ToggleCheckBox();
		assert.strictEqual(checkbox.IsCheckBoxChecked(), false);
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p),
			[
				[reviewtype_Common, "F"],
			],
			"Check toggle checkbox in review"
		);
		
		
		let p2 = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p2);
		AscTest.MoveCursorToParagraph(p2);
		
		let checkbox2 = logicDocument.AddContentControlCheckBox(checkboxPr);
		assert.strictEqual(checkbox2.IsUseInDocument(), true, "Check if checkbox is added to the document");
		assert.strictEqual(checkbox2.IsCheckBoxChecked(), false);
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p2),
			[
				[reviewtype_Add, "F"],
			],
			"Check adding a checkbox in review"
		);
		
		checkbox2.ToggleCheckBox();
		assert.strictEqual(checkbox2.IsCheckBoxChecked(), true);
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p2),
			[
				[reviewtype_Add, "T"]
			],
			"Check toggle checkbox in review (check box was added in review)"
		);
		
		checkbox2.ToggleCheckBox();
		assert.strictEqual(checkbox2.IsCheckBoxChecked(), false);
		assert.deepEqual(
			AscTest.GetParagraphReviewText(p2),
			[
				[reviewtype_Add, "F"]
			],
			"Check toggle checkbox in review (check box was added in review)"
		);
		
		
		AscTest.SetTrackRevisions(false);
		
	});
	
});
