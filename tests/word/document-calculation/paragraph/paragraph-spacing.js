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

	let dc = new AscWord.CDocumentContent();
	dc.ClearContent(false);

	let p1 = new AscWord.Paragraph();
	let p2 = new AscWord.Paragraph();

	dc.AddToContent(0, p1);
	dc.AddToContent(1, p2);

	let r1 = new AscWord.CRun();
	p1.AddToContent(0, r1);
	r1.AddText("Hello Word!");

	let r2 = new AscWord.CRun();
	p2.AddToContent(0, r2);
	r2.AddText("Абракадабра");

	const pageWidth = 20 * AscTest.CharWidth * AscTest.FontSize;
	function Recalculate()
	{
		dc.Reset(0, 0, 20 * AscTest.CharWidth * AscTest.FontSize, 10000);
		dc.Recalculate_Page(0, true);
	}

	QUnit.module("Paragraph Spacing");


	QUnit.test("Test: \"Paragraphs\"", function (assert)
	{
		p1.SetParagraphSpacing({Before : 0, After : 0});
		p2.SetParagraphSpacing({Before : 0, After : 0});

		Recalculate();
		assert.strictEqual(dc.GetElementsCount(), 2, "Check paragraphs count");
		assert.strictEqual(p1.GetPagesCount(), 1, "Check pages count of the first paragraph");
		assert.strictEqual(p2.GetPagesCount(), 1, "Check pages count of the second paragraph");

		assert.deepEqual(p1.GetPageBounds(0), new AscWord.CDocumentBounds(0, 0, pageWidth, AscTest.FontHeight), "Check page bounds of the first paragraph");
		assert.deepEqual(p2.GetPageBounds(0), new AscWord.CDocumentBounds(0, AscTest.FontHeight, pageWidth, AscTest.FontHeight * 2), "Check page bounds of the second paragraph");


		p1.SetParagraphSpacing({Before : 15, After : 20});
		p2.SetParagraphSpacing({Before : 0, After : 0});

		Recalculate();
		assert.deepEqual(p1.GetPageBounds(0), new AscWord.CDocumentBounds(0, 0, pageWidth, AscTest.FontHeight + 35), "Check page bounds of the first paragraph");
		assert.deepEqual(p2.GetPageBounds(0), new AscWord.CDocumentBounds(0, AscTest.FontHeight + 35, pageWidth, AscTest.FontHeight * 2 + 35), "Check page bounds of the second paragraph");

		p1.SetParagraphSpacing({Before : 15, After : 20});
		p2.SetParagraphSpacing({Before : 30, After : 0});

		Recalculate();
		assert.deepEqual(p1.GetPageBounds(0), new AscWord.CDocumentBounds(0, 0, pageWidth, AscTest.FontHeight + 35), "Check page bounds of the first paragraph");
		assert.deepEqual(p2.GetPageBounds(0), new AscWord.CDocumentBounds(0, AscTest.FontHeight + 35, pageWidth, AscTest.FontHeight * 2 + 45), "Check page bounds of the second paragraph");
	});
});
