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
	
	QUnit.test("Test temporary content control", function (assert)
	{
		let dateTime, p;
		function initDocument()
		{
			AscTest.ClearDocument();
			p = AscTest.CreateParagraph();
			logicDocument.AddToContent(0, p);
			AscTest.MoveCursorToParagraph(p);
			
			dateTime = logicDocument.AddContentControlDatePicker();
			dateTime.SetContentControlTemporary(true);
		}
		
		initDocument();
		assert.strictEqual(dateTime.IsUseInDocument(), true, "Check if date-time is added to the document");
		assert.strictEqual(AscTest.GetParagraphText(p), "Enter a date");
		dateTime.MoveCursorToContentControl();
		AscTest.EnterText("123");
		assert.strictEqual(dateTime.IsUseInDocument(), false, "Check if date-time is in the document after adding text");
		assert.strictEqual(AscTest.GetParagraphText(p), "123");
		
		initDocument();
		
		let dateTimePr = new AscWord.CSdtDatePickerPr();
		let date = new Date();
		date.setFullYear(2024, 6, 24);
		dateTimePr.SetDateFormat("mm/dd/yyyy");
		dateTimePr.SetFullDate(date);
		
		dateTime.SelectContentControl();
		dateTime.ApplyDatePickerPr(dateTimePr, true);
		assert.strictEqual(dateTime.IsUseInDocument(), false, "Check if date-time is in the document");
		assert.strictEqual(AscTest.GetParagraphText(p), "07/24/2024");
	});
	
});
