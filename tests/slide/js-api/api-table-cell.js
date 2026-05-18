/*
 * (c) Copyright Ascensio System SIA 2010-2025
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

"use strict";

$(function () {

	const logicDocument = AscTest.CreateLogicDocument();
	QUnit.module("ApiTableCell");

	function CreateSlide()
	{
		logicDocument.addNextSlide(0);
		editor.WordControl.Thumbnails.CalculatePlaces();
	}

	function createTableCell(text)
	{
		CreateSlide();

		const presentation = AscTest.JsApi.GetPresentation();
		const slide        = presentation.GetSlideByIndex(0);
		const table        = AscTest.JsApi.CreateTable(1, 1);
		slide.AddObject(table);

		const cell = table.GetRow(0).GetCell(0);
		if (typeof text === "string")
			cell.SetText(text);

		return cell;
	}

	// ── ApiTableCell.GetTextRange ─────────────────────────────────────────────

	QUnit.test("Test: ApiTableCell.GetTextRange returns an ApiTextRange", function (assert) {
		const cell  = createTableCell();
		const range = cell.GetTextRange();
		assert.ok(range !== null, "GetTextRange returns non-null");
		assert.strictEqual(range.GetClassType(), "textRange", "GetClassType returns 'textRange'");
	});

	QUnit.test("Test: ApiTableCell.GetTextRange reflects cell text content", function (assert) {
		const cell  = createTableCell("Hello");
		const range = cell.GetTextRange();
		assert.ok(range !== null, "GetTextRange returns non-null after SetText");
		assert.strictEqual(range.GetText(), "Hello\r\n", "GetText returns the cell text");
	});

	QUnit.test("Test: ApiTableCell.GetTextRange allows text modification via SetText", function (assert) {
		const cell  = createTableCell("Before");
		const range = cell.GetTextRange();
		range.SetText("After");
		assert.strictEqual(cell.GetTextRange().GetText(), "After\r\n", "SetText on range updates the cell text");
	});

	QUnit.test("Test: ApiTableCell.GetTextRange GetStartPos and GetEndPos", function (assert) {
		const cell  = createTableCell("Hi");
		const range = cell.GetTextRange();
		assert.strictEqual(range.GetStartPos(), 0, "GetStartPos is 0 for full-range");
		assert.strictEqual(range.GetEndPos(), 4, "GetEndPos equals text length (2 chars + \\r\\n)");
	});

});
