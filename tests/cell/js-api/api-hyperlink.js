/*
 * (c) Copyright Ascensio System SIA 2010-2024
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

AscCommonExcel.WorkbookView.prototype.restoreFocus = function () {};

(function (window)
{
	QUnit.module("ApiHyperlink");

	let ws;

	QUnit.testStart(function () {
		ws = AscTest.JsApi.GetActiveSheet();
	});

	function checkUndoRedo(assert, fBefore, fAfter, desc) {
		fAfter(assert, "after_" + desc);
		AscCommon.History.Undo();
		fBefore(assert, "undo_" + desc);
		AscCommon.History.Redo();
		fAfter(assert, "redo_" + desc);
		AscCommon.History.Undo();
	}

	QUnit.test("GetClassType", function (assert) {
		ws.SetHyperlink("A1", "https://onlyoffice.com", null, "tip", "link");
		assert.strictEqual(ws.GetHyperlinks()[0].GetClassType(), "hyperlink", "GetClassType returns 'hyperlink'");
	});

	QUnit.test("GetAddress", function (assert) {
		ws.SetHyperlink("A1", "https://onlyoffice.com", null, "tip", "link");
		assert.strictEqual(ws.GetHyperlinks()[0].GetAddress(), "https://onlyoffice.com", "GetAddress returns URL");
	});

	QUnit.test("SetAddress", function (assert) {
		ws.SetHyperlink("A1", "https://onlyoffice.com", null, "tip", "link");
		ws.GetHyperlinks()[0].SetAddress("https://helpcenter.onlyoffice.com");
		checkUndoRedo(assert,
			function (a, d) { a.strictEqual(ws.GetHyperlinks()[0].GetAddress(), "https://onlyoffice.com", d); },
			function (a, d) { a.strictEqual(ws.GetHyperlinks()[0].GetAddress(), "https://helpcenter.onlyoffice.com", d); },
			"SetAddress"
		);
	});

	QUnit.test("GetSubAddress", function (assert) {
		ws.SetHyperlink("A1", "", "Sheet1!B5", "tip", "link");
		assert.strictEqual(ws.GetHyperlinks()[0].GetSubAddress(), "Sheet1!B5", "GetSubAddress returns location");
	});

	QUnit.test("SetSubAddress", function (assert) {
		ws.SetHyperlink("A1", "", "Sheet1!B5", "tip", "link");
		ws.GetHyperlinks()[0].SetSubAddress("Sheet1!D10");
		checkUndoRedo(assert,
			function (a, d) { a.strictEqual(ws.GetHyperlinks()[0].GetSubAddress(), "Sheet1!B5", d); },
			function (a, d) { a.strictEqual(ws.GetHyperlinks()[0].GetSubAddress(), "Sheet1!D10", d); },
			"SetSubAddress"
		);
	});

	QUnit.test("GetScreenTip", function (assert) {
		ws.SetHyperlink("A1", "https://onlyoffice.com", null, "ONLYOFFICE website", "link");
		assert.strictEqual(ws.GetHyperlinks()[0].GetScreenTip(), "ONLYOFFICE website", "GetScreenTip returns tooltip");
	});

	QUnit.test("SetScreenTip", function (assert) {
		ws.SetHyperlink("A1", "https://onlyoffice.com", null, "ONLYOFFICE website", "link");
		ws.GetHyperlinks()[0].SetScreenTip("Go to ONLYOFFICE Help Center");
		checkUndoRedo(assert,
			function (a, d) { a.strictEqual(ws.GetHyperlinks()[0].GetScreenTip(), "ONLYOFFICE website", d); },
			function (a, d) { a.strictEqual(ws.GetHyperlinks()[0].GetScreenTip(), "Go to ONLYOFFICE Help Center", d); },
			"SetScreenTip"
		);
	});

	QUnit.test("GetTextToDisplay", function (assert) {
		ws.SetHyperlink("A1", "https://onlyoffice.com", null, "tip", "Visit ONLYOFFICE");
		assert.strictEqual(ws.GetHyperlinks()[0].GetTextToDisplay(), "Visit ONLYOFFICE", "GetTextToDisplay returns cell text");
	});

	QUnit.test("SetTextToDisplay", function (assert) {
		ws.SetHyperlink("A1", "https://onlyoffice.com", null, "tip", "Visit ONLYOFFICE");
		ws.GetHyperlinks()[0].SetTextToDisplay("Go to ONLYOFFICE");
		checkUndoRedo(assert,
			function (a, d) { a.strictEqual(ws.GetHyperlinks()[0].GetTextToDisplay(), "Visit ONLYOFFICE", d); },
			function (a, d) { a.strictEqual(ws.GetHyperlinks()[0].GetTextToDisplay(), "Go to ONLYOFFICE", d); },
			"SetTextToDisplay"
		);
	});

	QUnit.test("GetRange", function (assert) {
		ws.SetHyperlink("B2", "https://onlyoffice.com", null, "tip", "link");
		let range = ws.GetHyperlinks()[0].GetRange();
		assert.ok(range, "GetRange returns a range object");
		assert.strictEqual(range.Row, 2, "row is 2");
		assert.strictEqual(range.Col, 2, "col is 2 (column B)");
	});

	QUnit.test("GetName", function (assert) {
		ws.SetHyperlink("A1", "https://onlyoffice.com", null, "tip", "Visit ONLYOFFICE");
		assert.strictEqual(ws.GetHyperlinks()[0].GetName(), "Visit ONLYOFFICE", "GetName returns display text");
	});

	QUnit.test("GetType", function (assert) {
		ws.SetHyperlink("A1", "https://onlyoffice.com", null, "tip", "link");
		assert.strictEqual(ws.GetHyperlinks()[0].GetType(), Asc.c_oAscMsoHyperlinkType.Range, "GetType returns msoHyperlinkRange (0)");
	});

	QUnit.test("Delete", function (assert) {
		ws.SetHyperlink("A1", "https://onlyoffice.com", null, "tip", "link");
		ws.GetHyperlinks()[0].Delete();
		checkUndoRedo(assert,
			function (a, d) { a.strictEqual(ws.GetHyperlinks().length, 1, d + ": hyperlink restored after undo"); },
			function (a, d) { a.strictEqual(ws.GetHyperlinks().length, 0, d + ": hyperlink deleted"); },
			"Delete"
		);
	});

	QUnit.test("ApiWorksheet.GetHyperlinks", function (assert) {
		ws.SetHyperlink("A1", "https://onlyoffice.com", null, "tip1", "link1");
		ws.SetHyperlink("B1", "https://helpcenter.onlyoffice.com", null, "tip2", "link2");
		let hyperlinks = ws.GetHyperlinks();
		assert.strictEqual(hyperlinks.length, 2, "GetHyperlinks returns 2 hyperlinks");
		assert.strictEqual(hyperlinks[0].GetClassType(), "hyperlink", "first item is ApiHyperlink");
		assert.strictEqual(hyperlinks[1].GetClassType(), "hyperlink", "second item is ApiHyperlink");
	});

	QUnit.test("ApiRange.GetHyperlinks", function (assert) {
		ws.SetHyperlink("A1", "https://onlyoffice.com", null, "tip1", "link1");
		ws.SetHyperlink("C1", "https://helpcenter.onlyoffice.com", null, "tip2", "link2");
		let hyperlinks = ws.GetRange("A1").GetHyperlinks();
		assert.strictEqual(hyperlinks.length, 1, "GetHyperlinks on range A1 returns 1 hyperlink");
		assert.strictEqual(hyperlinks[0].GetAddress(), "https://onlyoffice.com", "correct hyperlink in range");
	});

})(window);
