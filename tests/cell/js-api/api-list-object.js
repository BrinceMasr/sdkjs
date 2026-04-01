/*
 * (c) Copyright Ascensio System SIA 2010-2026
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

$(function ()
{
	var ws = AscTest.JsApi.GetActiveSheet();

	function initializeTest()
	{
		var tables = ws.GetListObjects();
		for (var i = 0; i < tables.length; i++)
		{
			tables[i].Delete();
		}
	}

	QUnit.module("ApiListObject", function ()
	{
		QUnit.test("Active", function (assert)
		{
			initializeTest();

			var tbl = ws.AddListObject("xlSrcRange", "B2:D5");

			ws.GetRange("C3").Select();
			assert.equal(tbl.GetActive(), true, "GetActive returns true for cell inside table");
			assert.equal(tbl.Active, true, "Active property returns true");

			ws.GetRange("B2").Select();
			assert.equal(tbl.GetActive(), true, "Active is true for header row cell");

			ws.GetRange("A1").Select();
			assert.equal(tbl.GetActive(), false, "GetActive returns false for cell outside table");
			assert.equal(tbl.Active, false, "Active property returns false");

			ws.GetRange("E2").Select();
			assert.equal(tbl.GetActive(), false, "Active is false for cell just outside table column");

			ws.GetRange("B6").Select();
			assert.equal(tbl.GetActive(), false, "Active is false for cell just below table");
		});

		QUnit.test("AlternativeText", function (assert)
		{
			initializeTest();

			var tbl = ws.AddListObject("xlSrcRange", "A1:C3");

			assert.equal(tbl.GetAlternativeText(), "", "AlternativeText is empty by default");

			tbl.SetAlternativeText("Sales data");
			assert.equal(tbl.GetAlternativeText(), "Sales data", "GetAlternativeText returns set value");
			assert.equal(tbl.AlternativeText, "Sales data", "AlternativeText property matches GetAlternativeText()");

			tbl.AlternativeText = "Quarterly report";
			assert.equal(tbl.GetAlternativeText(), "Quarterly report", "AlternativeText set via property");
		});

		QUnit.test("Comment", function (assert)
		{
			initializeTest();

			var tbl = ws.AddListObject("xlSrcRange", "A1:C3");

			assert.equal(tbl.GetComment(), "", "Comment is empty by default");

			tbl.SetComment("Annual sales summary");
			assert.equal(tbl.GetComment(), "Annual sales summary", "GetComment returns set value");
			assert.equal(tbl.Comment, "Annual sales summary", "Comment property matches GetComment()");

			tbl.Comment = "Q3 report data";
			assert.equal(tbl.GetComment(), "Q3 report data", "Comment set via property");
		});

		QUnit.test("DataBodyRange", function (assert)
		{
			initializeTest();

			var tbl = ws.AddListObject("xlSrcRange", "B2:D5");
			assert.equal(tbl.GetDataBodyRange().GetAddress(true, true), "$B$3:$D$5", "DataBodyRange starts after header row");
			assert.equal(tbl.DataBodyRange.GetAddress(true, true), "$B$3:$D$5", "DataBodyRange property matches GetDataBodyRange()");

			initializeTest();

			tbl = ws.AddListObject("xlSrcRange", "B2:D5", false, "xlNo");
			assert.equal(tbl.GetDataBodyRange().GetAddress(true, true), "$B$2:$D$5", "DataBodyRange covers full range when no header");
		});

		QUnit.test("GetListObjects", function (assert)
		{
			initializeTest();

			assert.equal(ws.GetListObjects().length, 0, "No tables on a clean sheet");

			var tbl1 = ws.AddListObject("xlSrcRange", "A1:C3");
			var tbl2 = ws.AddListObject("xlSrcRange", "E1:G3");
			var tables = ws.GetListObjects();

			assert.equal(tables.length, 2, "GetListObjects returns 2 tables");
			assert.equal(tables[0].GetName(), tbl1.GetName(), "First table name matches");
			assert.equal(tables[1].GetName(), tbl2.GetName(), "Second table name matches");
		});

		QUnit.test("AddListObject", function (assert)
		{
			initializeTest();

			var tbl = ws.AddListObject("xlSrcRange", "A1:D5");
			assert.ok(tbl !== null, "AddListObject returns non-null");

			var name = tbl.GetName();
			assert.ok(name.length > 0, "Table has a non-empty name");
			assert.equal(tbl.Name, name, "Name property matches GetName()");

			initializeTest();

			ws.GetRange("A1").SetValue(10);
			tbl = ws.AddListObject("xlSrcRange", "A1:B3", false, "xlNo");
			assert.ok(tbl !== null, "Table created with xlNo headers");

			initializeTest();

			tbl = ws.AddListObject("xlSrcRange", "A1:C4", false, "xlYes", null, "TableStyleMedium9");
			assert.ok(tbl !== null, "Table created with custom style");
		});

		QUnit.test("GetRange", function (assert)
		{
			initializeTest();

			var tbl = ws.AddListObject("xlSrcRange", "B2:E6");
			var range = tbl.GetRange();

			assert.ok(range !== null, "GetRange returns non-null");
			assert.equal(range.GetAddress(true, true), "$B$2:$E$6", "GetRange returns correct address");
			assert.equal(tbl.Range.GetAddress(true, true), range.GetAddress(true, true), "Range property matches GetRange()");
		});

		QUnit.test("Delete", function (assert)
		{
			initializeTest();

			var tbl = ws.AddListObject("xlSrcRange", "A1:D5");
			assert.equal(ws.GetListObjects().length, 1, "One table before Delete");
			tbl.Delete();
			assert.equal(ws.GetListObjects().length, 0, "No tables after Delete");

			var tbl1 = ws.AddListObject("xlSrcRange", "A1:C3");
			var tbl2 = ws.AddListObject("xlSrcRange", "E1:G3");
			var name2 = tbl2.GetName();
			tbl1.Delete();
			var tables = ws.GetListObjects();
			assert.equal(tables.length, 1, "One table remains after deleting the first");
			assert.equal(tables[0].GetName(), name2, "Remaining table is the correct one");
		});

		QUnit.test("Unlist", function (assert)
		{
			initializeTest();

			ws.GetRange("A1").SetValue("Header");
			ws.GetRange("A2").SetValue(1);
			ws.GetRange("A3").SetValue(2);

			var tbl = ws.AddListObject("xlSrcRange", "A1:A3");
			tbl.Unlist();

			assert.equal(ws.GetListObjects().length, 0, "No tables after Unlist");
			assert.equal(ws.GetRange("A2").GetValue(), "1", "Cell data preserved after Unlist");
		});

		QUnit.test("Resize", function (assert)
		{
			initializeTest();

			var tbl = ws.AddListObject("xlSrcRange", "A1:C3");
			tbl.Resize("A1:C6");
			assert.equal(ws.GetListObjects()[0].GetRange().GetAddress(true, true), "$A$1:$C$6", "Range is $A$1:$C$6 after Resize to larger range");

			tbl.Resize("A1:B5");
			assert.equal(ws.GetListObjects()[0].GetRange().GetAddress(true, true), "$A$1:$B$5", "Range is $A$1:$B$5 after Resize to smaller range");
		});
	});
});
