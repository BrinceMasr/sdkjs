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

		QUnit.test("AutoFilter", function (assert)
		{
			initializeTest();

			var tbl = ws.AddListObject("xlSrcRange", "B2:D5");

			var af = tbl.GetAutoFilter();
			assert.ok(af !== null, "GetAutoFilter returns non-null for a table with autofilter");
			assert.ok(tbl.AutoFilter !== null, "AutoFilter property returns non-null");

			assert.equal(af.GetFilterMode(), true, "FilterMode is true when autofilter exists");
			assert.equal(af.GetFilters().length, 0, "GetFilters returns empty array when no filters applied");
			assert.equal(af.GetRange().GetAddress(true, true), "$B$2:$D$5", "AutoFilter range matches table range");
			assert.equal(af.GetParent().GetName(), tbl.GetName(), "GetParent returns the parent ListObject");

			// ShowAutoFilter = false removes AutoFilter → GetAutoFilter returns null
			tbl.SetShowAutoFilter(false);
			assert.equal(tbl.GetAutoFilter(), null, "GetAutoFilter returns null after ShowAutoFilter set to false");
		});

		QUnit.test("ShowAutoFilter and ShowAutoFilterDropDown", function (assert)
		{
			initializeTest();

			var tbl = ws.AddListObject("xlSrcRange", "A1:C4");

			// defaults
			assert.equal(tbl.GetShowAutoFilter(), true, "ShowAutoFilter is true by default");
			assert.equal(tbl.ShowAutoFilter, true, "ShowAutoFilter property matches GetShowAutoFilter()");
			assert.equal(tbl.GetShowAutoFilterDropDown(), true, "ShowAutoFilterDropDown is true by default");
			assert.equal(tbl.ShowAutoFilterDropDown, true, "ShowAutoFilterDropDown property matches GetShowAutoFilterDropDown()");

			// ShowAutoFilter = false removes AutoFilter entirely → both become false
			tbl.SetShowAutoFilter(false);
			assert.equal(tbl.GetShowAutoFilter(), false, "ShowAutoFilter is false after removing AutoFilter");
			assert.equal(tbl.GetShowAutoFilterDropDown(), false, "ShowAutoFilterDropDown is false when AutoFilter is removed");

			// ShowAutoFilter = true restores AutoFilter
			tbl.ShowAutoFilter = true;
			assert.equal(tbl.GetShowAutoFilter(), true, "ShowAutoFilter restored to true");

			// ShowAutoFilterDropDown = false hides buttons only, AutoFilter stays
			tbl.SetShowAutoFilterDropDown(false);
			assert.equal(tbl.GetShowAutoFilter(), true, "ShowAutoFilter stays true when only buttons hidden");
			assert.equal(tbl.GetShowAutoFilterDropDown(), false, "ShowAutoFilterDropDown is false after hiding buttons");

			tbl.ShowAutoFilterDropDown = true;
			assert.equal(tbl.GetShowAutoFilterDropDown(), true, "ShowAutoFilterDropDown restored via property setter");
		});

		QUnit.test("ShowHeaders", function (assert)
		{
			initializeTest();

			var tbl = ws.AddListObject("xlSrcRange", "B2:D5");
			assert.equal(tbl.GetShowHeaders(), true, "ShowHeaders is true by default");
			assert.equal(tbl.ShowHeaders, true, "ShowHeaders property matches GetShowHeaders()");

			tbl.SetShowHeaders(false);
			assert.equal(tbl.GetShowHeaders(), false, "ShowHeaders is false after SetShowHeaders(false)");

			tbl.ShowHeaders = true;
			assert.equal(tbl.GetShowHeaders(), true, "ShowHeaders restored to true via property setter");

			// calling SetShowHeaders with current value should not change state
			tbl.SetShowHeaders(true);
			assert.equal(tbl.GetShowHeaders(), true, "SetShowHeaders(true) when already true does not change state");
		});

		QUnit.test("ShowTableStyleColumnStripes", function (assert)
		{
			initializeTest();

			var tbl = ws.AddListObject("xlSrcRange", "A1:C4");
			assert.equal(tbl.GetShowTableStyleColumnStripes(), false, "ShowTableStyleColumnStripes is false by default");
			assert.equal(tbl.ShowTableStyleColumnStripes, false, "ShowTableStyleColumnStripes property matches getter");

			tbl.SetShowTableStyleColumnStripes(true);
			assert.equal(tbl.GetShowTableStyleColumnStripes(), true, "ShowTableStyleColumnStripes is true after setter");

			tbl.ShowTableStyleColumnStripes = false;
			assert.equal(tbl.GetShowTableStyleColumnStripes(), false, "ShowTableStyleColumnStripes reset via property");

			tbl.SetShowTableStyleColumnStripes(false);
			assert.equal(tbl.GetShowTableStyleColumnStripes(), false, "SetShowTableStyleColumnStripes(false) when already false does not toggle");
		});

		QUnit.test("ShowTableStyleFirstColumn, ShowTableStyleLastColumn, ShowTableStyleRowStripes", function (assert)
		{
			initializeTest();

			var tbl = ws.AddListObject("xlSrcRange", "A1:C4");

			assert.equal(tbl.GetShowTableStyleFirstColumn(), false, "ShowTableStyleFirstColumn is false by default");
			tbl.SetShowTableStyleFirstColumn(true);
			assert.equal(tbl.GetShowTableStyleFirstColumn(), true, "ShowTableStyleFirstColumn is true after setter");
			tbl.ShowTableStyleFirstColumn = false;
			assert.equal(tbl.GetShowTableStyleFirstColumn(), false, "ShowTableStyleFirstColumn reset via property");
			tbl.SetShowTableStyleFirstColumn(false);
			assert.equal(tbl.GetShowTableStyleFirstColumn(), false, "SetShowTableStyleFirstColumn(false) when already false does not toggle");

			assert.equal(tbl.GetShowTableStyleLastColumn(), false, "ShowTableStyleLastColumn is false by default");
			tbl.SetShowTableStyleLastColumn(true);
			assert.equal(tbl.GetShowTableStyleLastColumn(), true, "ShowTableStyleLastColumn is true after setter");
			tbl.ShowTableStyleLastColumn = false;
			assert.equal(tbl.GetShowTableStyleLastColumn(), false, "ShowTableStyleLastColumn reset via property");

			assert.equal(tbl.GetShowTableStyleRowStripes(), true, "ShowTableStyleRowStripes is true by default");
			tbl.SetShowTableStyleRowStripes(false);
			assert.equal(tbl.GetShowTableStyleRowStripes(), false, "ShowTableStyleRowStripes is false after setter");
			tbl.ShowTableStyleRowStripes = true;
			assert.equal(tbl.GetShowTableStyleRowStripes(), true, "ShowTableStyleRowStripes restored via property");
		});

		QUnit.test("ShowTotals", function (assert)
		{
			initializeTest();

			var tbl = ws.AddListObject("xlSrcRange", "A1:C4");

			assert.equal(tbl.GetShowTotals(), false, "ShowTotals is false by default");
			assert.equal(tbl.ShowTotals, false, "ShowTotals property matches getter");

			tbl.SetShowTotals(true);
			assert.equal(tbl.GetShowTotals(), true, "ShowTotals is true after setter");

			tbl.ShowTotals = false;
			assert.equal(tbl.GetShowTotals(), false, "ShowTotals reset via property");

			tbl.SetShowTotals(false);
			assert.equal(tbl.GetShowTotals(), false, "SetShowTotals(false) when already false does not toggle");
		});

		QUnit.test("HeaderRowRange", function (assert)
		{
			initializeTest();

			// xlYes: header row is B2
			var tbl = ws.AddListObject("xlSrcRange", "B2:D5");
			assert.equal(tbl.GetHeaderRowRange().GetAddress(true, true), "$B$2:$D$2", "HeaderRowRange returns the first row of the table");
			assert.equal(tbl.HeaderRowRange.GetAddress(true, true), "$B$2:$D$2", "HeaderRowRange property matches GetHeaderRowRange()");

			initializeTest();

			// xlNo: auto-header is inserted at B2, table expands to B2:D6
			tbl = ws.AddListObject("xlSrcRange", "B2:D5", false, "xlNo");
			assert.equal(tbl.GetHeaderRowRange().GetAddress(true, true), "$B$2:$D$2", "HeaderRowRange returns auto-generated header row for xlNo table");

			initializeTest();

			// table with no header: HeaderRowRange returns null
			tbl = ws.AddListObject("xlSrcRange", "B2:D5", false, "xlYes");
			tbl.tablePart.HeaderRowCount = 0;
			assert.equal(tbl.GetHeaderRowRange(), null, "HeaderRowRange returns null when table has no header row");
		});

		QUnit.test("DataBodyRange", function (assert)
		{
			initializeTest();

			// xlYes: B2:D5, header at B2, data at B3:D5
			var tbl = ws.AddListObject("xlSrcRange", "B2:D5");
			assert.equal(tbl.GetDataBodyRange().GetAddress(true, true), "$B$3:$D$5", "DataBodyRange starts after header row");
			assert.equal(tbl.DataBodyRange.GetAddress(true, true), "$B$3:$D$5", "DataBodyRange property matches GetDataBodyRange()");

			initializeTest();

			// xlNo: range expands by 1 row (B2:D5 → B2:D6), auto-header at B2, data shifted to B3:D6
			tbl = ws.AddListObject("xlSrcRange", "B2:D5", false, "xlNo");
			assert.equal(tbl.GetRange().GetAddress(true, true), "$B$2:$D$6", "xlNo expands table range by 1 row");
			assert.equal(tbl.GetDataBodyRange().GetAddress(true, true), "$B$3:$D$6", "DataBodyRange starts after auto-generated header row");
		});

		QUnit.test("DisplayName", function (assert)
		{
			initializeTest();

			var tbl = ws.AddListObject("xlSrcRange", "A1:C3");
			var originalName = tbl.GetDisplayName();
			assert.ok(originalName.length > 0, "DisplayName is non-empty by default");
			assert.equal(tbl.DisplayName, originalName, "DisplayName property matches GetDisplayName()");

			tbl.SetDisplayName("SalesTable");
			assert.equal(tbl.GetDisplayName(), "SalesTable", "GetDisplayName returns new name after SetDisplayName");
			assert.equal(ws.GetListObjects()[0].GetDisplayName(), "SalesTable", "Rename is reflected in GetListObjects");

			tbl.DisplayName = "RevenueTable";
			assert.equal(tbl.GetDisplayName(), "RevenueTable", "DisplayName set via property");
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

			tbl.SetName("MyTable");
			assert.equal(tbl.GetName(), "MyTable", "GetName returns new name after SetName");
			tbl.Name = "MyTable2";
			assert.equal(tbl.GetName(), "MyTable2", "Name set via property setter");

			initializeTest();

			// xlNo: range A1:B3 expands to A1:B4 (auto-header inserted, data shifts down)
			ws.GetRange("A1").SetValue(10);
			tbl = ws.AddListObject("xlSrcRange", "A1:B3", false, "xlNo");
			assert.ok(tbl !== null, "Table created with xlNo headers");
			assert.equal(tbl.GetRange().GetAddress(true, true), "$A$1:$B$4", "xlNo expands table range by 1 row");

			initializeTest();

			tbl = ws.AddListObject("xlSrcRange", "A1:C4", false, "xlYes", null, "TableStyleMedium9");
			assert.ok(tbl !== null, "Table created with custom style");
		});

		QUnit.test("Parent", function (assert)
		{
			initializeTest();

			var tbl = ws.AddListObject("xlSrcRange", "A1:C3");

			assert.ok(tbl.GetParent() !== null, "GetParent returns non-null");
			assert.equal(tbl.GetParent().GetName(), ws.GetName(), "GetParent returns the correct worksheet");
			assert.equal(tbl.Parent.GetName(), ws.GetName(), "Parent property matches GetParent()");
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

		QUnit.test("SourceType", function (assert)
		{
			initializeTest();

			var tbl = ws.AddListObject("xlSrcRange", "A1:C3");
			assert.equal(tbl.GetSourceType(), "xlSrcRange", "GetSourceType returns xlSrcRange");
			assert.equal(tbl.SourceType, "xlSrcRange", "SourceType property matches GetSourceType()");
		});

		QUnit.test("TableStyle", function (assert)
		{
			initializeTest();

			var tbl = ws.AddListObject("xlSrcRange", "A1:C3", false, "xlYes", null, "TableStyleMedium9");
			assert.equal(tbl.GetTableStyle(), "TableStyleMedium9", "GetTableStyle returns style set at creation");
			assert.equal(tbl.TableStyle, "TableStyleMedium9", "TableStyle property matches GetTableStyle()");

			tbl.SetTableStyle("TableStyleLight1");
			assert.equal(tbl.GetTableStyle(), "TableStyleLight1", "GetTableStyle returns new style after SetTableStyle");

			tbl.TableStyle = "TableStyleDark1";
			assert.equal(tbl.GetTableStyle(), "TableStyleDark1", "TableStyle set via property");
		});

		QUnit.test("TotalsRowRange", function (assert)
		{
			initializeTest();

			var tbl = ws.AddListObject("xlSrcRange", "B2:D5");
			assert.equal(tbl.GetTotalsRowRange(), null, "TotalsRowRange is null when ShowTotals is false");

			tbl.SetShowTotals(true);
			var range = tbl.GetTotalsRowRange();
			assert.ok(range !== null, "TotalsRowRange is non-null after ShowTotals = true");
			assert.equal(range.GetAddress(true, true), "$B$6:$D$6", "TotalsRowRange is the last row of the table");
			assert.equal(tbl.TotalsRowRange.GetAddress(true, true), "$B$6:$D$6", "TotalsRowRange property matches getter");
		});

		QUnit.test("Summary", function (assert)
		{
			initializeTest();

			var tbl = ws.AddListObject("xlSrcRange", "A1:C3");
			assert.equal(tbl.GetSummary(), "", "Summary is empty by default");

			tbl.SetSummary("Annual report data");
			assert.equal(tbl.GetSummary(), "Annual report data", "GetSummary returns set value");
			assert.equal(tbl.Summary, "Annual report data", "Summary property matches GetSummary()");
			assert.equal(tbl.GetComment(), "Annual report data", "Summary and Comment share the same value");

			tbl.Summary = "Q4 data";
			assert.equal(tbl.GetSummary(), "Q4 data", "Summary set via property");
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

		QUnit.test("GetListColumns", function (assert)
		{
			initializeTest();

			ws.GetRange("B2").SetValue("Name");
			ws.GetRange("C2").SetValue("Age");
			ws.GetRange("D2").SetValue("Score");
			var tbl = ws.AddListObject("xlSrcRange", "B2:D5");
			var cols = tbl.GetListColumns();

			assert.equal(cols.length, 3, "GetListColumns returns array with 3 elements");
			assert.equal(cols[0].GetName(), "Name",  "cols[0] is Name");
			assert.equal(cols[1].GetName(), "Age",   "cols[1] is Age");
			assert.equal(cols[2].GetName(), "Score", "cols[2] is Score");
		});

		QUnit.test("ListColumn - Index, Name, Parent, Range, DataBodyRange", function (assert)
		{
			initializeTest();

			ws.GetRange("B2").SetValue("Name");
			ws.GetRange("C2").SetValue("Age");
			ws.GetRange("D2").SetValue("Score");
			var tbl = ws.AddListObject("xlSrcRange", "B2:D5");
			var cols = tbl.GetListColumns();
			var col1 = cols[0];
			var col2 = cols[1];
			var col3 = cols[2];

			assert.equal(col1.GetIndex(), 1, "GetIndex returns 1 for first column");
			assert.equal(col2.GetIndex(), 2, "GetIndex returns 2 for second column");
			assert.equal(col3.Index, 3, "Index property returns 3 for third column");

			assert.strictEqual(col1.GetParent(), tbl, "GetParent returns the parent ApiListObject");
			assert.strictEqual(col1.Parent, tbl, "Parent property matches GetParent");

			assert.equal(col2.GetName(), "Age", "GetName returns Age");
			col2.SetName("Years");
			assert.equal(col2.GetName(), "Years", "GetName returns updated name after SetName");
			col2.Name = "Age";
			assert.equal(col2.Name, "Age", "Name property setter and getter work");

			// Table B2:D5, header row 2, data rows 3-5
			assert.equal(col1.GetRange().GetAddress(true, true), "$B$2:$B$5", "Range for first column includes header");
			assert.equal(col2.GetRange().GetAddress(true, true), "$C$2:$C$5", "Range for second column");
			assert.equal(col1.Range.GetAddress(true, true), "$B$2:$B$5", "Range property works");

			assert.equal(col1.GetDataBodyRange().GetAddress(true, true), "$B$3:$B$5", "DataBodyRange excludes header row");
			assert.equal(col2.DataBodyRange.GetAddress(true, true), "$C$3:$C$5", "DataBodyRange property works");
		});

		QUnit.test("ListColumn - TotalsCalculation, Total", function (assert)
		{
			initializeTest();

			ws.GetRange("B2").SetValue("Name");
			ws.GetRange("C2").SetValue("Age");
			ws.GetRange("D2").SetValue("Score");
			var tbl = ws.AddListObject("xlSrcRange", "B2:D5");
			var cols = tbl.GetListColumns();
			var col1 = cols[0];
			var col2 = cols[1];
			var col3 = cols[2];

			assert.equal(col2.GetTotalsCalculation(), "xlTotalsCalculationNone", "Default TotalsCalculation is None");
			assert.equal(col1.GetTotal(), null, "Total is null when table has no totals row");

			tbl.SetShowTotals(true);

			// when totals row is shown: Sum function is auto-assigned to last column
			assert.equal(col3.GetTotalsCalculation(), "xlTotalsCalculationSum", "Last column gets Sum when totals shown");
			assert.equal(col2.GetTotalsCalculation(), "xlTotalsCalculationNone", "Middle column stays None");

			// Table B2:D5 + totals row → B2:D6, totals row is row 6
			assert.equal(col1.GetTotal().GetAddress(true, true), "$B$6", "Total for first column is the totals row cell");
			assert.equal(col3.Total.GetAddress(true, true), "$D$6", "Total property for third column");

			col2.SetTotalsCalculation("xlTotalsCalculationCount");
			assert.equal(col2.GetTotalsCalculation(), "xlTotalsCalculationCount", "TotalsCalculation is Count after SetTotalsCalculation");
			assert.equal(col2.TotalsCalculation, "xlTotalsCalculationCount", "TotalsCalculation property getter works");

			col2.TotalsCalculation = "xlTotalsCalculationAverage";
			assert.equal(col2.TotalsCalculation, "xlTotalsCalculationAverage", "TotalsCalculation set via property");
		});

		QUnit.test("AddListColumn", function (assert)
		{
			initializeTest();

			ws.GetRange("B2").SetValue("Name");
			ws.GetRange("C2").SetValue("Age");
			ws.GetRange("D2").SetValue("Score");
			var tbl = ws.AddListObject("xlSrcRange", "B2:D5");

			assert.equal(tbl.GetListColumns().length, 3, "Initial count is 3");

			var newCol = tbl.AddListColumn();
			assert.equal(tbl.GetListColumns().length, 4, "Count is 4 after AddListColumn()");
			assert.ok(newCol !== null, "AddListColumn() returns the new column");
			assert.equal(newCol.GetIndex(), 4, "Appended column index is 4");

			var insertedCol = tbl.AddListColumn(2);
			assert.equal(tbl.GetListColumns().length, 5, "Count is 5 after AddListColumn(2)");
			assert.equal(insertedCol.GetIndex(), 2, "Inserted column index is 2");
			assert.equal(tbl.GetListColumns()[2].GetName(), "Age", "Former second column is now third");
		});

		QUnit.test("ListColumn - Delete", function (assert)
		{
			initializeTest();

			ws.GetRange("B2").SetValue("Name");
			ws.GetRange("C2").SetValue("Age");
			ws.GetRange("D2").SetValue("Score");
			var tbl = ws.AddListObject("xlSrcRange", "B2:D5");

			assert.equal(tbl.GetListColumns().length, 3, "Initial count is 3");

			tbl.GetListColumns()[1].Delete();
			var remaining = tbl.GetListColumns();
			assert.equal(remaining.length, 2, "Count is 2 after deleting middle column");
			assert.equal(remaining[0].GetName(), "Name",  "First column is still Name");
			assert.equal(remaining[1].GetName(), "Score", "Second column is now Score");
		});

		QUnit.test("GetListRows", function (assert)
		{
			initializeTest();

			// Table B2:D5: header row 2, data rows 3-5
			var tbl = ws.AddListObject("xlSrcRange", "B2:D5");
			var rows = tbl.GetListRows();

			assert.equal(rows.length, 3, "GetListRows returns 3 data rows");
			assert.equal(rows[0].GetIndex(), 1, "rows[0].GetIndex() is 1");
			assert.equal(rows[1].GetIndex(), 2, "rows[1].GetIndex() is 2");
			assert.equal(rows[2].Index, 3, "rows[2].Index property is 3");

			assert.equal(rows[0].GetRange().GetAddress(true, true), "$B$3:$D$3", "First data row range");
			assert.equal(rows[2].GetRange().GetAddress(true, true), "$B$5:$D$5", "Last data row range");
			assert.equal(rows[1].Range.GetAddress(true, true), "$B$4:$D$4", "Range property works");

			assert.strictEqual(rows[0].GetParent(), tbl, "GetParent returns the parent ApiListObject");
			assert.strictEqual(rows[0].Parent, tbl, "Parent property works");

			// Totals row is excluded from data rows
			tbl.SetShowTotals(true);
			var rowsWithTotals = tbl.GetListRows();
			assert.equal(rowsWithTotals.length, 3, "GetListRows still returns 3 when totals row is shown");
			assert.equal(rowsWithTotals[2].GetRange().GetAddress(true, true), "$B$5:$D$5", "Last data row is still row 5, not the totals row");
		});

		QUnit.test("AddListRow", function (assert)
		{
			initializeTest();

			var tbl = ws.AddListObject("xlSrcRange", "B2:D5");

			assert.equal(tbl.GetListRows().length, 3, "Initial data row count is 3");

			// Append at end
			var newRow = tbl.AddListRow();
			assert.ok(newRow !== null, "AddListRow() returns the new row");
			assert.equal(tbl.GetListRows().length, 4, "Count is 4 after AddListRow()");
			assert.equal(newRow.GetIndex(), 4, "Appended row index is 4");
			assert.equal(newRow.GetRange().GetAddress(true, true), "$B$6:$D$6", "Appended row range is row 6");

			// Insert at position 2
			var insertedRow = tbl.AddListRow(2);
			assert.equal(tbl.GetListRows().length, 5, "Count is 5 after AddListRow(2)");
			assert.equal(insertedRow.GetIndex(), 2, "Inserted row index is 2");
		});

		QUnit.test("ListRow - Delete", function (assert)
		{
			initializeTest();

			var tbl = ws.AddListObject("xlSrcRange", "B2:D5");
			ws.GetRange("B3").SetValue("Alice");
			ws.GetRange("B4").SetValue("Bob");
			ws.GetRange("B5").SetValue("Carol");

			assert.equal(tbl.GetListRows().length, 3, "Initial count is 3");

			tbl.GetListRows()[1].Delete();
			assert.equal(tbl.GetListRows().length, 2, "Count is 2 after deleting middle row");
			assert.equal(ws.GetRange("B3").GetValue(), "Alice", "First row still has Alice");
			assert.equal(ws.GetRange("B4").GetValue(), "Carol", "Second row now has Carol after Bob was deleted");
		});

		QUnit.test("GetSort - defaults", function (assert)
		{
			initializeTest();

			var tbl = ws.AddListObject("xlSrcRange", "B2:D5");
			var sort = tbl.GetSort();

			assert.ok(sort !== null, "GetSort returns an object");
			assert.equal(sort.GetHeader(), "xlYes", "Header is always xlYes");
			assert.equal(sort.GetMatchCase(), false, "Default MatchCase is false");
			assert.equal(sort.GetOrientation(), "xlTopToBottom", "Default orientation is xlTopToBottom");
			assert.equal(sort.GetSortMethod(), "xlPinYin", "Default sort method is xlPinYin");
			assert.ok(tbl.Sort !== null, "Sort property works");

			var rng = sort.GetRng();
			assert.ok(rng !== null, "GetRng returns a range");
			assert.equal(rng.GetAddress(true, true), "$B$3:$D$5", "Rng is the data body range");
		});

		QUnit.test("Sort - MatchCase/Orientation/SortMethod", function (assert)
		{
			initializeTest();

			var tbl  = ws.AddListObject("xlSrcRange", "B2:D5");
			var sort = tbl.GetSort();

			sort.SetMatchCase(true);
			assert.equal(sort.GetMatchCase(), true, "MatchCase set to true");

			sort.SetOrientation("xlLeftToRight");
			assert.equal(sort.GetOrientation(), "xlLeftToRight", "Orientation set to xlLeftToRight");

			sort.SetSortMethod("xlStroke");
			assert.equal(sort.GetSortMethod(), "xlStroke", "SortMethod set to xlStroke");

			sort.MatchCase = false;
			assert.equal(sort.MatchCase, false, "MatchCase property setter works");

			sort.Orientation = "xlTopToBottom";
			assert.equal(sort.Orientation, "xlTopToBottom", "Orientation property setter works");
		});

		QUnit.test("SortFields - Add/Add2/Count/Item/Clear", function (assert)
		{
			initializeTest();

			var tbl    = ws.AddListObject("xlSrcRange", "B2:D5");
			var sort   = tbl.GetSort();
			var fields = sort.GetSortFields();

			assert.equal(fields.GetCount(), 0, "Initially no sort fields");
			assert.equal(fields.Count, 0, "Count property works");

			var sf = fields.Add(ws.GetRange("B2"), "xlSortOnValues", "xlAscending");
			assert.ok(sf !== null, "Add returns a SortField");
			assert.equal(fields.Count, 1, "Count is 1 after Add");

			fields.Add(ws.GetRange("C2"), "xlSortOnValues", "xlDescending");
			assert.equal(fields.Count, 2, "Count is 2 after second Add");

			assert.strictEqual(fields.Add(null), null, "Add with non-ApiRange key returns null");

			var sf2 = fields.Add2(ws.GetRange("D2"), "xlSortOnValues", "xlAscending", null, "xlSortNormal", "Population");
			assert.ok(sf2 !== null, "Add2 returns a SortField");
			assert.equal(fields.Count, 3, "Count is 3 after Add2");

			var item = fields.Item(1);
			assert.ok(item !== null, "Item(1) returns a SortField");
			assert.equal(item.GetSortOn(), "xlSortOnValues", "SortField.GetSortOn works");
			assert.equal(item.GetOrder(), "xlAscending", "SortField.GetOrder works for ascending");
			assert.equal(fields.Item(2).GetOrder(), "xlDescending", "SortField.GetOrder works for descending");
			assert.equal(item.GetPriority(), 1, "SortField.GetPriority returns 1 for first field");
			assert.equal(fields.Item(2).GetPriority(), 2, "SortField.GetPriority returns 2 for second field");

			fields.Clear();
			assert.equal(fields.Count, 0, "Count is 0 after Clear");
		});

		QUnit.test("SortField - SortOn/Order/Priority/CustomOrder/DataOption writable", function (assert)
		{
			initializeTest();

			var tbl    = ws.AddListObject("xlSrcRange", "B2:D5");
			var fields = tbl.GetSort().SortFields;

			var sf = fields.Add(ws.GetRange("B2"), "xlSortOnValues", "xlAscending");

			sf.SetSortOn("xlSortOnCellColor");
			assert.equal(sf.GetSortOn(), "xlSortOnCellColor", "SetSortOn works");
			sf.SortOn = "xlSortOnValues";
			assert.equal(sf.SortOn, "xlSortOnValues", "SortOn property setter works");

			sf.SetOrder("xlDescending");
			assert.equal(sf.GetOrder(), "xlDescending", "SetOrder works");
			sf.Order = "xlAscending";
			assert.equal(sf.Order, "xlAscending", "Order property setter works");

			sf.SetCustomOrder("myList");
			assert.equal(sf.GetCustomOrder(), "myList", "SetCustomOrder works");
			sf.CustomOrder = null;
			assert.strictEqual(sf.CustomOrder, null, "CustomOrder property setter works");

			sf.SetDataOption("xlSortTextAsNumbers");
			assert.equal(sf.GetDataOption(), "xlSortTextAsNumbers", "SetDataOption works");
			sf.DataOption = "xlSortNormal";
			assert.equal(sf.DataOption, "xlSortNormal", "DataOption property setter works");
		});

		QUnit.test("SortField - Priority writable / SetPriority reorders", function (assert)
		{
			initializeTest();

			var tbl    = ws.AddListObject("xlSrcRange", "B2:D5");
			var fields = tbl.GetSort().SortFields;

			fields.Add(ws.GetRange("B2"), "xlSortOnValues", "xlAscending");
			fields.Add(ws.GetRange("C2"), "xlSortOnValues", "xlAscending");
			fields.Add(ws.GetRange("D2"), "xlSortOnValues", "xlAscending");

			// Move last field to priority 1
			var sf = fields.Item(3);
			sf.SetPriority(1);
			assert.equal(fields.Item(1).GetKey().GetAddress(true, true), "$D$2:$D$5", "After SetPriority(1): D is first");
			assert.equal(fields.Item(2).GetKey().GetAddress(true, true), "$B$2:$B$5", "After SetPriority(1): B is second");
			assert.equal(fields.Item(3).GetKey().GetAddress(true, true), "$C$2:$C$5", "After SetPriority(1): C is third");
			assert.equal(sf.GetPriority(), 1, "GetPriority reflects new position");

			sf.Priority = 3;
			assert.equal(fields.Item(3).GetKey().GetAddress(true, true), "$D$2:$D$5", "After Priority=3: D is back at end");
		});

		QUnit.test("SortField - SortOnValue/SetIcon", function (assert)
		{
			initializeTest();

			var tbl    = ws.AddListObject("xlSrcRange", "B2:D5");
			var fields = tbl.GetSort().SortFields;

			var sf = fields.Add(ws.GetRange("B2"), "xlSortOnValues", "xlAscending");
			assert.strictEqual(sf.GetSortOnValue(), null, "SortOnValue is null for value sort");
			assert.strictEqual(sf.SortOnValue, null, "SortOnValue property works");

			sf.SetIcon("someIcon");
			assert.equal(sf.GetSortOn(), "xlSortOnIcon", "SetIcon sets sortOn to xlSortOnIcon");
			assert.equal(sf.GetSortOnValue(), "someIcon", "SetIcon stores the icon in SortOnValue");
		});

		QUnit.test("SortField - GetKey/ModifyKey/Delete", function (assert)
		{
			initializeTest();

			var tbl    = ws.AddListObject("xlSrcRange", "B2:D5");
			var sort   = tbl.GetSort();
			var fields = sort.SortFields;

			fields.Add(ws.GetRange("B2"), "xlSortOnValues", "xlAscending");
			fields.Add(ws.GetRange("C2"), "xlSortOnValues", "xlAscending");

			var sf = fields.Item(1);
			var keyRng = sf.GetKey();
			assert.ok(keyRng !== null, "GetKey returns a range");
			assert.equal(keyRng.GetAddress(true, true), "$B$2:$B$5", "Key range spans full table column B");

			sf.ModifyKey(ws.GetRange("D2"));
			assert.equal(sf.GetKey().GetAddress(true, true), "$D$2:$D$5", "ModifyKey changed key to column D");

			fields.Item(1).Delete();
			assert.equal(fields.Count, 1, "Count is 1 after Delete of first field");
			assert.equal(fields.Item(1).GetKey().GetAddress(true, true), "$C$2:$C$5", "Remaining field is column C");
		});

		QUnit.test("Sort - Apply sorts table data", function (assert)
		{
			initializeTest();

			// Create table with header in B2, data in B3:D5
			var tbl = ws.AddListObject("xlSrcRange", "B2:D5");
			ws.GetRange("B2").SetValue("Name");
			ws.GetRange("B3").SetValue("Charlie");
			ws.GetRange("B4").SetValue("Alice");
			ws.GetRange("B5").SetValue("Bob");

			var sort   = tbl.GetSort();
			var fields = sort.GetSortFields();
			fields.Clear();
			fields.Add(ws.GetRange("B2"), "xlSortOnValues", "xlAscending");
			sort.Apply();

			assert.equal(ws.GetRange("B3").GetValue(), "Alice",   "After ascending sort: row 1 = Alice");
			assert.equal(ws.GetRange("B4").GetValue(), "Bob",     "After ascending sort: row 2 = Bob");
			assert.equal(ws.GetRange("B5").GetValue(), "Charlie", "After ascending sort: row 3 = Charlie");

			// Sort descending
			var sort2   = tbl.GetSort();
			var fields2 = sort2.GetSortFields();
			fields2.Clear();
			fields2.Add(ws.GetRange("B2"), "xlSortOnValues", "xlDescending");
			sort2.Apply();

			assert.equal(ws.GetRange("B3").GetValue(), "Charlie", "After descending sort: row 1 = Charlie");
			assert.equal(ws.GetRange("B4").GetValue(), "Bob",     "After descending sort: row 2 = Bob");
			assert.equal(ws.GetRange("B5").GetValue(), "Alice",   "After descending sort: row 3 = Alice");

			// Sort state is persisted — a fresh GetSort() reflects last applied
			var sort3 = tbl.GetSort();
			assert.equal(sort3.GetSortFields().Count, 1, "Persisted sort state has 1 field");
			assert.equal(sort3.GetSortFields().Item(1).GetOrder(), "xlDescending", "Persisted sort order is xlDescending");
		});
	});
});
