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

$(function () {
	
	var ws = AscTest.JsApi.GetActiveSheet();

    theRange = function (address) {
        return ws.GetRange(address);
    };
	
	function initializeTest(){}

    // ====== TESTS ======

    QUnit.module("ApiAutoFilter", function () {
        QUnit.test("AutoFilter filters length after Clear", function (assert) {
            initializeTest();
            let range = ws.GetRange("A1");
            range.SetAutoFilter();
            let filters = ws.AutoFilter.Filters;
            // length of filters should be 0 after Clear()
            assert.equal(filters.length, 0, "Filters length is 0");
        });

        QUnit.test("ApiAutoFilter check properties", function (assert) {
            initializeTest();

            theRange("A1").SetValue(10);
            theRange("A2").SetValue(20);
            theRange("A3").SetValue(2);
            theRange("A4").SetValue(5);
            theRange("A5").SetValue(4);
            theRange("A6").SetValue(7);

            let range = ws.GetRange("A1:A10");
            range.SetAutoFilter(1, [2, 5], "xlFilterValues");

            let filters = ws.AutoFilter.Filters;
            // length of filters should be 1 after setting filter on column 1

            assert.equal(filters.length, 1, "Filters length is 1");
            assert.equal(filters[0].Criteria1, "=2", "Criteria1 is 2");
            assert.equal(filters[0].Operator, "xlOr", "Operator is xlFilterValues");
            assert.equal(filters[0].Criteria2, "=5", "Criteria2 is 5");
            assert.equal(filters[0].On, true, "On is true");
        });

        QUnit.test("Remove AutoFilter when Field is null", function (assert) {
            initializeTest();

            let range = ws.GetRange("A1:B5");
            range.SetAutoFilter(1, [2,5], "xlFilterValues");
            assert.equal(ws.AutoFilter.FilterMode, true, "FilterMode true after add");

            range.SetAutoFilter(null); // remove whole AutoFilter
            assert.equal(ws.AutoFilter.FilterMode, false, "FilterMode false after remove");
            assert.equal(ws.AutoFilter.Range, null, "Range null after remove");
            assert.equal(ws.AutoFilter.Filters.length, 0, "Filters empty after remove");
        });

        QUnit.test("xlFilterValues with 3+ items keeps xlFilterValues", function (assert) {
            initializeTest();

            ["2", "5", "7", "9"].forEach((v, i) => theRange("A" + (i + 2)).SetValue(v));
            let range = ws.GetRange("A1:A10");
            range.SetAutoFilter(1, ["2", "5", "7"], "xlFilterValues");

            let f = ws.AutoFilter.Filters;
            assert.equal(f.length, 1, "One filter created");
            assert.deepEqual(f[0].Criteria1.sort(), ["2", "5", "7"], "Criteria1 returns visible values array");
            assert.equal(f[0].Criteria2, null, "Criteria2 is null for values filter");
            assert.equal(f[0].Operator, "xlFilterValues", "Operator is xlFilterValues");
            assert.equal(f[0].On, true, "On true");
        });

        QUnit.test("xlFilterValues with 2 items converts to xlOr custom filter", function (assert) {
            initializeTest();

            [10, 20, 2, 5, 4, 7].forEach((v, i) => theRange("A" + (i + 1)).SetValue(v));
            let range = ws.GetRange("A1:A10");
            range.SetAutoFilter(1, [2, 5], "xlFilterValues");

            let f = ws.AutoFilter.Filters;
            assert.equal(f.length, 1, "One filter created");
            assert.equal(f[0].Operator, "xlOr", "Operator converted to xlOr");
            assert.equal(f[0].Criteria1, "=2", "Criteria1 is '=2'");
            assert.equal(f[0].Criteria2, "=5", "Criteria2 is '=5'");
            assert.equal(f[0].On, true, "On true");
        });

        QUnit.test("Custom filter xlOr preserves signs and criteria", function (assert) {
            initializeTest();

            [1, 2, 3, 4, 5, 6, 7, 8, 9].forEach((v, i) => theRange("A" + (i + 1)).SetValue(v));
            let range = ws.GetRange("A1:A10");
            range.SetAutoFilter(1, ">3", "xlOr", "<=5");

            let f = ws.AutoFilter.Filters;
            assert.equal(f.length, 1, "One filter created");
            assert.equal(f[0].Operator, "xlOr");
            assert.equal(f[0].Criteria1, ">3");
            assert.equal(f[0].Criteria2, "<=5");
            assert.equal(f[0].On, true, "Filter is on, i.e. applied");
        });

        QUnit.test("Custom filter xlAnd", function (assert) {
            initializeTest();

            [1, 2, 3, 4, 5, 6, 7, 8, 9].forEach((v, i) => theRange("A" + (i + 1)).SetValue(v));
            let range = ws.GetRange("A1:A10");
            range.SetAutoFilter(1, ">=2", "xlAnd", "<8");

            let f = ws.AutoFilter.Filters;
            assert.equal(f.length, 1, "One filter created");
            assert.equal(f[0].Operator, "xlAnd");
            assert.equal(f[0].Criteria1, ">=2");
            assert.equal(f[0].Criteria2, "<8");
            assert.equal(f[0].On, true, "Filter is on, i.e. applied");
        });

        QUnit.test("Top10 items filter", function (assert) {
            initializeTest();

            [10, 20, 30, 40, 50, 60].forEach((v, i) => theRange("A" + (i + 1)).SetValue(v));
            let range = ws.GetRange("A1:A10");
            range.SetAutoFilter(1, "3", "xlTop10Items");

            let f = ws.AutoFilter.Filters;
            assert.equal(f.length, 1, "One filter created");
            assert.equal(f[0].Operator, "xlTop10Items");
            assert.equal(f[0].Criteria1, 3, "Criteria1 returns numeric Val");
            assert.equal(f[0].Criteria2, null, "Criteria2 should be null");
            assert.equal(f[0].On, true, "Filter is on");
        });

        QUnit.test("Bottom10 percent filter", function (assert) {
            initializeTest();

            [10, 20, 30, 40, 50, 60].forEach((v, i) => theRange("A" + (i + 1)).SetValue(v));
            let range = ws.GetRange("A1:A10");
            range.SetAutoFilter(1, "50", "xlBottom10Percent");

            let f = ws.AutoFilter.Filters;
            assert.equal(f.length, 1, "One filter created");
            assert.equal(f[0].Operator, "xlBottom10Percent");
            assert.equal(f[0].Criteria1, 50, "Criteria1 returns numeric Val");
            assert.equal(f[0].On, true, "Filter is on");
        });

        QUnit.test("Dynamic filter AboveAverage", function (assert) {
            initializeTest();

            [1, 2, 3, 100].forEach((v, i) => theRange("A" + (i + 1)).SetValue(v));
            let range = ws.GetRange("A1:A10");
            range.SetAutoFilter(1, "xlFilterAboveAverage", "xlFilterDynamic");

            let f = ws.AutoFilter.Filters;
            assert.equal(f.length, 1, "One filter created");
            assert.equal(f[0].Operator, "xlFilterDynamic");
            assert.equal(f[0].Criteria1, "xlFilterAboveAverage");
            assert.equal(f[0].Criteria2, null, "Criteria2 should be null");
            assert.equal(f[0].On, true, "Filter is on");
        });

        QUnit.test("Color filter CellColor (Criteria1 null by API)", function (assert) {
            initializeTest();

            // Make some colored cells to be realistic
            theRange("A1").SetValue("x");
            theRange("A2").SetValue("y");
            theRange("A1").SetFillColor(AscTest.JsApi.CreateColorFromRGB(255, 255, 0));
            theRange("A2").SetFillColor(AscTest.JsApi.CreateColorFromRGB(0, 255, 0));

            let range = ws.GetRange("A1:A5");
            range.SetAutoFilter(1, AscTest.JsApi.CreateColorFromRGB(255, 255, 0), "xlFilterCellColor");

            let f = ws.AutoFilter.Filters;
            assert.equal(f.length, 1, "One filter created");
            assert.equal(f[0].Operator, "xlFilterCellColor");
            assert.equal(f[0].Criteria1, null, "Criteria1 is null for color filters");
            assert.equal(f[0].Criteria2, null, "Criteria2 is null for color filters");
            assert.equal(f[0].On, true, "Filter is on");
        });

        QUnit.test("Clear specific column filter when Criteria1 is null", function (assert) {
            initializeTest();

            [10, 20, 2, 5, 4, 7].forEach((v, i) => theRange("A" + (i + 1)).SetValue(v));
            let range = ws.GetRange("A1:A10");

            range.SetAutoFilter(1, ">3", "xlOr", "<=7");
            assert.equal(ws.AutoFilter.Filters.length, 1, "Filter exists before clear");

            range.SetAutoFilter(1, null); // clear column filter
            assert.equal(ws.AutoFilter.FilterMode, true, "AutoFilter still exists");
            assert.equal(ws.AutoFilter.Filters.length, 1, "Filters empty after column clear");
        });

        QUnit.test("Invalid Field does not add AutoFilter", function (assert) {
            initializeTest();

            let range = ws.GetRange("A1:B5");
            try {
                range.SetAutoFilter("foo", ">1", "xlOr"); // invalid Field
            } catch (e) {
                // Expected error for invalid Field
            }

            assert.equal(ws.AutoFilter.FilterMode, false, "AutoFilter not added");
            assert.equal(ws.AutoFilter.Range, null, "Range null");
            assert.equal(ws.AutoFilter.Filters.length, 0, "No filters created");
        });

        // ---- ListObject AutoFilter tests (rows 6-29 to avoid interference) ----

        QUnit.test("ListObject - values filter hides rows", function (assert) {
            initializeTest();

            theRange("A6").SetValue("Product");
            theRange("A7").SetValue("Apples");
            theRange("A8").SetValue("Oranges");
            theRange("A9").SetValue("Bananas");
            ws.AddListObject("xlSrcRange", "A6:A9");

            ws.GetRange("A6:A9").SetAutoFilter(1, ["Apples", "Bananas"], "xlFilterValues");

            assert.equal(ws.worksheet.getRowHidden(6), false, "Row 7 (Apples) visible");
            assert.equal(ws.worksheet.getRowHidden(7), true,  "Row 8 (Oranges) hidden");
            assert.equal(ws.worksheet.getRowHidden(8), false, "Row 9 (Bananas) visible");
        });

        QUnit.test("ListObject - custom filter (>) hides rows", function (assert) {
            initializeTest();

            theRange("C6").SetValue("Price");
            theRange("C7").SetValue(100);
            theRange("C8").SetValue(200);
            theRange("C9").SetValue(150);
            ws.AddListObject("xlSrcRange", "C6:C9");

            ws.GetRange("C6:C9").SetAutoFilter(1, ">100", "xlOr");

            assert.equal(ws.worksheet.getRowHidden(6), true,  "Row 7 (100) hidden");
            assert.equal(ws.worksheet.getRowHidden(7), false, "Row 8 (200) visible");
            assert.equal(ws.worksheet.getRowHidden(8), false, "Row 9 (150) visible");
        });

        QUnit.test("ListObject - filter on field 2 of two-column table", function (assert) {
            initializeTest();

            theRange("A11").SetValue("Product"); theRange("B11").SetValue("Price");
            theRange("A12").SetValue("Apples");  theRange("B12").SetValue(100);
            theRange("A13").SetValue("Oranges"); theRange("B13").SetValue(200);
            theRange("A14").SetValue("Bananas"); theRange("B14").SetValue(100);
            ws.AddListObject("xlSrcRange", "A11:B14");

            ws.GetRange("A11:B14").SetAutoFilter(2, ">100", "xlOr");

            assert.equal(ws.worksheet.getRowHidden(11), true,  "Row 12 (Price=100) hidden");
            assert.equal(ws.worksheet.getRowHidden(12), false, "Row 13 (Price=200) visible");
            assert.equal(ws.worksheet.getRowHidden(13), true,  "Row 14 (Price=100) hidden");
        });

        QUnit.test("ListObject - filter does not create worksheet-level AutoFilter", function (assert) {
            initializeTest();

            theRange("D11").SetValue("Name");
            theRange("D12").SetValue("Alice");
            theRange("D13").SetValue("Bob");
            theRange("D14").SetValue("Alice");
            ws.AddListObject("xlSrcRange", "D11:D14");

            var afBefore = ws.AutoFilter.FilterMode;
            ws.GetRange("D11:D14").SetAutoFilter(1, ["Alice"], "xlFilterValues");

            assert.equal(afBefore, false, "ws.AutoFilter was not set before");
            assert.equal(ws.AutoFilter.FilterMode, false, "ws.AutoFilter still not set after table filter");
        });

        QUnit.test("ListObject - two filters on different columns are independent", function (assert) {
            initializeTest();

            theRange("A16").SetValue("Name");  theRange("B16").SetValue("Score");
            theRange("A17").SetValue("Alice");  theRange("B17").SetValue(80);
            theRange("A18").SetValue("Bob");    theRange("B18").SetValue(90);
            theRange("A19").SetValue("Alice");  theRange("B19").SetValue(40);
            ws.AddListObject("xlSrcRange", "A16:B19");

            // Filter field 1: only Alice
            ws.GetRange("A16:B19").SetAutoFilter(1, ["Alice"], "xlFilterValues");
            // Filter field 2: Score > 50  (second call must not wipe field-1 filter)
            ws.GetRange("A16:B19").SetAutoFilter(2, ">50", "xlOr");

            // Row 17: Alice / 80 — passes both filters → visible
            assert.equal(ws.worksheet.getRowHidden(16), false, "Alice/80 visible (passes both filters)");
            // Row 18: Bob / 90 — fails field-1 filter → hidden
            assert.equal(ws.worksheet.getRowHidden(17), true,  "Bob/90 hidden (fails name filter)");
            // Row 19: Alice / 40 — fails field-2 filter → hidden
            assert.equal(ws.worksheet.getRowHidden(18), true,  "Alice/40 hidden (fails score filter)");
        });

        QUnit.test("ListObject - SetAutoFilter() removes table AutoFilter and shows all rows", function (assert) {
            initializeTest();

            theRange("D16").SetValue("City");
            theRange("D17").SetValue("Paris");
            theRange("D18").SetValue("Rome");
            theRange("D19").SetValue("Paris");
            ws.AddListObject("xlSrcRange", "D16:D19");

            ws.GetRange("D16:D19").SetAutoFilter(1, ["Paris"], "xlFilterValues");
            assert.equal(ws.worksheet.getRowHidden(17), true, "Rome hidden after filter");

            // SetAutoFilter() with no args removes the table's AutoFilter and shows all rows
            ws.GetRange("D16:D19").SetAutoFilter();

            assert.equal(ws.worksheet.getRowHidden(17), false, "Rome visible after SetAutoFilter()");
        });

        QUnit.test("ListObject - SetAutoFilter(field, null) clears column filter", function (assert) {
            initializeTest();

            theRange("F16").SetValue("Item");
            theRange("F17").SetValue("X");
            theRange("F18").SetValue("Y");
            theRange("F19").SetValue("Z");
            ws.AddListObject("xlSrcRange", "F16:F19");

            ws.GetRange("F16:F19").SetAutoFilter(1, ["X"], "xlFilterValues");
            assert.equal(ws.worksheet.getRowHidden(17), true, "Y hidden after filter");
            assert.equal(ws.worksheet.getRowHidden(18), true, "Z hidden after filter");

            ws.GetRange("F16:F19").SetAutoFilter(null, null);

            assert.equal(ws.worksheet.getRowHidden(17), false, "Y visible after column clear");
            assert.equal(ws.worksheet.getRowHidden(18), false, "Z visible after column clear");
        });

        QUnit.test("ListObject - AND filter hides rows outside range", function (assert) {
            initializeTest();

            theRange("H16").SetValue("Score");
            theRange("H17").SetValue(30);
            theRange("H18").SetValue(70);
            theRange("H19").SetValue(90);
            ws.AddListObject("xlSrcRange", "H16:H19");

            ws.GetRange("H16:H19").SetAutoFilter(1, ">=50", "xlAnd", "<=80");

            assert.equal(ws.worksheet.getRowHidden(16), true,  "30 hidden (< 50)");
            assert.equal(ws.worksheet.getRowHidden(17), false, "70 visible (50-80)");
            assert.equal(ws.worksheet.getRowHidden(18), true,  "90 hidden (> 80)");
        });

        QUnit.test("Field out of range does not add AutoFilter", function (assert) {
            initializeTest();

            let range = ws.GetRange("A1:B5"); // only 2 columns
            try {
                range.SetAutoFilter(3, ">1", "xlOr"); // invalid: field > columns count
            } catch (e) {
                // Expected error for Field out of range
            }

            assert.equal(ws.AutoFilter.FilterMode, false, "AutoFilter not added");
            assert.equal(ws.AutoFilter.Filters.length, 0, "No filters created");
        });

    });

    QUnit.test("ApplyFilter recalculates visibility on new data", function (assert) {
        initializeTest();

        // Values: >2 are visible, <=2 are hidden after filter
        [1, 2, 2, 4, 5, "", ""].forEach((v, i) => theRange("A" + (i + 1)).SetValue(v));
        let range = ws.GetRange("A1:A10");
        range.SetAutoFilter(1, ">2", "xlOr");

        assert.equal(ws.AutoFilter.FilterMode, true, "FilterMode is true after SetAutoFilter");

        // Header row is row 0 in ws.worksheet; data below
        assert.equal(ws.worksheet.getRowHidden(0), false, "Row 1 visible (1)");
        assert.equal(ws.worksheet.getRowHidden(1), true,  "Row 2 hidden (2)");
        assert.equal(ws.worksheet.getRowHidden(2), true,  "Row 3 hidden (2)");
        assert.equal(ws.worksheet.getRowHidden(3), false, "Row 4 visible (4)");
        assert.equal(ws.worksheet.getRowHidden(4), false, "Row 5 visible (5)");

        // Change data after filter has been applied
        [1, 2, 2, 4, 5, 2, 3].forEach((v, i) => theRange("A" + (i + 1)).SetValue(v));

        ws.worksheet.setRowHidden(false, 5, 6);

        // Still using old visibility until ApplyFilter is called again
        assert.equal(ws.worksheet.getRowHidden(5), false, "Row 7 visible before ApplyFilter");
        assert.equal(ws.worksheet.getRowHidden(6), false, "Row 8 visible before ApplyFilter");

        // Reapply filter – should reevaluate all rows
        ws.AutoFilter.ApplyFilter();

        assert.equal(ws.worksheet.getRowHidden(0), false, "Row 1 visible (1)");
        assert.equal(ws.worksheet.getRowHidden(1), true,  "Row 2 hidden (2)");
        assert.equal(ws.worksheet.getRowHidden(2), true,  "Row 3 hidden (2)");
        assert.equal(ws.worksheet.getRowHidden(3), false,  "Row 4 visible (4)");
        assert.equal(ws.worksheet.getRowHidden(4), false, "Row 5 visible (5)");
        assert.equal(ws.worksheet.getRowHidden(5), true, "Row 6 hidden (2)");
        assert.equal(ws.worksheet.getRowHidden(6), false,  "Row 7 visible (3)");
    });

    QUnit.test("ShowAllData makes all rows visible but keeps AutoFilter", function (assert) {
        initializeTest();

        // Values: only >2 should be visible before ShowAllData
        [1, 2, 3, 4, 5].forEach((v, i) => theRange("A" + (i + 1)).SetValue(v));
        let range = ws.GetRange("A1:A10");
        range.SetAutoFilter(1, "<1", "xlOr");

        // Sanity check: some rows are hidden
        assert.equal(ws.worksheet.getRowHidden(0), false, "Row 1 visible (header)");
        assert.equal(ws.worksheet.getRowHidden(1), true, "Row 2 hidden (2)");
        assert.equal(ws.worksheet.getRowHidden(2), true, "Row 3 hidden (3)");
        assert.equal(ws.worksheet.getRowHidden(3), true, "Row 4 hidden (4)");
        assert.equal(ws.worksheet.getRowHidden(4), true, "Row 5 hidden (5)");

        ws.AutoFilter.ShowAllData();

        // All rows of the AutoFilter range must be visible now
        assert.equal(ws.worksheet.getRowHidden(0), false, "Row 1 visible (header)");
        assert.equal(ws.worksheet.getRowHidden(1), false, "Row 2 visible (2)");
        assert.equal(ws.worksheet.getRowHidden(2), false, "Row 3 visible (3)");
        assert.equal(ws.worksheet.getRowHidden(3), false, "Row 4 visible (4)");
        assert.equal(ws.worksheet.getRowHidden(4), false, "Row 5 visible (5)");

        // AutoFilter object should still exist (drop-downs preserved)
        assert.equal(ws.AutoFilter.FilterMode, true, "AutoFilter still present after ShowAllData");
    });

});
