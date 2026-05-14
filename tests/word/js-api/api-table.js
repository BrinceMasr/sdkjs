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
	QUnit.module('ApiTable');

	QUnit.test('GetRow', function (assert)
	{
		const table = AscTest.JsApi.CreateTable(3, 4);

		assert.strictEqual(table.GetRow(0).GetClassType(), 'tableRow', 'GetRow(0) returns the first ApiTableRow');
		assert.strictEqual(table.GetRow(1).GetClassType(), 'tableRow', 'GetRow(1) returns the second ApiTableRow');
		assert.strictEqual(table.GetRow(2).GetClassType(), 'tableRow', 'GetRow(2) returns the last ApiTableRow');

		assert.throws(
			function() { table.GetRow(-1); },
			'GetRow throws on negative index'
		);
		assert.throws(
			function() { table.GetRow(3); },
			'GetRow throws on out-of-bounds index'
		);
	});

	QUnit.test('Rows', function (assert)
	{
		const table = AscTest.JsApi.CreateTable(3, 4);
		const rows = table.Rows;

		assert.strictEqual(rows.length, 3, 'Rows returns an array with the correct number of rows');
		assert.strictEqual(rows[0].GetClassType(), 'tableRow', 'First element is an ApiTableRow');
		assert.strictEqual(rows[1].GetClassType(), 'tableRow', 'Second element is an ApiTableRow');
		assert.strictEqual(rows[2].GetClassType(), 'tableRow', 'Third element is an ApiTableRow');

		const singleRowTable = AscTest.JsApi.CreateTable(1, 2);
		assert.strictEqual(singleRowTable.Rows.length, 1, 'Rows returns a single-element array for a one-row table');
	});

	QUnit.test('GetCell', function (assert)
	{
		const table = AscTest.JsApi.CreateTable(3, 4);

		assert.strictEqual(table.GetCell(0, 0).GetClassType(), 'tableCell', 'GetCell(0, 0) returns the first ApiTableCell');
		assert.strictEqual(table.GetCell(2, 3).GetClassType(), 'tableCell', 'GetCell(2, 3) returns the last ApiTableCell');
		assert.strictEqual(table.GetCell(1, 2).GetClassType(), 'tableCell', 'GetCell(1, 2) returns a middle cell');

		assert.throws(
			function() { table.GetCell(-1, 0); },
			'GetCell throws on negative row index'
		);
		assert.throws(
			function() { table.GetCell(3, 0); },
			'GetCell throws on out-of-bounds row index'
		);
		assert.throws(
			function() { table.GetCell(0, 4); },
			'GetCell throws on out-of-bounds cell index'
		);
	});

	QUnit.test('Cells', function (assert)
	{
		const table = AscTest.JsApi.CreateTable(3, 4);
		const cells = table.Cells;

		assert.strictEqual(cells.length, 3, 'Cells outer array length matches row count');
		assert.strictEqual(cells[0].length, 4, 'Cells inner array length matches column count');
		assert.strictEqual(cells[0][0].GetClassType(), 'tableCell', 'cells[0][0] is an ApiTableCell');
		assert.strictEqual(cells[2][3].GetClassType(), 'tableCell', 'cells[2][3] (last cell) is an ApiTableCell');
		assert.strictEqual(cells[1][2].GetClassType(), 'tableCell', 'cells[1][2] (middle cell) is an ApiTableCell');
	});
});
