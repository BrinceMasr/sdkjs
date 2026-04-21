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
});
