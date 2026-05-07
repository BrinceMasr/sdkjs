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
	QUnit.module('ApiTableRow');

	QUnit.test('GetCell', function (assert)
	{
		const table = AscTest.JsApi.CreateTable(2, 3);
		const row = table.GetRow(0);

		assert.strictEqual(row.GetCell(0).GetClassType(), 'tableCell', 'GetCell(0) returns the first ApiTableCell');
		assert.strictEqual(row.GetCell(1).GetClassType(), 'tableCell', 'GetCell(1) returns a middle ApiTableCell');
		assert.strictEqual(row.GetCell(2).GetClassType(), 'tableCell', 'GetCell(2) returns the last ApiTableCell');
	});

	QUnit.test('Cells', function (assert)
	{
		const table = AscTest.JsApi.CreateTable(2, 3);
		const row = table.GetRow(0);
		const cells = row.Cells;

		assert.strictEqual(cells.length, 3, 'Cells length matches column count');
		assert.strictEqual(cells[0].GetClassType(), 'tableCell', 'cells[0] is an ApiTableCell');
		assert.strictEqual(cells[1].GetClassType(), 'tableCell', 'cells[1] is an ApiTableCell');
		assert.strictEqual(cells[2].GetClassType(), 'tableCell', 'cells[2] is an ApiTableCell');

		assert.throws(
			function() { cells[-1]; },
			'Cells throws on negative index'
		);
		assert.throws(
			function() { cells[3]; },
			'Cells throws on out-of-bounds index'
		);
	});
});
