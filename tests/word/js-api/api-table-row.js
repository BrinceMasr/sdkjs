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
