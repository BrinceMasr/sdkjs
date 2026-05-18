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
	QUnit.module("Test a correction of bad tables");
	
	QUnit.test("Test: bad vMerge", function (assert)
	{
		let table = AscTest.CreateTable(3, 3);
		
		assert.strictEqual(table.GetRowsCount(), 3, "Create table 3x3 and check number of table rows");
		table.CorrectBadTable();
		assert.strictEqual(table.GetRowsCount(), 3, "Process correction. Still should be 3 rows");
		
		for (let iRow = 0, nRows = table.GetRowsCount(); iRow < nRows; ++iRow)
		{
			let row = table.GetRow(iRow);
			for (let iCell = 0, nCells = row.GetCellsCount(); iCell < nCells; ++iCell)
			{
				let cell = row.GetCell(iCell);
				cell.SetVMerge(vmerge_Continue);
			}
		}
		
		table.CorrectBadTable();
		assert.strictEqual(table.GetRowsCount(), 1, "Set vMerge=continue for all cells and process the correction. There should be one row");
		
		let row = table.GetRow(0);
		if (row)
		{
			for (let iCell = 0, nCells = row.GetCellsCount(); iCell < nCells; ++iCell)
			{
				let cell = row.GetCell(iCell);
				assert.strictEqual(cell.GetVMerge(), vmerge_Restart, "Check vMerge parameter for cell_" + iCell);
			}
		}
	});
});
