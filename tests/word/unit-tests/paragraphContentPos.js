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

$(function () {

	QUnit.module("Unit-tests for AscWord.CParagraphContentPos");


	QUnit.test("Test:", function (assert)
	{
		let oPos = new AscWord.CParagraphContentPos();
		assert.strictEqual(oPos.GetDepth(), -1, "Create new pos and check the depth (must be negative)");

		oPos.Add(4);
		oPos.Add(8);
		oPos.Add(15);
		oPos.Add(16);
		oPos.Add(23);
		oPos.Add(42);

		assert.strictEqual(oPos.GetDepth(), 5, "Fill the pos and check the depth");

		assert.strictEqual(oPos.Get(0), 4, "[0]");
		assert.strictEqual(oPos.Get(1), 8, "[1]");
		assert.strictEqual(oPos.Get(2), 15, "[2]");
		assert.strictEqual(oPos.Get(3), 16, "[3]");
		assert.strictEqual(oPos.Get(4), 23, "[4]");
		assert.strictEqual(oPos.Get(5), 42, "[5]");

		let oPos2 = oPos.Copy();

		assert.strictEqual(oPos2.Compare(oPos), 0, "Make a copy and check for equality");

		oPos2.Update2(20, 3);
		assert.strictEqual(oPos2.Get(3), 20, "Check [3] after update");
		assert.strictEqual(oPos2.GetDepth(), 5, "Check depth after Update2");
		assert.strictEqual(oPos2.Compare(oPos), 1, "Compare pos +1");

		oPos2.Update2(2, 3);
		assert.strictEqual(oPos2.Compare(oPos), -1, "Compare pos -1");

		oPos2.DecreaseDepth(3);
		assert.strictEqual(oPos2.GetDepth(), 2, "Check depth after decreasing by 3");
		assert.strictEqual(oPos2.Compare(oPos), -1, "Compare decreased pos and original pos");
		assert.strictEqual(oPos2.IsPartOf(oPos), true, "Check decreased pos as part of original");

		oPos2.Update2(2, 2);
		assert.strictEqual(oPos2.IsPartOf(oPos), false, "Spoil decreased pos and check it as part of original");

		let oPos3 = oPos.Copy();
		oPos3.DecreaseDepth(1);
		assert.strictEqual(oPos3.Compare(oPos), -1, "Make a copy and decrease pos and check for equality");
	});
});
