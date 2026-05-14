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
	QUnit.module("Test the ApiDateForm methods");

	function createApiDateForm(pr)
	{
		return AscTest.JsApi.CreateDateForm(pr || {"key": "BirthDate", "format": "mm.dd.yyyy", "lang": "en-US"});
	}

	QUnit.test("SetValue, GetValue, Value", function (assert)
	{
		const form = createApiDateForm();

		assert.strictEqual(form.GetValue(), undefined, "Check GetValue returns undefined for a new placeholder date form");

		const testDate = new Date(2024, 0, 15);
		const result = form.SetValue(testDate);
		assert.strictEqual(result, true, "Check SetValue returns true on success");

		const value = form.GetValue();
		assert.strictEqual(form.Value.getTime(), testDate.getTime(), "Check Value getter returns the correct date");
		assert.strictEqual(value.getTime(), testDate.getTime(), "Check GetValue returns the correct date after SetValue");

		const newDate = new Date(2025, 5, 20);
		form.Value = newDate;
		assert.strictEqual(form.GetValue().getTime(), newDate.getTime(), "Check Value setter updates the date");
	});
});
