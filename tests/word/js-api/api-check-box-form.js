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
	QUnit.module("Test the ApiCheckBoxForm methods");

	function createApiCheckBoxForm(pr)
	{
		return AscTest.JsApi.CreateCheckBoxForm(pr || {"key": "Agree"});
	}

	QUnit.test("SetValue, GetValue, Value", function (assert)
	{
		const form = createApiCheckBoxForm();

		assert.strictEqual(form.GetValue(), false, "Check GetValue returns false for a newly created checkbox");

		const result = form.SetValue(true);
		assert.strictEqual(result, true, "Check SetValue returns true on success");
		assert.strictEqual(form.GetValue(), true, "Check GetValue returns true after SetValue(true)");

		assert.strictEqual(form.Value, true, "Check Value getter returns the current state");

		form.Value = false;
		assert.strictEqual(form.GetValue(), false, "Check Value setter updates the checkbox state");
	});

	QUnit.test("SetGroupValue, GetGroupValue, GroupValue", function (assert)
	{
		AscTest.ClearDocument();
		const document = AscTest.JsApi.GetDocument();
		const p = AscTest.JsApi.CreateParagraph();
		document.Push(p);

		const radio1 = createApiCheckBoxForm({"radio": true});
		radio1.SetRadioGroup("Gender");
		radio1.SetChoiceName("Male");
		p.Push(radio1);

		const radio2 = createApiCheckBoxForm({"radio": true});
		radio2.SetRadioGroup("Gender");
		radio2.SetChoiceName("Female");
		p.Push(radio2);

		assert.strictEqual(radio1.GetGroupValue(), "", "Check GetGroupValue returns empty string when nothing is selected");

		radio1.SetGroupValue("Male");
		assert.strictEqual(radio1.GetGroupValue(), "Male", "Check GetGroupValue returns the selected choice name");
		assert.strictEqual(radio1.GroupValue, "Male", "Check GroupValue getter returns the current selection");

		radio1.GroupValue = "Female";
		assert.strictEqual(radio1.GetGroupValue(), "Female", "Check GroupValue setter updates the selection");
	});
});
