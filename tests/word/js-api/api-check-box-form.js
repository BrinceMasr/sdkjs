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
