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
	QUnit.module("Test the ApiSignatureForm methods");

	function createApiSignatureForm(pr)
	{
		return AscTest.JsApi.CreateSignatureForm(pr || {});
	}

	QUnit.test("GetClassType", function (assert)
	{
		const form = createApiSignatureForm();
		assert.strictEqual(form.GetClassType(), "signatureForm", "Check class type of a signature form");
	});

	QUnit.test("SetImage, GetImage", function (assert)
	{
		const form = createApiSignatureForm();

		assert.strictEqual(typeof form.GetImage(), "string", "Check GetImage returns a string when no image is set");

		const result = form.SetImage("https://static.onlyoffice.com/assets/docs/samples/img/onlyoffice_logo.png");
		assert.strictEqual(typeof result, "boolean", "Check SetImage returns a boolean");
	});

	QUnit.test("Copy", function (assert)
	{
		const form = createApiSignatureForm({"key": "Signature", "placeholder": "Sign here"});
		const copy = form.Copy();
		assert.strictEqual(copy.GetClassType(), "signatureForm", "Check copy has the same class type");
	});

	QUnit.test("SetPlaceholderText, GetPlaceholderText", function (assert)
	{
		const form = createApiSignatureForm({"placeholder": "Sign here"});
		assert.strictEqual(form.GetPlaceholderText(), "Sign here", "Check placeholder text after creation");

		form.SetPlaceholderText("Updated signature placeholder");
		assert.strictEqual(form.GetPlaceholderText(), "Updated signature placeholder", "Check placeholder text after update");
	});

	QUnit.test("SetFormKey, GetFormKey", function (assert)
	{
		const form = createApiSignatureForm({"key": "Signature"});
		assert.strictEqual(form.GetFormKey(), "Signature", "Check form key after creation");

		form.SetFormKey("UpdatedKey");
		assert.strictEqual(form.GetFormKey(), "UpdatedKey", "Check form key after update");
	});

	QUnit.test("SetBorderColor, GetBorderColor", function (assert)
	{
		const form = createApiSignatureForm();

		assert.strictEqual(form.GetBorderColor(), null, "Check border color for a newly created signature form");

		form.SetBorderColor(255, 122, 100);
		assert.equalRgb(form.GetBorderColor(), { r: 255, g: 122, b: 100 }, "Check border color after setting with RGB components");

		const hexColor = AscTest.JsApi.HexColor("a1b2c3");
		form.SetBorderColor(hexColor);
		assert.equalRgb(form.GetBorderColor(), { r: 161, g: 178, b: 195 }, "Check border color after setting with ApiColor");

		form.SetBorderColor(0, 0, 0, true);
		assert.strictEqual(form.GetBorderColor(), null, "Check border color after resetting it");
	});

	QUnit.test("SetBackgroundColor, GetBackgroundColor", function (assert)
	{
		const form = createApiSignatureForm();

		assert.strictEqual(form.GetBackgroundColor(), null, "Check background color for a newly created signature form");

		form.SetBackgroundColor(255, 122, 100);
		assert.equalRgb(form.GetBackgroundColor(), { r: 255, g: 122, b: 100 }, "Check background color after setting with RGB components");

		const hexColor = AscTest.JsApi.HexColor("a1b2c3");
		form.SetBackgroundColor(hexColor);
		assert.equalRgb(form.GetBackgroundColor().GetRGB(), { r: 161, g: 178, b: 195 }, "Check background color after setting with ApiColor");

		const themeColor = AscTest.JsApi.ThemeColor("accent3");
		form.SetBackgroundColor(themeColor);
		assert.strictEqual(form.GetBackgroundColor().IsThemeColor(), true, "Check background color after setting with theme color");

		form.SetBackgroundColor(0, 0, 0, true);
		assert.strictEqual(form.GetBackgroundColor(), null, "Check background color after resetting it");
	});
});
