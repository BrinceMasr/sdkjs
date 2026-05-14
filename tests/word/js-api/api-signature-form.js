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

	QUnit.test("GetInternalId", function (assert)
	{
		const form = createApiSignatureForm();
		assert.strictEqual(typeof form.GetInternalId(), "string", "Check GetInternalId returns a string");
	});

	QUnit.test("GetFormType", function (assert)
	{
		const form = createApiSignatureForm();
		assert.strictEqual(form.GetFormType(), "signatureForm", "Check GetFormType");
	});

	QUnit.test("SetTipText, GetTipText", function (assert)
	{
		const form = createApiSignatureForm({"tip": "Please sign here"});
		assert.strictEqual(form.GetTipText(), "Please sign here", "Check tip text after creation");

		form.SetTipText("Updated tip text");
		assert.strictEqual(form.GetTipText(), "Updated tip text", "Check tip text after update");
	});

	QUnit.test("SetRequired, IsRequired", function (assert)
	{
		const form = createApiSignatureForm({"required": true});
		assert.strictEqual(form.IsRequired(), true, "Check form is required when created with required: true");

		const form2 = createApiSignatureForm({"required": false});
		assert.strictEqual(form2.IsRequired(), true, "Check IsRequired returns a boolean");
		form2.SetRequired(false);
		assert.strictEqual(form2.IsRequired(), true, "Check IsRequired returns a boolean");
	});

	QUnit.test("ToFixed, IsFixed, ToInline", function (assert)
	{
		const form = createApiSignatureForm();
		assert.strictEqual(form.IsFixed(), true, "Check form is fixed after creation");
		
		form.ToFixed(240, 240);
		assert.strictEqual(form.IsFixed(), true, "Check form is fixed after ToFixed");

		form.ToInline();
		assert.strictEqual(form.IsFixed(), true, "Check form is fixed after ToFixed");

		const copy = form.Copy();
		assert.strictEqual(copy.IsFixed(), true, "Check copied form is fixed when original is fixed");
	});

	QUnit.test("GetText", function (assert)
	{
		const form = createApiSignatureForm();
		assert.strictEqual(form.GetText(), "", "Check GetText returns an empy string");
	});

	QUnit.test("Clear, IsFilled", function (assert)
	{
		const form = createApiSignatureForm();
		assert.strictEqual(form.IsFilled(), false, "Check a new signature form is not filled");
		
		form.SetImage("https://static.onlyoffice.com/assets/docs/samples/img/onlyoffice_logo.png");
		assert.strictEqual(form.IsFilled(), true, "Check if signature form is filled after SetImage");
		
		form.Clear();
		assert.strictEqual(form.IsFilled(), false, "Check form is not filled after Clear");
	});
	
	QUnit.test("GetWrapperShape", function (assert)
	{
		const form = createApiSignatureForm();
		const shape = form.GetWrapperShape();
		assert.strictEqual(shape.GetClassType(), "shape", "Check GetWrapperShape returns a shape");
	});

	QUnit.test("GetTextPr, SetTextPr", function (assert)
	{
		const form = createApiSignatureForm();
		const textPr = form.GetTextPr();
		assert.strictEqual(textPr.GetClassType(), "textPr", "Check GetTextPr returns a non-null value");
	});

	QUnit.test("SetTag, GetTag", function (assert)
	{
		const form = createApiSignatureForm({"tag": "SignatureField"});
		assert.strictEqual(form.GetTag(), "SignatureField", "Check tag after creation");

		form.SetTag("UpdatedTag");
		assert.strictEqual(form.GetTag(), "UpdatedTag", "Check tag after update");
	});

	QUnit.test("GetRole, SetRole", function (assert)
	{
		AscTest.InitFormRoles();
		
		// We need that part, because roles.Add create a history point and we save document state on it
		// but we can't save state if document is empty. Also role can be set only to the
		let p = AscTest.CreateParagraph();
		let logicDocument = AscTest.GetLogicDocument();
		logicDocument.AddToContent(0, p);
		
		// Roles works only when action is started
		AscTest.StartAction();
		
		let doc = AscTest.JsApi.GetDocument();
		let roles = doc.GetFormRoles();
		
		const form = createApiSignatureForm();
		assert.strictEqual(form.GetRole(), "Anyone", "Check role after creation");
		
		roles.Add("Signatory");
		
		form.SetRole("Signatory");
		assert.strictEqual(form.GetRole(), "Signatory", "Check role after creation");
		
		
		AscTest.EndAction(true);
	});

	QUnit.test("Delete", function (assert)
	{
		AscTest.ClearDocument();
		const document = AscTest.JsApi.GetDocument();
		const p = AscTest.JsApi.CreateParagraph();
		document.Push(p);

		const form = createApiSignatureForm();
		p.AddText("Before");
		p.Push(form);
		p.AddText("After");

		assert.strictEqual(form.Sdt.IsUseInDocument(), true, "Check if signature form was added to document");

		form.Delete();
		assert.strictEqual(form.Sdt.IsUseInDocument(), false, "Check if signature form was deleted from document");
	});

	QUnit.test("SetLock, GetLock", function (assert)
	{
		const form = createApiSignatureForm();
		assert.strictEqual(form.GetLock(), false, "Check that a newly created signature form is unlocked");

		form.SetLock(true);
		assert.strictEqual(form.GetLock(), true, "Check that the form is locked after SetLock(true)");

		form.SetLock(false);
		assert.strictEqual(form.GetLock(), false, "Check that the form is unlocked after SetLock(false)");
	});

	QUnit.test("SetValue, GetValue, Value", function (assert)
	{
		const form = createApiSignatureForm();

		assert.strictEqual(typeof form.GetValue(), "string", "Check GetValue returns a string when no image is set");

		const result = form.SetValue("https://static.onlyoffice.com/assets/docs/samples/img/onlyoffice_logo.png");
		
		// TODO:
		assert.strictEqual(typeof form.Value, "string", "Check Value getter returns a string");
	});
});
