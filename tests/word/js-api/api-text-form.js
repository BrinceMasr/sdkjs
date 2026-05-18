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
	QUnit.module("Test the ApiTextForm methods");
	
	let logicDocument = AscTest.CreateLogicDocument();
	
	function createApiTextForm(pr)
	{
		pr = pr ? pr : {"key": "Name", "placeholder": "Enter your name"};
		return AscTest.JsApi.CreateTextForm(pr);
	}
	
	QUnit.test("Placeholder", function (assert)
	{
		let textForm = createApiTextForm({
			"key": "Name",
			"placeholder": "Enter your name"
		});
		
		assert.strictEqual(textForm.GetPlaceholderText(), "Enter your name" , "Check text form placeholder after the creation");
		
		textForm.SetPlaceholderText("TEST");
		assert.strictEqual(textForm.GetPlaceholderText(), "TEST" , "Check text form placeholder after reset placeholder text");
	});
	
	QUnit.test("Delete", function (assert)
	{
		AscTest.ClearDocument();
		let document = AscTest.JsApi.GetDocument();
		let p = AscTest.JsApi.CreateParagraph();
		document.Push(p);
		
		let textForm = createApiTextForm();
		
		p.AddText("Before");
		p.Push(textForm);
		p.AddText("After");
		
		assert.strictEqual(textForm.Sdt.IsUseInDocument(), true, "Check if text form were added");
		
		textForm.Delete(false);
		assert.strictEqual(textForm.Sdt.IsUseInDocument(), false, "Check if text form were deleted");
		assert.strictEqual(p.GetText(), "BeforeAfter\r\n", "Check paragraph text");
		
		p.RemoveAllElements();
		textForm.SetText("Inside");
		p.AddText("Before");
		p.Push(textForm);
		p.AddText("After");
		
		textForm.Delete(true);
		assert.strictEqual(textForm.Sdt.IsUseInDocument(), false, "Check if text form were deleted");
		assert.strictEqual(p.GetText(), "BeforeInsideAfter\r\n", "Check paragraph text");
	});
	
	QUnit.test('SetBorderColor, GetBorderColor', function (assert) 
	{
		const textForm = AscTest.JsApi.CreateTextForm();

		assert.strictEqual(textForm.GetBorderColor(), null, 'Check border color for a newly created text form');

		textForm.SetBorderColor(255, 122, 100);
		assert.equalRgb(textForm.GetBorderColor(), { r: 255, g: 122, b: 100 }, 'Check border color after setting it with rgba components');

		const hexColor = AscTest.JsApi.HexColor('a1b2c3');
		textForm.SetBorderColor(hexColor);
		assert.equalRgb(textForm.GetBorderColor(), { r: 161, g: 178, b: 195 }, 'Check border color after setting it with ApiColor (rgba)');

		textForm.SetBorderColor(0, 0, 0, true);
		assert.strictEqual(textForm.GetBorderColor(), null, 'Check border color after resetting it');
	});

	QUnit.test('SetBackgroundColor, GetBackgroundColor', function (assert) {
		const textForm = AscTest.JsApi.CreateTextForm();

		assert.strictEqual(textForm.GetBackgroundColor(), null, 'Check background color for a newly created text form');

		textForm.SetBackgroundColor(255, 122, 100);
		assert.equalRgb(textForm.GetBackgroundColor(), { r: 255, g: 122, b: 100 }, 'Check background color after setting it with rgba components');

		const hexColor = AscTest.JsApi.HexColor('a1b2c3');
		textForm.SetBackgroundColor(hexColor);
		assert.equalRgb(textForm.GetBackgroundColor().GetRGB(), { r: 161, g: 178, b: 195 }, 'Check background color after setting it with ApiColor (rgba)');

		const themeColor = AscTest.JsApi.ThemeColor('accent3');
		textForm.SetBackgroundColor(themeColor);
		assert.strictEqual(textForm.GetBackgroundColor().IsThemeColor(), true, 'Check background color after setting it with theme color');

		textForm.SetBackgroundColor(0, 0, 0, true);
		assert.strictEqual(textForm.GetBackgroundColor(), null, 'Check background color after resetting it');
	});
	
	QUnit.test('SetLock/GetLock', function (assert)
	{
		const textForm = AscTest.JsApi.CreateTextForm();

		assert.strictEqual(textForm.GetLock(), false, 'Check that a newly created text form is unlocked');

		textForm.SetLock(true);
		assert.strictEqual(textForm.GetLock(), true, 'Check that the text form is locked after SetLock(true)');

		let sdt = textForm.private_GetImpl();
		assert.strictEqual(sdt.GetContentControlLock(), c_oAscSdtLockType.SdtLocked, 'Check that the internal lock type is SdtLocked');
	});

	QUnit.test('IsFilled', function (assert)
	{
		const textForm = createApiTextForm();

		assert.strictEqual(textForm.IsFilled(), false, 'Check that a newly created text form is not filled');

		textForm.SetText("John Smith");
		assert.strictEqual(textForm.IsFilled(), true, 'Check that the text form is filled after SetText');

		textForm.Clear();
		assert.strictEqual(textForm.IsFilled(), false, 'Check that the text form is not filled after Clear');
	});

	QUnit.test('SetAllowedSymbols, GetAllowedSymbols', function (assert)
	{
		const textForm = createApiTextForm();

		assert.strictEqual(textForm.GetAllowedSymbols(), '', 'Check that a newly created text form has no allowed symbols restriction');

		textForm.SetAllowedSymbols('abc');
		assert.strictEqual(textForm.GetAllowedSymbols(), 'abc', 'Check allowed symbols after setting "abc"');

		textForm.SetAllowedSymbols('0123456789');
		assert.strictEqual(textForm.GetAllowedSymbols(), '0123456789', 'Check allowed symbols after setting digits string');

		textForm.SetAllowedSymbols('');
		assert.strictEqual(textForm.GetAllowedSymbols(), '', 'Check allowed symbols after clearing with empty string');
	});

	QUnit.test('SetFormat, GetFormat', function (assert)
	{
		const textForm = createApiTextForm();

		assert.deepEqual(textForm.GetFormat(), { type: 'none' }, 'Check that a newly created text form has no format');

		textForm.SetFormat({ type: 'digit' });
		assert.deepEqual(textForm.GetFormat(), { type: 'digit' }, 'Check format after setting digit type');

		textForm.SetFormat({ type: 'letter' });
		assert.deepEqual(textForm.GetFormat(), { type: 'letter' }, 'Check format after setting letter type');

		textForm.SetFormat({ type: 'mask', value: '9-9-9' });
		assert.deepEqual(textForm.GetFormat(), { type: 'mask', value: '9-9-9' }, 'Check format after setting mask type with value');

		textForm.SetFormat({ type: 'regExp', value: '\\d+' });
		assert.deepEqual(textForm.GetFormat(), { type: 'regExp', value: '\\d+' }, 'Check format after setting regExp type with value');

		textForm.SetFormat({ type: 'none' });
		assert.deepEqual(textForm.GetFormat(), { type: 'none' }, 'Check format after resetting to none');
	});

	QUnit.test('SetValue, GetValue, Value', function (assert)
	{
		const form = createApiTextForm({"key" : "key", "placeholder" : "123"});

		assert.strictEqual(form.GetValue(), '123', 'Check GetValue returns empty string for a newly created text form');

		const result = form.SetValue('Hello');
		assert.strictEqual(result, true, 'Check SetValue returns true on success');
		assert.strictEqual(form.GetValue(), 'Hello', 'Check GetValue returns the value set by SetValue');

		assert.strictEqual(form.Value, 'Hello', 'Check Value getter returns the current value');

		form.Value = 'World';
		assert.strictEqual(form.GetValue(), 'World', 'Check Value setter updates the text form value');
	});
});
