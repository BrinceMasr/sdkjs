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
	QUnit.module("ApiRun");

	function createApiRun()
	{
		return AscTest.JsApi.CreateRun();
	}
	
	QUnit.test('GetText/AddText', function (assert)
	{
		let run = createApiRun();
		
		assert.strictEqual(run.GetText(), "", "Check text for an empty run");
		run.AddText("1");
		run.AddTabStop();
		run.AddText("2");
		run.AddLineBreak();
		run.AddText("3");
		assert.strictEqual(run.GetText(), "1\t2\r3", "Check text");
		assert.strictEqual(run.GetText({
			"TabSymbol" : "_t_",
			"NewLineSeparator" : "_nl_",
		}), "1_t_2_nl_3", "Check text");
	});

	QUnit.test('SetColor, GetColor', function (assert) {
		const apiRun = createApiRun();

		const hexColor = AscTest.JsApi.HexColor('#bada55');
		const themeColor = AscTest.JsApi.ThemeColor('accent2');
		const autoColor = AscTest.JsApi.AutoColor();

		assert.strictEqual(apiRun.GetColor(), null, 'Color check for a newly created run');

		apiRun.SetColor(255, 127, 0);
		assert.equalRgb(apiRun.GetColor(), { r: 255, g: 127, b: 0 }, 'Color check after setting color with RGB components');

		apiRun.SetColor(hexColor);
		assert.equalRgba(apiRun.GetColor(), { r: 186, g: 218, b: 85, a: 255 }, 'Color check after setting color with ApiColor (hex)');

		apiRun.SetColor(themeColor);
		assert.strictEqual(apiRun.GetColor().IsThemeColor(), true, 'Color check after setting color with ApiColor (theme)');

		apiRun.SetColor(autoColor);
		assert.strictEqual(apiRun.GetColor().IsAutoColor(), true, 'Color check after setting color with ApiColor (auto)');
	});

	QUnit.test('SetShd, GetShd', function (assert) {
		const apiRun = createApiRun();

		const hexColor = AscTest.JsApi.HexColor('#bada55');
		const themeColor = AscTest.JsApi.ThemeColor('accent2');
		const autoColor = AscTest.JsApi.AutoColor();

		assert.equalShd(apiRun.GetTextPr().GetShd(), undefined, 'Shading check for a newly created run');

		apiRun.SetShd('clear', 255, 127, 0);
		assert.equalShd(apiRun.GetTextPr().GetShd(), { 'Type': 'clear', 'Color': AscTest.JsApi.RGB(255, 127, 0) }, 'Shading check after setting shading with RGB components');

		apiRun.SetShd('clear', hexColor);
		assert.equalShd(apiRun.GetTextPr().GetShd(), { 'Type': 'clear', 'Color': AscTest.JsApi.HexColor('#bada55') }, 'Shading check after setting shading with ApiColor (hex)');

		apiRun.SetShd('clear', themeColor);
		assert.equalShd(apiRun.GetTextPr().GetShd(), { 'Type': 'clear', 'Color': AscTest.JsApi.ThemeColor('accent2') }, 'Shading check after setting shading with ApiColor (theme)');

		apiRun.SetShd('clear', autoColor);
		assert.equalShd(apiRun.GetTextPr().GetShd(), { 'Type': 'clear', 'Color': AscTest.JsApi.AutoColor() }, 'Shading check after setting shading with ApiColor (auto)');
	});
});
