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
	QUnit.module('Test the ApiBlockLevelSdt methods');

	QUnit.test('SetColor, GetColor', function (assert)
	{
		const apiBlockCC = AscTest.JsApi.CreateBlockLvlSdt();

		assert.strictEqual(apiBlockCC.GetColor(), null, 'Check color for a newly created block-level content control');

		const rgbColor = AscTest.JsApi.RGB(60, 120, 180);
		apiBlockCC.SetColor(rgbColor);
		assert.equalRgb(apiBlockCC.GetColor(), { r: 60, g: 120, b: 180 }, 'Check color after setting with ApiColor');

		apiBlockCC.SetColor(null);
		assert.strictEqual(apiBlockCC.GetColor(), null, 'Check color after removing it');
	});
});
