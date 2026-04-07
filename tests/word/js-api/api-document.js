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
	QUnit.module('Test the ApiDocument methods');

	QUnit.test('SetControlsHighlight, GetControlsHighlight', function (assert)
	{
		const doc = AscTest.JsApi.GetDocument();

		assert.strictEqual(doc.GetControlsHighlight(), null, 'Check that controls highlight is null by default');

		doc.SetControlsHighlight(255, 122, 100);
		assert.equalRgb(doc.GetControlsHighlight().GetRGB(), { r: 255, g: 122, b: 100 }, 'Check controls highlight after setting with RGB components');

		const hexColor = AscTest.JsApi.HexColor('a1b2c3');
		doc.SetControlsHighlight(hexColor);
		assert.equalRgb(doc.GetControlsHighlight().GetRGB(), { r: 161, g: 178, b: 195 }, 'Check controls highlight after setting with ApiColor');

		doc.SetControlsHighlight(0, 0, 0, true);
		assert.strictEqual(doc.GetControlsHighlight(), null, 'Check that controls highlight is null after disabling it');
	});
});
