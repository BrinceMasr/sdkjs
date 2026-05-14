/*
 * (c) Copyright Ascensio System SIA 2010-2024
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

$(function () {
	const CU = AscCommon.ColorUtils;

	function rgb(R, G, B) { return {R: R, G: G, B: B}; }

	QUnit.module("ColorUtils :: WCAG luminance");
	QUnit.test("anchors", function (assert) {
		assert.strictEqual(CU.relativeLuminance(rgb(0, 0, 0)), 0, "black -> 0");
		assert.strictEqual(CU.relativeLuminance(rgb(255, 255, 255)), 1, "white -> 1");
		assert.strictEqual(CU.wcagContrastRatio(1, 0), 21, "white/black ratio === 21");
	});

	QUnit.module("ColorUtils :: RGB <-> HSL");
	QUnit.test("gray maps to S=0", function (assert) {
		assert.deepEqual(CU.rgbToHsl(rgb(128, 128, 128)), {H: 0, S: 0, L: 128}, "gray 128");
		assert.deepEqual(CU.hslToRgb({H: 0, S: 0, L: 128}, true), rgb(128, 128, 128), "S=0 returns R=G=B=L");
	});
	QUnit.test("round-trip primaries", function (assert) {
		[rgb(255, 0, 0), rgb(0, 255, 0), rgb(0, 0, 255), rgb(255, 255, 255), rgb(0, 0, 0)].forEach(function (c) {
			assert.deepEqual(CU.hslToRgb(CU.rgbToHsl(c), true), c, "round-trip " + JSON.stringify(c));
		});
	});

	QUnit.module("ColorUtils :: LAB");
	QUnit.test("anchors", function (assert) {
		const black = CU.rgbToLab(rgb(0, 0, 0));
		assert.strictEqual(black.L, 0, "L(black) === 0");
		assert.strictEqual(black.a, 0, "a(black) === 0");
		assert.strictEqual(black.b, 0, "b(black) === 0");

		const white = CU.rgbToLab(rgb(255, 255, 255));
		assert.ok(Math.abs(white.L - 100) < 1e-2, "L(white) ≈ 100");
		assert.ok(Math.abs(white.a) < 1e-2, "a(white) ≈ 0");
		assert.ok(Math.abs(white.b) < 1e-2, "b(white) ≈ 0");
	});
	QUnit.module("ColorUtils :: hex");
	QUnit.test("byteToHex", function (assert) {
		assert.strictEqual(CU.byteToHex(0), "00");
		assert.strictEqual(CU.byteToHex(15), "0F");
		assert.strictEqual(CU.byteToHex(255), "FF");
	});
});
