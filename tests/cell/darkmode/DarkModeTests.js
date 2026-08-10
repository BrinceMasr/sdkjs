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
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

$(function () {

	// =====================================================================
	// AscCommonExcel.isColorAutomatic
	// =====================================================================
	QUnit.module('isColorAutomatic');

	QUnit.test('null/undefined color is automatic', function (assert) {
		assert.strictEqual(AscCommonExcel.isColorAutomatic(null), true, 'null');
		assert.strictEqual(AscCommonExcel.isColorAutomatic(undefined), true, 'undefined');
	});

	QUnit.test('RgbColor with isAutoColor=true is automatic, regardless of rgb value', function (assert) {
		var autoBlack = AscCommonExcel.createRgbColor(0, 0, 0);
		autoBlack.isAutoColor = true;
		assert.strictEqual(AscCommonExcel.isColorAutomatic(autoBlack), true, 'auto-flagged black');

		var autoWhite = AscCommonExcel.createRgbColor(255, 255, 255);
		autoWhite.isAutoColor = true;
		assert.strictEqual(AscCommonExcel.isColorAutomatic(autoWhite), true, 'auto-flagged white');
	});

	QUnit.test('explicit RgbColor(0,0,0) is NOT automatic - literal black is not the auto default', function (assert) {
		var explicitBlack = AscCommonExcel.createRgbColor(0, 0, 0);
		assert.strictEqual(explicitBlack.isAutoColor, false, 'isAutoColor defaults to false on a fresh RgbColor');
		assert.strictEqual(AscCommonExcel.isColorAutomatic(explicitBlack), false,
			'explicit literal black must stay explicit, not be treated as automatic');
	});

	QUnit.test('RgbColor.clone() preserves isAutoColor', function (assert) {
		var auto = AscCommonExcel.createRgbColor(12, 34, 56);
		auto.isAutoColor = true;
		var clonedAuto = auto.clone();
		assert.strictEqual(clonedAuto.isAutoColor, true, 'cloned auto color stays automatic');
		assert.strictEqual(AscCommonExcel.isColorAutomatic(clonedAuto), true, 'isColorAutomatic agrees on the clone');

		var explicit = AscCommonExcel.createRgbColor(12, 34, 56);
		var clonedExplicit = explicit.clone();
		assert.strictEqual(clonedExplicit.isAutoColor, false, 'cloned explicit color stays explicit');
		assert.strictEqual(AscCommonExcel.isColorAutomatic(clonedExplicit), false, 'isColorAutomatic agrees on the clone');
	});

	QUnit.test('ThemeColor identity against g_oDefaultFormat.Font.c survives .clone()', function (assert) {
		var defaultFormat = AscCommonExcel.g_oDefaultFormat;
		var savedFont = defaultFormat.Font;

		try {
			var defaultFontColor = new AscCommonExcel.ThemeColor();
			defaultFormat.Font = {c: defaultFontColor};

			assert.strictEqual(AscCommonExcel.isColorAutomatic(defaultFontColor), true,
				'the workbook default font color is automatic by identity, not by a flag');

			// ThemeColor.clone() intentionally returns `this` (see WorkbookElems.js) - this is
			// the mechanism isColorAutomatic's identity check relies on to survive cloning
			// through charProperties. If a future change makes clone() return a real copy,
			// this assertion is what should start failing.
			var cloned = defaultFontColor.clone();
			assert.strictEqual(cloned, defaultFontColor, 'ThemeColor.clone() returns the same instance');
			assert.strictEqual(AscCommonExcel.isColorAutomatic(cloned), true,
				'identity survives the clone, so the cloned color is still seen as automatic');

			var otherThemeColor = new AscCommonExcel.ThemeColor();
			assert.strictEqual(AscCommonExcel.isColorAutomatic(otherThemeColor), false,
				'a different ThemeColor instance (not the default font color) is not automatic');
		} finally {
			defaultFormat.Font = savedFont;
		}
	});

	// =====================================================================
	// AscCommonExcel / Asc.DrawingContext.prototype.getDarkModeCorrectedColor
	// =====================================================================
	QUnit.module('getDarkModeCorrectedColor');

	// getDarkModeCorrectedColor only reads _darkModeRgbCache/_darkModeColorShuttle off `this`,
	// so a real DrawingContext (which needs a canvas, fmgrGraphics and font to construct) isn't
	// required - a manufactured `this` matching that shape is enough to exercise the real method.
	function makeFakeDrawingContext() {
		return {
			_darkModeRgbCache: {},
			_darkModeColorShuttle: new AscCommon.CColor(0, 0, 0, 1)
		};
	}

	QUnit.test('returns the shared shuttle instance, not a new object', function (assert) {
		var ctx = makeFakeDrawingContext();
		var result = Asc.DrawingContext.prototype.getDarkModeCorrectedColor.call(ctx, 10, 20, 30, 1);
		assert.strictEqual(result, ctx._darkModeColorShuttle, 'shuttle instance is returned as-is');
	});

	QUnit.test('repeated calls with the same rgb return the same corrected value (cache hit)', function (assert) {
		var ctx = makeFakeDrawingContext();
		Asc.DrawingContext.prototype.getDarkModeCorrectedColor.call(ctx, 10, 20, 30, 1);
		var first = {r: ctx._darkModeColorShuttle.getR(), g: ctx._darkModeColorShuttle.getG(), b: ctx._darkModeColorShuttle.getB()};

		Asc.DrawingContext.prototype.getDarkModeCorrectedColor.call(ctx, 10, 20, 30, 1);
		var second = {r: ctx._darkModeColorShuttle.getR(), g: ctx._darkModeColorShuttle.getG(), b: ctx._darkModeColorShuttle.getB()};

		assert.deepEqual(second, first, 'same input rgb produces the same corrected rgb on a cache hit');
		assert.strictEqual(Object.keys(ctx._darkModeRgbCache).length, 1, 'only one cache entry was created for one distinct color');
	});

	QUnit.test('shuttle is overwritten (not stale) across consecutive different-color calls', function (assert) {
		var ctx = makeFakeDrawingContext();

		Asc.DrawingContext.prototype.getDarkModeCorrectedColor.call(ctx, 0, 0, 0, 1);
		var forBlack = {r: ctx._darkModeColorShuttle.getR(), g: ctx._darkModeColorShuttle.getG(), b: ctx._darkModeColorShuttle.getB()};

		Asc.DrawingContext.prototype.getDarkModeCorrectedColor.call(ctx, 255, 255, 255, 1);
		var forWhite = {r: ctx._darkModeColorShuttle.getR(), g: ctx._darkModeColorShuttle.getG(), b: ctx._darkModeColorShuttle.getB()};

		assert.notDeepEqual(forWhite, forBlack,
			'the shuttle reflects the second call\'s color, not a leftover from the first call');
		assert.strictEqual(Object.keys(ctx._darkModeRgbCache).length, 2, 'two distinct colors produced two cache entries');
	});

	QUnit.test('alpha is passed through unmodified, only rgb goes through correction', function (assert) {
		var ctx = makeFakeDrawingContext();
		Asc.DrawingContext.prototype.getDarkModeCorrectedColor.call(ctx, 100, 100, 100, 0.5);
		assert.strictEqual(ctx._darkModeColorShuttle.getA(), 0.5, 'alpha is carried through as given, not corrected');
	});
});
