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

	// =====================================================================
	// WorksheetView.prototype._getKeepsAutomaticTextColorAsIs
	// =====================================================================
	QUnit.module('_getKeepsAutomaticTextColorAsIs');

	// Only this.handlers/this.workbook/this.settings.findFillColor and the passed-in cell's
	// Fill are touched, so a real WorksheetView (which needs a full editor boot to construct)
	// isn't required - a manufactured `this` matching that shape is enough.
	function makeFakeWorksheetView(opts) {
		opts = opts || {};
		return {
			handlers: {
				trigger: function (name) {
					return name === 'selectSearchingResults' ? !!opts.searchHighlightOn : false;
				}
			},
			workbook: {
				inFindResults: function () {
					return opts.isFindResult ? true : undefined;
				}
			},
			settings: {
				findFillColor: opts.findFillColor
			}
		};
	}

	function callKeepsAsIs(wsOpts, fill, resolvedFallbackBg) {
		var ws = makeFakeWorksheetView(wsOpts);
		var cell = {getFill: function () { return fill; }};
		return AscCommonExcel.WorksheetView.prototype._getKeepsAutomaticTextColorAsIs.call(ws, cell, 0, 0, resolvedFallbackBg);
	}

	function makeSolidFill(r, g, b) {
		var fill = new AscCommonExcel.Fill();
		fill.fromColor(AscCommonExcel.createRgbColor(r, g, b));
		return fill;
	}

	function makeGradientFill() {
		var fill = new AscCommonExcel.Fill();
		fill.gradientFill = new AscCommonExcel.GradientFill();
		return fill;
	}

	QUnit.test('no fill at all: does not keep as-is (needs correction, same as bare canvas)', function (assert) {
		var noFill = new AscCommonExcel.Fill();
		assert.notOk(noFill.hasFill(), 'sanity check: this Fill really has no fill (hasFill() can be null, not just false)');
		assert.strictEqual(callKeepsAsIs({}, noFill), false);
	});

	QUnit.test('solid light fill keeps automatic text as-is', function (assert) {
		assert.strictEqual(callKeepsAsIs({}, makeSolidFill(255, 255, 255)), true, 'white fill');
		assert.strictEqual(callKeepsAsIs({}, makeSolidFill(255, 255, 0)), true, 'yellow fill');
	});

	QUnit.test('solid dark fill does NOT keep automatic text as-is', function (assert) {
		assert.strictEqual(callKeepsAsIs({}, makeSolidFill(0, 0, 0)), false, 'black fill');
		assert.strictEqual(callKeepsAsIs({}, makeSolidFill(10, 10, 10)), false, 'near-black fill');
	});

	QUnit.test('pattern/gradient fill without resolvedFallbackBg: conservative true', function (assert) {
		assert.strictEqual(callKeepsAsIs({}, makeGradientFill()), true,
			'unknown per-pixel contrast is exempted from correction by default');
	});

	QUnit.test('pattern/gradient fill WITH resolvedFallbackBg: resolves against it instead of the conservative default', function (assert) {
		var gradient = makeGradientFill();
		assert.strictEqual(callKeepsAsIs({}, gradient, new AscCommon.CColor(255, 255, 255)), true,
			'light resolvedFallbackBg overrides the conservative true with a true - but for the real reason (contrast), not by default');
		assert.strictEqual(callKeepsAsIs({}, gradient, new AscCommon.CColor(0, 0, 0)), false,
			'dark resolvedFallbackBg flips the result to false, proving the fallback is actually consulted, not ignored');
	});

	QUnit.test('search-highlighted cell overrides the fill-based result entirely', function (assert) {
		// dark fill would normally return false; a light findFillColor while highlighted
		// must still return true, proving the highlight branch short-circuits before the
		// fill is ever consulted (not that it coincidentally agrees with it)
		var darkFill = makeSolidFill(0, 0, 0);
		assert.strictEqual(
			callKeepsAsIs({searchHighlightOn: true, isFindResult: true, findFillColor: new AscCommon.CColor(255, 255, 0)}, darkFill),
			true,
			'light search-highlight color wins over the cell\'s own dark fill'
		);

		// light fill would normally return true; a dark findFillColor while highlighted must
		// still return false, for the same reason in the opposite direction
		var lightFill = makeSolidFill(255, 255, 255);
		assert.strictEqual(
			callKeepsAsIs({searchHighlightOn: true, isFindResult: true, findFillColor: new AscCommon.CColor(0, 0, 0)}, lightFill),
			false,
			'dark search-highlight color wins over the cell\'s own light fill'
		);
	});

	QUnit.test('selectSearchingResults off, or cell not actually a find result: falls through to the fill', function (assert) {
		var lightFill = makeSolidFill(255, 255, 255);
		assert.strictEqual(
			callKeepsAsIs({searchHighlightOn: false, isFindResult: true, findFillColor: new AscCommon.CColor(0, 0, 0)}, lightFill),
			true,
			'search highlighting disabled entirely: falls through to the fill-based result'
		);
		assert.strictEqual(
			callKeepsAsIs({searchHighlightOn: true, isFindResult: false, findFillColor: new AscCommon.CColor(0, 0, 0)}, lightFill),
			true,
			'highlighting is on but this specific cell is not a find result: falls through to the fill-based result'
		);
	});

	// =====================================================================
	// AscCommonExcel.drawFillCell - regression pin for the opt-in guard (point 1 of the
	// PR #70 review) and for the exact bug class already fixed once in this branch's
	// history (commit 42705a02f4, conditional-formatting Data Bar colors)
	// =====================================================================
	QUnit.module('drawFillCell dark-mode guard');

	// setFillStyle/fillRect are the only ctx methods drawFillCell calls for a solid fill, so a
	// manufactured ctx recording what it was asked to draw is enough - no real canvas needed.
	function makeFakeFillCtx(isDarkMode) {
		var lastFillColor = null;
		return {
			isDarkMode: !!isDarkMode,
			_darkModeRgbCache: {},
			_darkModeColorShuttle: new AscCommon.CColor(0, 0, 0, 1),
			getDarkModeCorrectedColor: Asc.DrawingContext.prototype.getDarkModeCorrectedColor,
			setFillStyle: function (color) {
				lastFillColor = color;
				return this;
			},
			fillRect: function () {
			},
			getLastFillColor: function () {
				return lastFillColor;
			}
		};
	}

	function makeRect() {
		return new AscCommon.asc_CRect(0, 0, 10, 10);
	}

	QUnit.test('caller that omits the flag draws the literal color unmodified in dark mode (the fixed bug)', function (assert) {
		var ctx = makeFakeFillCtx(true);
		var fill = new AscCommonExcel.Fill();
		fill.fromColor(AscCommonExcel.createRgbColor(10, 20, 30));

		AscCommonExcel.drawFillCell(ctx, null, fill, makeRect());

		var drawn = ctx.getLastFillColor();
		assert.strictEqual(drawn.getR(), 10, 'red channel untouched');
		assert.strictEqual(drawn.getG(), 20, 'green channel untouched');
		assert.strictEqual(drawn.getB(), 30, 'blue channel untouched');
	});

	QUnit.test('caller that explicitly opts in gets the dark-mode-corrected color', function (assert) {
		var ctx = makeFakeFillCtx(true);
		var fill = new AscCommonExcel.Fill();
		fill.fromColor(AscCommonExcel.createRgbColor(0, 0, 0));

		AscCommonExcel.drawFillCell(ctx, null, fill, makeRect(), true);

		var drawn = ctx.getLastFillColor();
		assert.notStrictEqual(drawn.getR(), 0, 'literal black must have actually been corrected, not left as-is');
	});

	QUnit.test('outside dark mode, the flag has no effect either way', function (assert) {
		var ctx = makeFakeFillCtx(false);
		var fill = new AscCommonExcel.Fill();
		fill.fromColor(AscCommonExcel.createRgbColor(0, 0, 0));

		AscCommonExcel.drawFillCell(ctx, null, fill, makeRect(), true);

		var drawn = ctx.getLastFillColor();
		assert.strictEqual(drawn.getR(), 0, 'light mode never corrects, even when bIsFillRecolorable is true');
	});

	// =====================================================================
	// AscCommon.updateGlobalSkin - regression pin for the theme-corruption bug fixed in
	// commit 0f55894fb7 (interface theme switches were silently corrupting content dark
	// mode's cell background/grid color through a shared GlobalSkin/EditorSkins alias)
	// =====================================================================
	QUnit.module('interface theme switch must not corrupt cell background/grid colors');

	QUnit.test('canvas-cell-background/canvas-cell-grid no longer alias into GlobalSkin.CellBackground/CellGrid', function (assert) {
		var before = AscCommon.GlobalSkin;
		var savedCellBackground = before.CellBackground;
		var savedCellGrid = before.CellGrid;
		var savedBackground = before.Background;

		try {
			// simulates a web-apps interface-theme switch supplying CSS-custom-property-backed
			// values, including the two keys that used to alias straight into CellBackground/CellGrid
			AscCommon.updateGlobalSkin({
				type: 'light',
				'canvas-cell-background': '#010203',
				'canvas-cell-grid': '#040506',
				'canvas-cell-title-background': '#0a0b0c'
			});

			var after = AscCommon.GlobalSkin;
			assert.strictEqual(after.CellBackground, savedCellBackground,
				'interface theme switch must not overwrite content dark mode\'s cell background color');
			assert.strictEqual(after.CellGrid, savedCellGrid,
				'interface theme switch must not overwrite content dark mode\'s cell grid color');

			// sanity check the update mechanism itself still works for a color that IS still
			// coupled (Background <- canvas-cell-title-background), so the two assertions above
			// are proven by an actually-removed mapping, not by updateGlobalSkin silently no-oping
			assert.notStrictEqual(after.Background, savedBackground,
				'a genuinely-mapped color changes, proving updateGlobalSkin is not a global no-op');
		} finally {
			var skin = AscCommon.GlobalSkin;
			skin.CellBackground = savedCellBackground;
			skin.CellGrid = savedCellGrid;
			skin.Background = savedBackground;
			delete skin['canvas-cell-background'];
			delete skin['canvas-cell-grid'];
			delete skin['canvas-cell-title-background'];
			delete skin.type;
		}
	});
});
