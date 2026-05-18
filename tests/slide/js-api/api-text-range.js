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

"use strict";

$(function () {

	const logicDocument = AscTest.CreateLogicDocument();
	QUnit.module("ApiTextRange");

	function CreateSlide()
	{
		logicDocument.addNextSlide(0);
		editor.WordControl.Thumbnails.CalculatePlaces();
	}

	function createTextShape(text)
	{
		CreateSlide();

		const presentation = AscTest.JsApi.GetPresentation();
		const slide        = presentation.GetSlideByIndex(0);
		const fill         = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.CreateRGBColor(200, 200, 200));
		const stroke       = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());
		const shape        = AscTest.JsApi.CreateShape("rect", 300 * 36000, 150 * 36000, fill, stroke);
		slide.AddObject(shape);

		const docContent = shape.GetDocContent();
		const lines      = typeof text === "string" ? text.split("\r") : [""];

		let firstPara = docContent.GetElement(0);
		if (firstPara)
		{
			firstPara.AddText(lines[0]);
		}
		else
		{
			firstPara = AscTest.JsApi.CreateParagraph();
			firstPara.AddText(lines[0]);
			docContent.Push(firstPara);
		}

		for (let i = 1; i < lines.length; i++)
		{
			const para = AscTest.JsApi.CreateParagraph();
			para.AddText(lines[i]);
			docContent.Push(para);
		}

		return shape;
	}

	// ── Basic accessors ──────────────────────────────────────────────────────

	QUnit.test("Test: GetClassType", function (assert) {
		const range = createTextShape("Hello").GetTextRange();
		assert.strictEqual(range.GetClassType(), "textRange", "GetClassType returns 'textRange'");
	});

	QUnit.test("Test: GetText returns full content", function (assert) {
		const range = createTextShape("Hello World").GetTextRange();
		assert.strictEqual(range.GetText(), "Hello World\r\n", "GetText returns the shape text");
	});

	QUnit.test("Test: GetText with multi-paragraph content", function (assert) {
		const range = createTextShape("First\rSecond\rThird").GetTextRange();
		assert.strictEqual(range.GetText(), "First\r\nSecond\r\nThird\r\n", "GetText joins paragraphs with \\r\\n");
	});

	QUnit.test("Test: GetStartPos and GetEndPos", function (assert) {
		const range = createTextShape("Hello").GetTextRange();
		assert.strictEqual(range.GetStartPos(), 0, "GetStartPos is 0 for a full-range");
		assert.strictEqual(range.GetEndPos(), 5 + 2, "GetEndPos equals text length for -1 end");
	});

	QUnit.test("Test: Start and End property accessors", function (assert) {
		const range = createTextShape("Hello").GetTextRange();
		assert.strictEqual(range.Start, 0, "Start property equals GetStartPos()");
		assert.strictEqual(range.End, 5 + 2, "End property equals GetEndPos()");
	});

	QUnit.test("Test: SetStartPos and SetEndPos", function (assert) {
		const range = createTextShape("Hello World").GetTextRange();

		assert.strictEqual(range.SetStartPos(2), true, "SetStartPos returns true on valid input");
		assert.strictEqual(range.GetStartPos(), 2, "StartPos updated");

		assert.strictEqual(range.SetEndPos(7), true, "SetEndPos returns true on valid input");
		assert.strictEqual(range.GetEndPos(), 7, "EndPos updated");

		assert.strictEqual(range.GetText(), "llo W", "GetText reflects new range bounds");
	});

	QUnit.test("Test: SetStartPos rejects invalid input", function (assert) {
		const range = createTextShape("Hello").GetTextRange();
		assert.strictEqual(range.SetStartPos(-1), false, "SetStartPos returns false for negative value");
		assert.strictEqual(range.SetStartPos("x"), false, "SetStartPos returns false for non-number");
	});

	QUnit.test("Test: SetEndPos rejects invalid input", function (assert) {
		const range = createTextShape("Hello").GetTextRange();
		assert.strictEqual(range.SetEndPos(-1), false, "SetEndPos returns false for negative value");
		assert.strictEqual(range.SetEndPos(null), false, "SetEndPos returns false for null");
	});

	QUnit.test("Test: Start and End property setters", function (assert) {
		const range = createTextShape("Hello World").GetTextRange();
		range.Start = 1;
		range.End   = 6;
		assert.strictEqual(range.GetStartPos(), 1, "Start setter updates StartPos");
		assert.strictEqual(range.GetEndPos(), 6, "End setter updates EndPos");
	});

	// ── Paragraph access ─────────────────────────────────────────────────────

	QUnit.test("Test: GetParagraph returns paragraph at index", function (assert) {
		const range = createTextShape("Hello\rWorld").GetTextRange();
		const para0 = range.GetParagraph(0);
		const para1 = range.GetParagraph(1);
		assert.ok(para0 !== null, "GetParagraph(0) is not null");
		assert.ok(para0.GetText() === "Hello\r\n", "GetParagraph(0) is \"Hello\\r\\n\"");
		assert.ok(para1 !== null, "GetParagraph(1) is not null");
		assert.ok(para1.GetText() === "World\r\n", "GetParagraph(1) is \"World\\r\\n\"");
	});

	QUnit.test("Test: GetParagraph returns null for out-of-bounds index", function (assert) {
		const range = createTextShape("Hello").GetTextRange();
		assert.strictEqual(range.GetParagraph(5),  null, "Returns null for too-large index");
		assert.strictEqual(range.GetParagraph(-1), null, "Returns null for negative index");
	});

	QUnit.test("Test: GetAllParagraphs returns all paragraphs", function (assert) {
		const range = createTextShape("A\rB\rC").GetTextRange();
		const paras = range.GetAllParagraphs();
		assert.strictEqual(paras.length, 3, "GetAllParagraphs returns 3 paragraphs");
	});

	// ── Range operations ─────────────────────────────────────────────────────

	QUnit.test("Test: GetRange returns a sub-range by relative offsets", function (assert) {
		const range = createTextShape("Hello World").GetTextRange();
		const sub   = range.GetRange(6, 11);
		assert.ok(sub !== null, "GetRange returns a non-null range");
		assert.strictEqual(sub.GetText(), "World\r\n", "Sub-range text is 'World'");
		assert.strictEqual(sub.GetStartPos(), 6, "Sub-range start is 6");
		assert.strictEqual(sub.GetEndPos(), 11, "Sub-range end is 11");
	});

	QUnit.test("Test: GetRange with default end covers to range end", function (assert) {
		const range = createTextShape("Hello World").GetTextRange();
		const sub   = range.GetRange(6);
		assert.ok(sub !== null, "GetRange with no end returns non-null");
		assert.strictEqual(sub.GetText(), "World\r\n", "Sub-range text without explicit end is 'World'");
	});

	QUnit.test("Test: GetRange swaps nStart and nEnd when inverted", function (assert) {
		const range = createTextShape("Hello World").GetTextRange();
		const sub   = range.GetRange(11, 6);
		assert.ok(sub !== null, "GetRange with swapped args returns non-null");
		assert.strictEqual(sub.GetStartPos(), 6, "Start is corrected to 6");
	});

	QUnit.test("Test: GetRange returns null when start equals end", function (assert) {
		const range = createTextShape("Hello").GetTextRange();
		assert.strictEqual(range.GetRange(3, 3), null, "Returns null for zero-length range");
	});

	QUnit.test("Test: ExpandTo returns union of two ranges", function (assert) {
		const shape = createTextShape("Hello World");
		const range = shape.GetTextRange();
		range.SetStartPos(0);
		range.SetEndPos(5);
		const range2 = shape.GetTextRange();
		range2.SetStartPos(6);
		range2.SetEndPos(11);
		const union = range.ExpandTo(range2);
		assert.ok(union !== null, "ExpandTo returns a range");
		assert.strictEqual(union.GetText(), "Hello World\r\n", "Union spans both sub-ranges");
	});

	QUnit.test("Test: ExpandTo returns null for ranges from different doc contents", function (assert) {
		const range1 = createTextShape("Foo").GetTextRange();
		const range2 = createTextShape("Bar").GetTextRange();
		assert.strictEqual(range1.ExpandTo(range2), null, "Returns null for different doc contents");
	});

	QUnit.test("Test: ExpandTo accepts non-ApiTextRange argument", function (assert) {
		const range = createTextShape("Hello").GetTextRange();
		assert.strictEqual(range.ExpandTo(null), null, "Returns null for null argument");
		assert.strictEqual(range.ExpandTo("text"), null, "Returns null for string argument");
	});

	QUnit.test("Test: IntersectWith returns intersection of two ranges", function (assert) {
		const shape  = createTextShape("Hello World");
		const range1 = shape.GetTextRange();
		range1.SetStartPos(0);
		range1.SetEndPos(7);
		const range2 = shape.GetTextRange();
		range2.SetStartPos(5);
		range2.SetEndPos(11);
		const inter = range1.IntersectWith(range2);
		assert.ok(inter !== null, "IntersectWith returns a range");
		assert.strictEqual(inter.GetText(), " W", "Intersection is the overlapping part");
	});

	QUnit.test("Test: IntersectWith returns null when ranges do not overlap", function (assert) {
		const shape  = createTextShape("Hello World");
		const range1 = shape.GetTextRange();
		range1.SetStartPos(0);
		range1.SetEndPos(3);
		const range2 = shape.GetTextRange();
		range2.SetStartPos(6);
		range2.SetEndPos(11);
		assert.strictEqual(range1.IntersectWith(range2), null, "Returns null when ranges do not overlap");
	});

	QUnit.test("Test: IntersectWith returns null for different doc contents", function (assert) {
		const range1 = createTextShape("Foo").GetTextRange();
		const range2 = createTextShape("Bar").GetTextRange();
		assert.strictEqual(range1.IntersectWith(range2), null, "Returns null for different doc contents");
	});

	// ── Text manipulation ─────────────────────────────────────────────────────

	QUnit.test("Test: SetText replaces full content", function (assert) {
		const range = createTextShape("Old text").GetTextRange();
		range.SetText("New text");
		assert.strictEqual(range.GetText(), "New text\r\n", "SetText replaces full content");
	});

	QUnit.test("Test: SetText with multi-paragraph string", function (assert) {
		const range = createTextShape("one").GetTextRange();
		range.SetText("A\rB\rC");
		assert.strictEqual(range.GetText(), "A\r\nB\r\nC\r\n", "SetText accepts \\r as paragraph separator");
		assert.strictEqual(range.GetAllParagraphs().length, 3, "Creates three paragraphs");
	});

	QUnit.test("Test: SetText updates StartPos and EndPos", function (assert) {
		const range = createTextShape("Hello").GetTextRange();
		range.SetText("Hi");
		assert.strictEqual(range.GetStartPos(), 0, "StartPos is preserved");
		assert.strictEqual(range.GetEndPos(), 2, "EndPos reflects new text length");
	});

	QUnit.test("Test: SetText returns this", function (assert) {
		const range = createTextShape("Hi").GetTextRange();
		assert.ok(range.SetText("Hello") === range, "SetText returns this");
	});

	QUnit.test("Test: Find returns matching sub-range", function (assert) {
		const range = createTextShape("Hello World").GetTextRange();
		const found = range.Find("World");
		assert.ok(found !== null, "Find returns a range for existing text");
		assert.strictEqual(found.GetText(), "World\r\n", "Found range contains the search text");
	});

	QUnit.test("Test: Find returns null when text is not found", function (assert) {
		const range = createTextShape("Hello World").GetTextRange();
		assert.strictEqual(range.Find("xyz"), null, "Find returns null when text is not found");
	});

	QUnit.test("Test: Find is case-insensitive by default", function (assert) {
		const range = createTextShape("Hello World").GetTextRange();
		const found = range.Find("hello");
		assert.ok(found !== null, "Find matches case-insensitively by default");
		assert.strictEqual(found.GetText(), "Hello", "Found range contains original casing");
	});

	QUnit.test("Test: Find with bMatchCase=true is case-sensitive", function (assert) {
		const range = createTextShape("Hello World").GetTextRange();
		assert.strictEqual(range.Find("hello", 1, true), null, "Case-sensitive search does not match wrong case");
		assert.ok(range.Find("Hello", 1, true) !== null, "Case-sensitive search matches correct case");
	});

	QUnit.test("Test: Find with bWholeWords skips partial matches", function (assert) {
		const range = createTextShape("hell hello").GetTextRange();
		const found = range.Find("hell", 1, false, true);
		assert.ok(found !== null, "Find with bWholeWords finds the standalone word");
		assert.strictEqual(found.GetText(), "hell", "Matched standalone 'hell'");
	});

	QUnit.test("Test: Find with nAfter starts search from that position", function (assert) {
		const range  = createTextShape("aa bb aa").GetTextRange();
		const found1 = range.Find("aa", 1);
		const found2 = range.Find("aa", found1.GetEndPos() + 1);
		assert.ok(found2 !== null, "Second Find picks up the next occurrence");
		assert.ok(found2.GetStartPos() > found1.GetStartPos(), "Second match is further in the text");
	});

	QUnit.test("Test: Find returns null for empty search string", function (assert) {
		const range = createTextShape("Hello").GetTextRange();
		assert.strictEqual(range.Find(""), null, "Find returns null for empty string");
	});

	QUnit.test("Test: Replace replaces all occurrences", function (assert) {
		const range = createTextShape("foo bar foo").GetTextRange();
		range.Replace("foo", "baz");
		assert.strictEqual(range.GetText(), "baz bar baz\r\n", "Replace substitutes all occurrences");
	});

	QUnit.test("Test: Replace with case-sensitive mode", function (assert) {
		const range = createTextShape("Foo foo FOO").GetTextRange();
		range.Replace("foo", "bar", true);
		assert.strictEqual(range.GetText(), "Foo bar FOO\r\n", "Case-sensitive Replace only matches 'foo'");
	});

	QUnit.test("Test: Replace with whole-word mode skips partial matches", function (assert) {
		const range = createTextShape("foobar foo").GetTextRange();
		range.Replace("foo", "X", false, true);
		assert.strictEqual(range.GetText(), "foobar X\r\n", "Whole-word Replace skips 'foobar'");
	});

	QUnit.test("Test: Replace returns this", function (assert) {
		const range = createTextShape("hello").GetTextRange();
		assert.ok(range.Replace("hello", "world") === range, "Replace returns this");
	});

	QUnit.test("Test: Delete removes text and returns true", function (assert) {
		const shape  = createTextShape("Hello World");
		const full   = shape.GetTextRange();
		full.SetStartPos(0);
		full.SetEndPos(5);
		const result = full.Delete();
		assert.strictEqual(result, true, "Delete returns true");
	});

	QUnit.test("Test: Delete on empty range returns false", function (assert) {
		const range = createTextShape("Hi").GetTextRange();
		range.SetStartPos(2);
		range.SetEndPos(2);
		assert.strictEqual(range.Delete(), false, "Delete on empty range returns false");
	});

	// ── Text formatting ───────────────────────────────────────────────────────

	QUnit.test("Test: SetBold returns the range", function (assert) {
		const range = createTextShape("Bold text").GetTextRange();
		assert.ok(range.GetParagraph().GetElement(1).TextPr.Bold !== true, "Run isn't Bold");
		assert.ok(range.SetBold(true) === range, "SetBold returns this");
		assert.ok(range.GetParagraph().GetElement(1).TextPr.Bold === true, "Run is Bold");
	});

	QUnit.test("Test: SetItalic returns the range", function (assert) {
		const range = createTextShape("Italic text").GetTextRange();
		assert.ok(range.GetParagraph().GetElement(1).TextPr.Italic !== true, "Run isn't Italic");
		assert.ok(range.SetItalic(true) === range, "SetItalic returns this");
		assert.ok(range.GetParagraph().GetElement(1).TextPr.Italic === true, "Run is Italic");
	});

	QUnit.test("Test: SetUnderline returns the range", function (assert) {
		const range = createTextShape("Underline text").GetTextRange();
		assert.ok(range.GetParagraph().GetElement(1).TextPr.Underline !== true, "Run isn't Underline");
		assert.ok(range.SetUnderline(true) === range, "SetUnderline returns this");
		assert.ok(range.GetParagraph().GetElement(1).TextPr.Underline === true, "Run is Underline");
	});

	QUnit.test("Test: SetStrikeout returns the range", function (assert) {
		const range = createTextShape("Strike text").GetTextRange();
		assert.ok(range.GetParagraph().GetElement(1).TextPr.Strikeout !== true, "Run isn't Strikeout");
		assert.ok(range.SetStrikeout(true) === range, "SetStrikeout returns this");
		assert.ok(range.GetParagraph().GetElement(1).TextPr.Strikeout === true, "Run is Strikeout");
	});

	QUnit.test("Test: SetDoubleStrikeout returns the range", function (assert) {
		const range = createTextShape("Text").GetTextRange();
		assert.ok(range.GetParagraph().GetElement(1).TextPr.DStrikeout !== true, "Run isn't DStrikeout");
		assert.ok(range.SetDoubleStrikeout(true) === range, "SetDoubleStrikeout returns this");
		assert.ok(range.GetParagraph().GetElement(1).TextPr.DStrikeout === true, "Run is DStrikeout");
	});

	QUnit.test("Test: SetSmallCaps returns the range", function (assert) {
		const range = createTextShape("Text").GetTextRange();
		assert.ok(range.GetParagraph().GetElement(1).TextPr.SmallCaps !== true, "Run isn't SmallCaps");
		assert.ok(range.SetSmallCaps(true) === range, "SetSmallCaps returns this");
		assert.ok(range.GetParagraph().GetElement(1).TextPr.SmallCaps === true, "Run is SmallCaps");
	});

	QUnit.test("Test: SetCaps returns the range", function (assert) {
		const range = createTextShape("Text").GetTextRange();
		assert.ok(range.GetParagraph().GetElement(1).TextPr.Caps !== true, "Run isn't Caps");
		assert.ok(range.SetCaps(true) === range, "SetCaps returns this");
		assert.ok(range.GetParagraph().GetElement(1).TextPr.Caps === true, "Run is Caps");
	});

	QUnit.test("Test: SetFontSize returns the range", function (assert) {
		const range = createTextShape("Text").GetTextRange();
		assert.ok(range.GetParagraph().GetElement(1).TextPr.FontSize !== 24, "FontSize isn't 24 before");
		assert.ok(range.SetFontSize(24) === range, "SetFontSize returns this");
		assert.strictEqual(range.GetParagraph().GetElement(1).TextPr.FontSize, 24, "FontSize is 24 after");
	});

	QUnit.test("Test: SetFontFamily returns the range for valid font name", function (assert) {
		const range = createTextShape("Text").GetTextRange();
		assert.ok(range.GetParagraph().GetElement(1).TextPr.RFonts !== "Arial", "FontFamily isn't Arial before");
		assert.ok(range.SetFontFamily("Arial") === range, "SetFontFamily returns this for a valid font name");
		assert.strictEqual(range.GetParagraph().GetElement(1).TextPr.RFonts.Ascii.Name, "Arial", "FontFamily is Arial after");
	});

	QUnit.test("Test: SetFontFamily returns null for non-string argument", function (assert) {
		const range = createTextShape("Text").GetTextRange();
		assert.strictEqual(range.SetFontFamily(42), null, "SetFontFamily returns null for number");
	});

	QUnit.test("Test: SetVertAlign returns the range for valid type", function (assert) {
		const range = createTextShape("Text").GetTextRange();
		assert.ok(range.GetParagraph().GetElement(1).TextPr.VertAlign !== AscCommon.vertalign_SuperScript, "VertAlign isn't superscript before");
		assert.ok(range.SetVertAlign("superscript") === range, "SetVertAlign('superscript') returns this");
		assert.strictEqual(range.GetParagraph().GetElement(1).TextPr.VertAlign, AscCommon.vertalign_SuperScript, "VertAlign is superscript after");
		assert.ok(range.SetVertAlign("subscript") === range, "SetVertAlign('subscript') returns this");
		assert.strictEqual(range.GetParagraph().GetElement(1).TextPr.VertAlign, AscCommon.vertalign_SubScript, "VertAlign is subscript after");
		assert.ok(range.SetVertAlign("baseline") === range, "SetVertAlign('baseline') returns this");
		assert.strictEqual(range.GetParagraph().GetElement(1).TextPr.VertAlign, AscCommon.vertalign_Baseline, "VertAlign is baseline after");
	});

	QUnit.test("Test: SetVertAlign returns null for invalid type", function (assert) {
		const range = createTextShape("Text").GetTextRange();
		assert.strictEqual(range.SetVertAlign("invalid"), null, "SetVertAlign returns null for unknown type");
	});

	QUnit.test("Test: SetPosition returns null for non-number", function (assert) {
		const range = createTextShape("Text").GetTextRange();
		assert.strictEqual(range.SetPosition("x"), null, "SetPosition returns null for string argument");
	});

	QUnit.test("Test: SetPosition returns the range for valid input", function (assert) {
		const range = createTextShape("Text").GetTextRange();
		assert.ok(range.GetParagraph().GetElement(1).TextPr.Position === undefined || range.GetParagraph().GetElement(1).TextPr.Position === null || range.GetParagraph().GetElement(1).TextPr.Position === 0, "Position isn't set before");
		assert.ok(range.SetPosition(10) === range, "SetPosition returns this for a numeric argument");
		assert.ok(range.GetParagraph().GetElement(1).TextPr.Position !== undefined && range.GetParagraph().GetElement(1).TextPr.Position !== null && range.GetParagraph().GetElement(1).TextPr.Position !== 0, "Position is set after");
	});

	QUnit.test("Test: SetSpacing returns the range", function (assert) {
		const range = createTextShape("Text").GetTextRange();
		assert.ok(range.GetParagraph().GetElement(1).TextPr.Spacing === undefined || range.GetParagraph().GetElement(1).TextPr.Spacing === null || range.GetParagraph().GetElement(1).TextPr.Spacing === 0, "Spacing isn't set before");
		assert.ok(range.SetSpacing(40) === range, "SetSpacing returns this");
		assert.ok(range.GetParagraph().GetElement(1).TextPr.Spacing !== undefined && range.GetParagraph().GetElement(1).TextPr.Spacing !== null && range.GetParagraph().GetElement(1).TextPr.Spacing !== 0, "Spacing is set after");
	});

	QUnit.test("Test: SetColor with RGB values returns the range", function (assert) {
		const range = createTextShape("Text").GetTextRange();
		assert.ok(!range.GetParagraph().GetElement(1).TextPr.Color || range.GetParagraph().GetElement(1).TextPr.Color.r !== 255, "Color isn't red before");
		assert.ok(range.SetColor(255, 0, 0) === range, "SetColor returns this for RGB arguments");
		assert.strictEqual(range.GetParagraph().GetElement(1).TextPr.Color.r, 255, "Color.r is 255 after");
		assert.strictEqual(range.GetParagraph().GetElement(1).TextPr.Color.g, 0,   "Color.g is 0 after");
		assert.strictEqual(range.GetParagraph().GetElement(1).TextPr.Color.b, 0,   "Color.b is 0 after");
	});

	QUnit.test("Test: SetColor with ApiColor returns the range", function (assert) {
		const range = createTextShape("Text").GetTextRange();
		const color = AscTest.JsApi.CreateRGBColor(0, 128, 255);
		assert.ok(!range.GetParagraph().GetElement(1).TextPr.Color || pr.Color.b !== 255, "Color isn't blue before");
		assert.ok(range.SetColor(color) === range, "SetColor returns this for ApiColor argument");
		assert.strictEqual(range.GetParagraph().GetElement(1).TextPr.Color.r, 0,   "Color.r is 0 after");
		assert.strictEqual(range.GetParagraph().GetElement(1).TextPr.Color.g, 128, "Color.g is 128 after");
		assert.strictEqual(range.GetParagraph().GetElement(1).TextPr.Color.b, 255, "Color.b is 255 after");
	});

	QUnit.test("Test: GetTextPr returns a text property object", function (assert) {
		const range  = createTextShape("Text").GetTextRange();
		const textPr = range.GetTextPr();
		assert.ok(textPr !== null && textPr !== undefined, "GetTextPr returns a non-null object");
		assert.ok(typeof textPr.GetBold === "function", "Returned object has GetBold method");
		assert.ok(typeof textPr.GetItalic === "function", "Returned object has GetItalic method");
	});

	QUnit.test("Test: SetTextPr applies ApiTextPr properties", function (assert) {
		const range  = createTextShape("Text").GetTextRange();
		const textPr = range.GetTextPr();
		const result = range.SetTextPr(textPr);
		assert.ok(result === range, "SetTextPr returns this when given a valid ApiTextPr");
	});

	QUnit.test("Test: SetTextPr returns null for non-ApiTextPr argument", function (assert) {
		const range = createTextShape("Text").GetTextRange();
		assert.strictEqual(range.SetTextPr({}), null, "SetTextPr returns null for plain object");
		assert.strictEqual(range.SetTextPr(null), null, "SetTextPr returns null for null");
	});

	QUnit.test("Test: SetHighlight returns null for invalid color name", function (assert) {
		const range = createTextShape("Text").GetTextRange();
		assert.strictEqual(range.SetHighlight("notacolor"), null, "SetHighlight returns null for unknown color");
	});

});
