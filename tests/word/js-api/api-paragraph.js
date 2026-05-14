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
	QUnit.module("ApiParagraph");
	
	function createApiParagraph()
	{
		return AscTest.JsApi.CreateParagraph();
	}
	
	QUnit.test("ParaId", function (assert)
	{
		let apiParagraph = createApiParagraph();
		apiParagraph.SetParaId(0x48151623);
		assert.strictEqual(apiParagraph.GetParaId(), 0x48151623, "Check paraId");
	});
	
	QUnit.test("GetText", function (assert)
	{
		let p = createApiParagraph();
		let run = p.AddText("123");
		run.AddTabStop();
		run.AddText("456");
		run.AddLineBreak();
		run.AddText("789");
		assert.strictEqual(p.GetText(), "123\t456\r789\r\n", "Check GetText");
		assert.strictEqual(p.GetText({
			"TabSymbol" : "_t_",
			"NewLineSeparator" : "_nl_"
		}), "123_t_456_nl_789\r\n", "Check GetText");
	});
	
	QUnit.test('SetShd, GetShd', function (assert)
	{
		const apiParagraph = createApiParagraph();

		assert.equalShd(apiParagraph.GetShd(), {'Type' : 'nil', 'Color' : AscTest.JsApi.RGB(255, 255, 255)}, 'Shading check for a newly created paragraph');

		apiParagraph.SetShd('clear', 255, 122, 100);
		assert.equalShd(apiParagraph.GetShd(), { 'Type' : 'clear', 'Color' : AscTest.JsApi.RGB(255, 122, 100) }, 'Check shd color set with RGB components');

		apiParagraph.SetShd('clear', AscTest.JsApi.HexColor('55aa00'));
		assert.equalShd(apiParagraph.GetShd(), { 'Type': 'clear', 'Color': AscTest.JsApi.RGB(85, 170, 0) }, 'Check shd color set with ApiColor (hex)');

		apiParagraph.SetShd('clear', AscTest.JsApi.ThemeColor('accent2'));
		assert.equalShd(apiParagraph.GetShd(), { 'Type': 'clear', 'Color': AscTest.JsApi.ThemeColor('accent2') }, 'Check shd color set with ApiColor (theme)');

		apiParagraph.SetShd('clear', AscTest.JsApi.AutoColor());
		assert.equalShd(apiParagraph.GetShd(), { 'Type': 'clear', 'Color': AscTest.JsApi.AutoColor() }, 'Check shd color set with ApiColor (auto)');
	});
	
	QUnit.test('GetRange', function (assert)
	{
		const detachedParagraph = AscTest.JsApi.CreateParagraph();
		assert.throws(
			function() { detachedParagraph.GetRange(); },
			/Paragraph must be attached to document before getting its range/,
			"GetRange throws when paragraph is not attached to document"
		);
		
		const doc = AscTest.JsApi.GetDocument();
		const apiParagraph = AscTest.JsApi.CreateParagraph();
		apiParagraph.AddText("Hello World");
		doc.Push(apiParagraph);
		
		const fullRange = apiParagraph.GetRange();
		assert.ok(fullRange !== null, "GetRange returns non-null range for attached paragraph");
		assert.strictEqual(fullRange.GetText(), "Hello World\r\n", "Full range text matches paragraph text");
		
		const partialRange = apiParagraph.GetRange(0, 5);
		assert.ok(partialRange !== null, "GetRange with bounds returns non-null range");
		assert.strictEqual(partialRange.GetText(), "Hello", "Partial range text matches expected substring");
	});
	
	QUnit.test('Delete', function (assert)
	{
		const detachedParagraph = createApiParagraph();
		assert.strictEqual(detachedParagraph.Delete(), false, "Delete returns false for a detached paragraph");

		const doc = AscTest.JsApi.GetDocument();
		doc.Push(createApiParagraph());
		const p = createApiParagraph();
		p.AddText("To be deleted");
		doc.Push(p);

		assert.strictEqual(doc.GetElementsCount(), 2, "Document has 2 paragraphs before deletion");
		assert.strictEqual(p.Delete(), true, "Delete returns true for an attached paragraph");
		assert.strictEqual(doc.GetElementsCount(), 1, "Document has 1 paragraph after deletion");
		assert.strictEqual(p.Delete(), false, "Delete returns false — paragraph is already detached");
		assert.strictEqual(doc.GetElementsCount(), 1, "Document element count unchanged after deleting a detached paragraph");
	});

	QUnit.test('Delete with TrackRevisions', function (assert)
	{
		const doc = AscTest.JsApi.GetDocument();
		const p = createApiParagraph();
		const run = p.AddText("To be deleted");
		doc.Push(p);
		doc.Push(createApiParagraph());
		doc.Push(createApiParagraph());
		
		AscTest.SetTrackRevisions(true);
		assert.strictEqual(doc.GetElementsCount(), 3, "Document has 2 paragraphs before deletion");
		assert.strictEqual(p.Delete(), true, "Delete returns true when TrackRevisions is on");
		assert.strictEqual(doc.GetElementsCount(), 3, "Paragraph stays in the document as a pending deletion revision");
		assert.strictEqual(p.Paragraph.GetReviewType(), reviewtype_Remove, "Paragraph is marked as removed in review");
		assert.strictEqual(run.Run.GetReviewType(), reviewtype_Remove, "Text run inside the paragraph is marked as removed in review");
	});

	QUnit.test("SetText", function (assert)
	{
		let p = createApiParagraph();
		p.AddText("Hello World");
		assert.strictEqual(p.GetText(), "Hello World\r\n", "Check initial text");
		p.SetText("Replaced");
		assert.strictEqual(p.GetText(), "Replaced\r\n", "Check SetText replaces all content");

		p.SetText("");
		assert.strictEqual(p.GetText(), "\r\n", "Check SetText with empty string");
	});

	QUnit.test('SetColor, GetColor', function (assert) {
		const hexColor = AscTest.JsApi.HexColor('#bada55');
		const themeColor = AscTest.JsApi.ThemeColor('accent2');
		const autoColor = AscTest.JsApi.AutoColor();

		const apiParagraph = createApiParagraph();
		apiParagraph.AddText('Run for testing paragraph color');

		let apiRun;

		apiRun = apiParagraph.GetElement(0);
		assert.strictEqual(apiRun.GetColor(), null, 'Color check for a newly created paragraph');

		apiParagraph.SetColor(80, 160, 240);
		apiRun = apiParagraph.GetElement(0);
		assert.equalRgb(apiRun.GetColor(), { r: 80, g: 160, b: 240 }, 'Color check after setting color with RGB components');

		apiParagraph.SetColor(hexColor);
		apiRun = apiParagraph.GetElement(0);
		assert.strictEqual(apiRun.GetColor().GetHex(), '#BADA55', 'Color check after setting color with ApiColor (hex)');

		apiParagraph.SetColor(themeColor);
		apiRun = apiParagraph.GetElement(0);
		assert.strictEqual(apiRun.GetColor().IsThemeColor(), true, 'Color check after setting color with ApiColor (theme)');

		apiParagraph.SetColor(autoColor);
		apiRun = apiParagraph.GetElement(0);
		assert.strictEqual(apiRun.GetColor().IsAutoColor(), true, 'Color check after setting color with ApiColor (auto)');
	});

});
