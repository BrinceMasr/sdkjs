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
	QUnit.module('ApiParaPr');

	QUnit.test('SetContextualSpacing, GetContextualSpacing', function (assert)
	{
		const paragraph = AscTest.JsApi.CreateParagraph();
		const paraPr = paragraph.GetParaPr();

		assert.strictEqual(paraPr.GetContextualSpacing(), false, 'Contextual spacing is false by default for a newly created paragraph');

		paraPr.SetContextualSpacing(true);
		assert.strictEqual(paraPr.GetContextualSpacing(), true, 'Contextual spacing is true after setting it to true');

		paraPr.SetContextualSpacing(false);
		assert.strictEqual(paraPr.GetContextualSpacing(), false, 'Contextual spacing is false after setting it to false');
	});

	QUnit.test('SetKeepLines, GetKeepLines', function (assert)
	{
		const paragraph = AscTest.JsApi.CreateParagraph();
		const paraPr = paragraph.GetParaPr();

		assert.strictEqual(paraPr.GetKeepLines(), false, 'Keep lines is false by default for a newly created paragraph');

		paraPr.SetKeepLines(true);
		assert.strictEqual(paraPr.GetKeepLines(), true, 'Keep lines is true after setting it to true');

		paraPr.SetKeepLines(false);
		assert.strictEqual(paraPr.GetKeepLines(), false, 'Keep lines is false after setting it to false');
	});

	QUnit.test('SetKeepNext, GetKeepNext', function (assert)
	{
		const paragraph = AscTest.JsApi.CreateParagraph();
		const paraPr = paragraph.GetParaPr();

		assert.strictEqual(paraPr.GetKeepNext(), false, 'Keep next is false by default for a newly created paragraph');

		paraPr.SetKeepNext(true);
		assert.strictEqual(paraPr.GetKeepNext(), true, 'Keep next is true after setting it to true');

		paraPr.SetKeepNext(false);
		assert.strictEqual(paraPr.GetKeepNext(), false, 'Keep next is false after setting it to false');
	});

	QUnit.test('SetPageBreakBefore, GetPageBreakBefore', function (assert)
	{
		const paragraph = AscTest.JsApi.CreateParagraph();
		const paraPr = paragraph.GetParaPr();

		assert.strictEqual(paraPr.GetPageBreakBefore(), false, 'Page break before is false by default for a newly created paragraph');

		paraPr.SetPageBreakBefore(true);
		assert.strictEqual(paraPr.GetPageBreakBefore(), true, 'Page break before is true after setting it to true');

		paraPr.SetPageBreakBefore(false);
		assert.strictEqual(paraPr.GetPageBreakBefore(), false, 'Page break before is false after setting it to false');
	});

	QUnit.test('SetBottomBorder, GetBottomBorder', function (assert)
	{
		const paragraph = AscTest.JsApi.CreateParagraph();
		const paraPr = paragraph.GetParaPr();

		assert.equalBorder(paraPr.GetBottomBorder(), { 'Type': 'none' }, 'Bottom border is none for a newly created paragraph');

		paraPr.SetBottomBorder('single', 24, 0, 255, 111, 61);
		assert.equalBorder(paraPr.GetBottomBorder(), { 'Type': 'single', 'Size': 24, 'Space': 0, 'Color': AscTest.JsApi.RGB(255, 111, 61) }, 'Bottom border after set');
	});

	QUnit.test('SetLeftBorder, GetLeftBorder', function (assert)
	{
		const paragraph = AscTest.JsApi.CreateParagraph();
		const paraPr = paragraph.GetParaPr();

		assert.equalBorder(paraPr.GetLeftBorder(), { 'Type': 'none' }, 'Left border is none for a newly created paragraph');

		paraPr.SetLeftBorder('single', 16, 2, 0, 128, 255);
		assert.equalBorder(paraPr.GetLeftBorder(), { 'Type': 'single', 'Size': 16, 'Space': 2, 'Color': AscTest.JsApi.RGB(0, 128, 255) }, 'Left border after set');
	});

	QUnit.test('SetRightBorder, GetRightBorder', function (assert)
	{
		const paragraph = AscTest.JsApi.CreateParagraph();
		const paraPr = paragraph.GetParaPr();

		assert.equalBorder(paraPr.GetRightBorder(), { 'Type': 'none' }, 'Right border is none for a newly created paragraph');

		paraPr.SetRightBorder('single', 8, 0, 100, 200, 50);
		assert.equalBorder(paraPr.GetRightBorder(), { 'Type': 'single', 'Size': 8, 'Space': 0, 'Color': AscTest.JsApi.RGB(100, 200, 50) }, 'Right border after set');
	});

	QUnit.test('SetTopBorder, GetTopBorder', function (assert)
	{
		const paragraph = AscTest.JsApi.CreateParagraph();
		const paraPr = paragraph.GetParaPr();

		assert.equalBorder(paraPr.GetTopBorder(), { 'Type': 'none' }, 'Top border is none for a newly created paragraph');

		paraPr.SetTopBorder('single', 24, 0, 255, 0, 0);
		assert.equalBorder(paraPr.GetTopBorder(), { 'Type': 'single', 'Size': 24, 'Space': 0, 'Color': AscTest.JsApi.RGB(255, 0, 0) }, 'Top border after set');
	});

	QUnit.test('SetBetweenBorder, GetBetweenBorder', function (assert)
	{
		const paragraph = AscTest.JsApi.CreateParagraph();
		const paraPr = paragraph.GetParaPr();

		assert.equalBorder(paraPr.GetBetweenBorder(), { 'Type': 'none' }, 'Between border is none for a newly created paragraph');

		paraPr.SetBetweenBorder('single', 12, 0, 0, 0, 128);
		assert.equalBorder(paraPr.GetBetweenBorder(), { 'Type': 'single', 'Size': 12, 'Space': 0, 'Color': AscTest.JsApi.RGB(0, 0, 128) }, 'Between border after set');
	});

	QUnit.test('SetWidowControl, GetWidowControl', function (assert)
	{
		const paragraph = AscTest.JsApi.CreateParagraph();
		const paraPr = paragraph.GetParaPr();

		assert.strictEqual(paraPr.GetWidowControl(), true, 'Widow control is true for a newly created paragraph. Inherited from the default paragraph style.');

		paraPr.SetWidowControl(true);
		assert.strictEqual(paraPr.GetWidowControl(), true, 'Widow control is true after setting it to true');

		paraPr.SetWidowControl(false);
		assert.strictEqual(paraPr.GetWidowControl(), false, 'Widow control is false after setting it to false');
	});

	QUnit.test('SetTabs, GetTabs', function (assert)
	{
		const paragraph = AscTest.JsApi.CreateParagraph();
		const paraPr = paragraph.GetParaPr();

		assert.deepEqual(paraPr.GetTabs(), [], 'Tabs are empty for a newly created paragraph');

		paraPr.SetTabs([1440, 4320], ['left', 'right']);
		const tabs = paraPr.GetTabs();
		assert.strictEqual(tabs.length, 2, 'Two tab stops are set');
		assert.strictEqual(tabs[0].Pos, 1440, 'First tab stop position is 1440 twips');
		assert.strictEqual(tabs[0].Val, 'left', 'First tab stop type is left');
		assert.strictEqual(tabs[0].Leader, 'none', 'First tab stop leader is none');
		assert.strictEqual(tabs[1].Pos, 4320, 'Second tab stop position is 4320 twips');
		assert.strictEqual(tabs[1].Val, 'right', 'Second tab stop type is right');
		assert.strictEqual(tabs[1].Leader, 'none', 'Second tab stop leader is none');
	});

	QUnit.test('SetNumPr, GetNumPr', function (assert)
	{
		const paragraph = AscTest.JsApi.CreateParagraph();
		const paraPr = paragraph.GetParaPr();

		assert.strictEqual(paraPr.GetNumPr(), undefined, 'Numbering is undefined for a newly created paragraph');

		const doc = AscTest.JsApi.GetDocument();
		const numbering = doc.CreateNumbering('bullet');
		paraPr.SetNumPr(numbering, 0);
		const numPr = paraPr.GetNumPr();
		assert.ok(numPr !== undefined, 'Numbering is set after calling SetNumPr');
		assert.strictEqual(numPr.GetClassType(), 'numberingLevel', 'GetNumPr returns an ApiNumberingLevel object');
	});
});
