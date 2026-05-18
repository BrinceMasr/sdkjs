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

"use strict";

$(function () {

    const logicDocument = AscTest.CreateLogicDocument();
	QUnit.module("ApiShape");

    function CreateSlide()
	{
		logicDocument.addNextSlide(0);
		editor.WordControl.Thumbnails.CalculatePlaces();
	}

	QUnit.test("Test: SetPaddings", function (assert) {
		CreateSlide();

		const presentation = AscTest.JsApi.GetPresentation();
		const slide = presentation.GetSlideByIndex(0);

		const fill = AscTest.JsApi.CreateSolidFill(AscTest.JsApi.CreateRGBColor(200, 200, 200));
		const stroke = AscTest.JsApi.CreateStroke(0, AscTest.JsApi.CreateNoFill());
		const shape = AscTest.JsApi.CreateShape("rect", 300 * 36000, 150 * 36000, fill, stroke);

		slide.AddObject(shape);

		const docContent = shape.GetDocContent();
		const paragraph = AscTest.JsApi.CreateParagraph();
		paragraph.AddText("Test text with paddings");
		docContent.Push(paragraph);

		const result = shape.SetPaddings(4 * 36000, 2 * 36000, 3 * 36000, 5 * 36000);

		assert.strictEqual(result, true, 'SetPaddings should return true');

		const bodyPr = shape.Shape.getBodyPr();

		assert.ok(bodyPr.lIns === 4, 'Left padding should be set 4');
		assert.ok(bodyPr.tIns === 2, 'Top padding should be set 2');
		assert.ok(bodyPr.rIns === 3, 'Right padding should be set 3');
		assert.ok(bodyPr.bIns === 5, 'Bottom padding should be set 5');
	});
});
