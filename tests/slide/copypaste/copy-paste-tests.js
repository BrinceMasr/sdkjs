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

	QUnit.module("Slide Copy Paste Tests");

	let logicDocument = AscTest.CreateLogicDocument();
	AscTest.Editor.WordControl.m_oDrawingDocument.m_oLogicDocument = logicDocument;
	AscTest.Editor.WordControl.m_oLogicDocument = logicDocument;
    AscCommon.PasteElementsId.g_bIsDocumentCopyPaste = false;

    //TODO added only negative tests. need add slide and positive tests
	QUnit.test("Test: \"callback tests paste plain text\"", function (assert) {
		let done = assert.async();
		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.Text, "test", undefined, undefined, undefined, function (success) {
			assert.ok(true);
			done();
		});
	});

	QUnit.test("Test: \"callback tests paste HTML\"", function (assert) {
		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "test HTML content";
		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function (success) {
			assert.ok(true);
			done();
		});
	});

	QUnit.test("Test: \"callback tests paste Internal format\"", function (assert) {
		let done = assert.async();
		let binaryData = "";
		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.Internal, binaryData, undefined, undefined, undefined, function () {
			assert.ok(true);
			done();
		});
	});

	QUnit.module("Slide Copy Paste Tests");
});
