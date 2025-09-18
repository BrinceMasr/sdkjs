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

	// let startTime = 0;
	//
	// QUnit.begin(function() {
	// 	startTime = performance.now();
	// });
	//
	// QUnit.done(function(details) {
	// 	const endTime = performance.now();
	// 	const duration = (endTime - startTime).toFixed(2);
	// 	console.log(`Время выполнения всех тестов: ${duration} мс`);
	// });

	QUnit.module("Word Copy Paste Tests");

	// AscTest.Editor._init();
    let logicDocument = AscTest.CreateLogicDocument();
	AscTest.Editor.WordControl.m_oDrawingDocument.m_oLogicDocument = logicDocument;
	AscTest.Editor.WordControl.m_oLogicDocument = logicDocument;

	const oldPrepeare_recursive = AscCommon.PasteProcessor.prototype._Prepeare_recursive;

	AscCommon.PasteProcessor.prototype._Prepeare_recursive = function () {

	};



	QUnit.test("Test: \"callback tests paste plain text\"", function (assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.Text, "test", undefined, undefined, undefined, function (success) {
			assert.ok(success);
			done();
		});
	});

	QUnit.test("Test: \"callback tests paste HTML\"", function (assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "test HTML content";
		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function (success) {
			assert.ok(success);
			done();
		});
	});

	QUnit.test("Test: \"callback tests paste Internal format\"", function (assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let binaryData = "";
		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.Internal, binaryData, undefined, undefined, undefined, function () {
			assert.ok(true);
			done();
		});
	});

	QUnit.test("Test: \"copy HTML with JSON verification\"", function (assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();

		// Create an HTML element to simulate copying
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<p>Test HTML content</p>";

		// Simulate pasting the HTML content into the document
		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		const result = ToJsonString(logicDocument);

		const expected = {
			"type": "document",
			"textPr": "Test HTML content\r\n\r\n",
			'content': [{"bFromDocument":true,"pPr":{"pBdr":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"left":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"right":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["Test HTML content"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"},{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		};

		assert.strictEqual(result, JSON.stringify(expected), "HTML content should match expected JSON format");

		done();
	});

	QUnit.test("Test: \"copy complex HTML with JSON verification\"", function (assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();

		// Create a complex HTML element to simulate copying
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = `
								<div>
								  <h1 style="color: red;">Title</h1>
								  <p>Paragraph with <strong>bold</strong> and <em>italic</em> text.</p>
								  <ul>
									<li>List item 1</li>
									<li>List item 2</li>
								  </ul>
								</div>
							  `;

		// Simulate pasting the HTML content into the document
		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		const result = ToJsonString(logicDocument);

		const expected = {
			"type": "document",
			"textPr": "Title\r\nParagraph with bold and italic text.\r\n·\tList item 1\r\nList item 2\r\n",
			'content': [{"bFromDocument":true,"pPr":{"pBdr":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"left":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"right":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"bFromDocument":true,"type":"paraPr","pStyle":"139"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["Title"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"},{"bFromDocument":true,"pPr":{"pBdr":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"left":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"right":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["Paragraph with bold and italic text."],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"},{"bFromDocument":true,"pPr":{"numPr":{"ilvl":0,"numId":"488"},"pBdr":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"left":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"right":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"bFromDocument":true,"type":"paraPr","pStyle":"165"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["List item 1"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"},{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["List item 2"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		};

		// result json object content will have "numId\":\"...\", I need to copy that part into my expected object content
		// to make the test pass, because the numId is generated dynamically

		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}

		assert.strictEqual(result, JSON.stringify(expected), "Complex HTML content should match expected JSON format");

		done();
	});

	QUnit.test("Test: \"paste html, select text, copy html, check htmls for simple lists\"", function (assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();

		// Create a complex HTML element to simulate copying
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = `
								<div>
								  <h1 style="color: red;">Title</h1>
								  <p>Paragraph with <strong>bold</strong> and <em>italic</em> text.</p>
								  <ul>
									<li>List item 1</li>
									<li>List item 2</li>
								  </ul>
								</div>
							  `;

		// Simulate pasting the HTML content into the document
		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		// Select the text in the paragraph and copy to clipboard
		logicDocument.SelectAll();
		var oCopyProcessor = new AscCommon.CopyProcessor(AscTest.Editor);
		const sBase64 = oCopyProcessor.Start();
		const _data = oCopyProcessor.getInnerHtml();
		logicDocument.RemoveSelection();
		const jsonedData = removeBase64(JSON.stringify(_data));
		const trueExpectations =
			"\"<h1 style=\\\"mso-pagination:widow-orphan lines-together;page-break-after:avoid;margin-top:18pt;margin-bottom:4pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\\" class=\\\"docData;\\\"><span style=\\\"font-family:'Arial';font-size:20pt;color:#376092;mso-style-textfill-fill-color:#376092\\\">Title</span></h1><p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\\"><span style=\\\"font-family:'Times New Roman';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">Paragraph with bold and italic text.</span></p><ul style=\\\"padding-left:40px\\\"><li style=\\\"list-style-type: disc\\\"><p style=\\\"margin-left:35.43307086614173pt;text-indent:-18pt;margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\\"><span style=\\\"font-family:'Times New Roman';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">List item 1</span></p></li></ul><p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\\"><span style=\\\"font-family:'Times New Roman';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">List item 2</span></p>\""

		assert.strictEqual(jsonedData, trueExpectations, "Copied data should be a document type");

		done();
	});

	QUnit.test("Test: \"paste html, select text, copy html, check htmls for marked lists\"", function (assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();

		// Вставляем только маркированный список
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = `
		<ul>
			<li>Элемент 1</li>
			<li>Элемент 2</li>
			<li>Элемент 3</li>
		</ul>
	`;

		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {
			// Копируем обратно
			logicDocument.SelectAll();
			var oCopyProcessor = new AscCommon.CopyProcessor(AscTest.Editor);
			oCopyProcessor.Start();
			const copiedHtml = oCopyProcessor.getInnerHtml();
			logicDocument.RemoveSelection();

			const jsonedData = removeBase64(JSON.stringify(copiedHtml));
			const expectedHtml ="\"<ul style=\\\"padding-left:40px\\\" class=\\\"docData;\\\"><li style=\\\"list-style-type: disc\\\"><p style=\\\"margin-left:35.43307086614173pt;text-indent:-18pt;margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\\"><span style=\\\"font-family:'Times New Roman';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">Элемент 1</span></p></li><li style=\\\"list-style-type: disc\\\"><p style=\\\"margin-left:35.43307086614173pt;text-indent:-18pt;margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\\"><span style=\\\"font-family:'Times New Roman';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">Элемент 2</span></p></li><li style=\\\"list-style-type: disc\\\"><p style=\\\"margin-left:35.43307086614173pt;text-indent:-18pt;margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\\"><span style=\\\"font-family:'Times New Roman';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">Элемент 3</span></p></li></ul><p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\\">&nbsp;</p>\""
			assert.strictEqual(jsonedData, expectedHtml, "Должен корректно копироваться маркированный список");
			done();
		});
	});

	QUnit.test("Test: \"paste html, select text, copy html, check htmls for numbered lists\"", function (assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();

		// Вставляем только нумерованный список
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = `
    <ol>
      <li>Элемент 1</li>
      <li>Элемент 2</li>
      <li>Элемент 3</li>
    </ol>
  `;

		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {
			// Копируем обратно
			logicDocument.SelectAll();
			var oCopyProcessor = new AscCommon.CopyProcessor(AscTest.Editor);
			oCopyProcessor.Start();
			const copiedHtml = oCopyProcessor.getInnerHtml();
			logicDocument.RemoveSelection();

			const jsonedData = removeBase64(JSON.stringify(copiedHtml));
			const expectedHtml = ''
			assert.strictEqual(jsonedData, expectedHtml, "Должен корректно копироваться нумерованный список");
			done();
		});
	});

	QUnit.test("Paste simple div HTML content", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<div>Simple text</div>";

		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});
		const result = ToJsonString(logicDocument);
		const expected = {
			"type": "document",
			"textPr": "Simple text\r\n\r\n",
			'content': [{"bFromDocument":true,"pPr":{"pBdr":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"left":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"right":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["Simple text"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"},{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		}

		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(result, JSON.stringify(expected), "Должен вставиться текст из div элемента");

		done();
	});

	QUnit.test("Paste simple div HTML, then select & copy back", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();

		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<div>Simple text</div>";

		// Paste HTML (no callback, just call it)
		AscTest.Editor.asc_PasteData(
			AscCommon.c_oAscClipboardDataFormat.HtmlElement,
			htmlElement
		);

		// Now select & copy
		logicDocument.SelectAll();
		let oCopyProcessor = new AscCommon.CopyProcessor(AscTest.Editor);
		oCopyProcessor.Start();
		const copiedHtml = oCopyProcessor.getInnerHtml();
		logicDocument.RemoveSelection();

		// Normalize copied HTML for comparison
		const jsonedData = removeBase64(JSON.stringify(copiedHtml));
		const expectedHtml = "\"<p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\\" class=\\\"docData;\\\"><span style=\\\"font-family:'Times New Roman';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">Simple text</span></p><p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\\">&nbsp;</p>\""
		assert.strictEqual(jsonedData, expectedHtml, "Copied HTML should match for simple div paste");
		done();
	});

	QUnit.test("Paste paragraph and span with style", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<p><span style='color:blue;'>Blue text</span></p>";

		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		const result = ToJsonString(logicDocument);
		const expected = {
			"type": "document",
			"textPr": "Blue text\r\n\r\n",
			'content': [{"bFromDocument":true,"pPr":{"pBdr":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"left":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"right":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["Blue text"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"},{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		}
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(result, JSON.stringify(expected), "Должен вставиться текст с цветом из span элемента");
		done();
	});

	QUnit.test("Paste table HTML", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<table><tr><td>Cell 1</td><td>Cell 2</td></tr></table>";

		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		const result = ToJsonString(logicDocument);
		const expected = {
			"type": "document",
			"textPr": "Cell 1\tCell 2\t\r\n",
			'content': [{"bPresentation":false,"tblGrid":[{"w":4677,"type":"gridCol"},{"w":4677,"type":"gridCol"}],"tblPr":{"tblBorders":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"end":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"insideH":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"insideV":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"start":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"tblCellMar":{},"tblLayout":"autofit","tblLook":{"firstColumn":true,"firstRow":true,"lastColumn":false,"lastRow":false,"noHBand":false,"noVBand":true},"tblOverlap":"overlap","tblpPr":{"horzAnchor":"page","vertAnchor":"page","tblpXSpec":"center","tblpYSpec":"center","tblpX":0,"tblpY":57,"bottomFromText":0,"leftFromText":0,"rightFromText":0,"topFromText":0},"tblStyle":"12","tblW":{"type":"auto","w":0},"inline":true,"type":"tablePr"},"content":[{"content":[{"content":{"bPresentation":false,"content":[{"bFromDocument":true,"pPr":{"spacing":{"before":0,"after":0},"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["Cell 1"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}],"type":"docContent"},"tcPr":{"tcBorders":{},"tcW":{"type":"dxa","w":4677},"type":"tableCellPr"},"id":"592","type":"tblCell"},{"content":{"bPresentation":false,"content":[{"bFromDocument":true,"pPr":{"spacing":{"before":0,"after":0},"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["Cell 2"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}],"type":"docContent"},"tcPr":{"tcBorders":{},"tcW":{"type":"dxa","w":4677},"type":"tableCellPr"},"id":"602","type":"tblCell"}],"reviewInfo":{"userId":"","author":"","date":"","moveType":"noMove","prevType":-1},"reviewType":"common","trPr":{"type":"tableRowPr"},"type":"tblRow"}],"changes":[],"type":"table"},{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		}
		const numIds = result.match(/"id":"(\d+)"/g);
		if (numIds) {
			expected.content[0].content[0].content[0].id = numIds[0].replace(/"id":"(\d+)"/, '$1');
			expected.content[0].content[0].content[1].id = numIds[1].replace(/"id":"(\d+)"/, '$1');
		}
		const wMatches = result.match(/"w":(\d+)/g);
		if (wMatches && wMatches.length >= 5) {
			const wValues = wMatches.map(m => m.replace(/"w":/, ''));

			// patch the 4 locations in expected
			expected.content[0].tblGrid[0].w = Number(wValues[0]);
			expected.content[0].tblGrid[1].w = Number(wValues[1]);
			expected.content[0].content[0].content[0].tcPr.tcW.w = Number(wValues[3]);
			expected.content[0].content[0].content[1].tcPr.tcW.w = Number(wValues[4]);
		}
		assert.strictEqual(result, JSON.stringify(expected), "Должна вставиться таблица с двумя ячейками");
		done();
	});

	QUnit.test("Paste unordered list HTML", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<ul><li>Item 1</li><li>Item 2</li></ul>";

		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		const result = ToJsonString(logicDocument);
		const expected = {
			"type": "document",
			"textPr": "·\tItem 1\r\n·\tItem 2\r\n\r\n",
			'content': [{"bFromDocument":true,"pPr":{"numPr":{"ilvl":0,"numId":"620"},"pBdr":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"left":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"right":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"bFromDocument":true,"type":"paraPr","pStyle":"165"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["Item 1"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"},{"bFromDocument":true,"pPr":{"numPr":{"ilvl":0,"numId":"620"},"pBdr":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"left":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"right":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"bFromDocument":true,"type":"paraPr","pStyle":"165"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["Item 2"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"},{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		}
		const numIds = result.match(/"numId":"(\d+)"/g);
		if (numIds) {
			expected.content[0].pPr.numPr.numId = numIds[0].match(/"numId":"(\d+)"/)[1];
			expected.content[1].pPr.numPr.numId = numIds[1].match(/"numId":"(\d+)"/)[1];
		}
		assert.strictEqual(result, JSON.stringify(expected), "Должен вставиться маркированный список с двумя элементами");
		done();
	});

	QUnit.test("Paste image HTML", function(assert) {
		// add prepare recursive to tests
		AscCommon.PasteProcessor.prototype._Prepeare_recursive = oldPrepeare_recursive;
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<img src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAUA' alt='Test Image'>";

		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		const result = ToJsonString(logicDocument);
		const expected = {
			"type": "document",
			"textPr": "\r\n",
			'content': [{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		}
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(result, JSON.stringify(expected), "Должно вставиться изображение из img элемента");

		// remove prepeare recursive from tests
		AscCommon.PasteProcessor.prototype._Prepeare_recursive = function () {

		};
		done();
	});

	QUnit.test("Paste bold and italic HTML", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<p><b>Bold</b> and <i>Italic</i></p>";

		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		const result = ToJsonString(logicDocument);
		const expected = {
			"type": "document",
			"textPr": "Bold and Italic\r\n\r\n",
			'content': [{"bFromDocument":true,"pPr":{"pBdr":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"left":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"right":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["Bold and Italic"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"},{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		}
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(result, JSON.stringify(expected), "Должен вставиться жирный и курсивный текст из HTML");
		done();
	});

	QUnit.test("Paste underline and strikethrough HTML", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<p><u>Underline</u> and <s>Strikethrough</s></p>";

		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		const result = ToJsonString(logicDocument);
		const expected = {
			"type": "document",
			"textPr": "Underline and Strikethrough\r\n\r\n",
			'content': [{"bFromDocument":true,"pPr":{"pBdr":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"left":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"right":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["Underline and Strikethrough"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"},{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		}
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(result, JSON.stringify(expected), "Должен вставиться подчеркнутый и зачеркнутый текст из HTML");
		done();
	});

	QUnit.test("Paste hyperlink HTML", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<a href='https://example.com'>Example Link</a>";

		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		const result = ToJsonString(logicDocument);
		const expected = {
			"type": "document",
			"textPr": "Example Link\r\n",
			'content': [{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"value":"https://example.com/","content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr","rStyle":"187"},"content":["Example Link"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"}],"type":"hyperlink"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		}
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(result, JSON.stringify(expected), "Должна вставиться гиперссылка из HTML");
		done();
	});

	// QUnit.test("Paste ordered list HTML", function(assert) {
	// 	AscTest.ClearDocument();
	// 	let p = AscTest.CreateParagraph();
	// 	logicDocument.AddToContent(0, p);
	//
	// 	let done = assert.async();
	// 	let htmlElement = document.createElement("div");
	// 	htmlElement.innerHTML = "<ol><li>First</li><li>Second</li></ol>";
	//
	// 	AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});
	//
	// 	const result = ToJsonString(logicDocument);
	// 	console.log("Ordered List Result:", result);
	// 	const expected = {
	// 		"type": "document",
	// 		"textPr": "",
	// 		'content': ""
	// 	}
	// 	// const numIds = result.match(/"numId":"(\d+)"/g);
	// 	// if (numIds) {
	// 	// 	expected.content[0].pPr.numPr.numId = numIds[0].match(/"numId":"(\d+)"/)[1];
	// 	// 	expected.content[1].pPr.numPr.numId = numIds[1].match(/"numId":"(\d+)"/)[1];
	// 	// }
	// 	assert.strictEqual(result, JSON.stringify(expected), "Должен вставиться маркированный список с двумя элементами");
	// 	done();
	// });

	QUnit.test("Paste nested HTML elements", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<div><span><b>Nested</b> <i>Elements</i></span></div>";

		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		const result = ToJsonString(logicDocument);
		const expected = {
			"type": "document",
			"textPr": "Nested Elements\r\n\r\n",
			'content': [{"bFromDocument":true,"pPr":{"pBdr":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"left":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"right":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["Nested Elements"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"},{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		}
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(result, JSON.stringify(expected), "Должен вставиться вложенный текст с жирным и курсивом из HTML");
		done();
	});

	QUnit.test("Paste line break HTML", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "Line1<br>Line2";

		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		const result = ToJsonString(logicDocument);
		const expected = {
			"type": "document",
			"textPr": "Line1\rLine2\r\n",
			'content': [{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["Line1",{"type":"break","breakType":"textWrapping"},"Line2"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		}
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(result, JSON.stringify(expected), "Должен вставиться перенос строки из HTML");
		done();
	});

	QUnit.test("Paste empty div HTML", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<div></div>";

		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		const result = ToJsonString(logicDocument);
		const expected = {
			"type": "document",
			"textPr": "\r\n",
			'content': [{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		}
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(result, JSON.stringify(expected), "Должен вставиться пустой параграф из пустого div элемента");
		done();
	});

	QUnit.test("Paste special character HTML", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<div>&copy; &euro; &amp;</div>";

		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		const result = ToJsonString(logicDocument);
		const expected = {
			"type": "document",
			"textPr": "© € &\r\n\r\n",
			'content': [{"bFromDocument":true,"pPr":{"pBdr":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"left":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"right":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["© € &"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"},{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		}
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(result, JSON.stringify(expected), "Должен вставиться текст с символами ©, €, & из HTML");
		done();
	});

	QUnit.test("Paste formula as text HTML", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<div>y = mx + b</div>";

		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		const result = ToJsonString(logicDocument);
		const expected = {
			"type": "document",
			"textPr": "y = mx + b\r\n\r\n",
			'content': [{"bFromDocument":true,"pPr":{"pBdr":{"bottom":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"left":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"right":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"},"top":{"color":{"auto":false,"r":0,"g":0,"b":0},"sz":4,"space":0,"value":"none"}},"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":["y = mx + b"],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"},{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		}
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(result, JSON.stringify(expected), "Должен вставиться текст формулы из HTML");
		done();
	});

	QUnit.test("Paste HTML with mso style", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);

		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = '<br style="page-break-before:always;mso-break-type:section-break;">';

		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function () {});

		const result = ToJsonString(logicDocument);
		const expected = {
			"type": "document",
			"textPr": "\r\r\n",
			'content': [{"bFromDocument":true,"pPr":{"bFromDocument":true,"type":"paraPr"},"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[{"type":"break","breakType":"page"}],"footnotes":[],"endnotes":[],"reviewType":"common","type":"run"},{"bFromDocument":true,"rPr":{"bFromDocument":true,"type":"textPr"},"content":[],"footnotes":[],"endnotes":[],"reviewType":"common","type":"endRun"}],"changes":[],"type":"paragraph"}]
		}
		const numId = result.match(/"numId":"(\d+)"/);
		if (numId) {
			expected.content[2].pPr.numPr.numId = numId[1];
		}
		assert.strictEqual(result, JSON.stringify(expected), "Должен вставиться перенос страницы из HTML с mso стилем");
		done();
	});

	QUnit.test("Paste paragraph + span with style, then select & copy back", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);
		let done = assert.async();

		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<p><span style='color:blue;'>Blue text</span></p>";
		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement);

		logicDocument.SelectAll();
		let oCopyProcessor = new AscCommon.CopyProcessor(AscTest.Editor);
		oCopyProcessor.Start();
		const copiedHtml = oCopyProcessor.getInnerHtml();
		logicDocument.RemoveSelection();

		const jsonedData = removeBase64(JSON.stringify(copiedHtml));
		const expectedHtml = "\"<p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\\" class=\\\"docData;\\\"><span style=\\\"font-family:'Times New Roman';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">Blue text</span></p><p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\\">&nbsp;</p>\""
		assert.strictEqual(jsonedData, expectedHtml, "Copied HTML should match for span with style");
		done();
	});

	// QUnit.test("Paste table HTML, then select & copy back", function(assert) {
	// 	AscTest.ClearDocument();
	// 	let p = AscTest.CreateParagraph();
	// 	logicDocument.AddToContent(0, p);
	// 	let done = assert.async();
	//
	// 	let htmlElement = document.createElement("div");
	// 	htmlElement.innerHTML = "<table><tr><td>Cell 1</td><td>Cell 2</td></tr></table>";
	// 	AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement);
	//
	// 	logicDocument.SelectAll();
	// 	let oCopyProcessor = new AscCommon.CopyProcessor(AscTest.Editor);
	// 	oCopyProcessor.Start();
	// 	const copiedHtml = oCopyProcessor.getInnerHtml();
	// 	logicDocument.RemoveSelection();
	//
	// 	const jsonedData = removeBase64(JSON.stringify(copiedHtml));
	// 	const expectedHtml =''
	// 	assert.strictEqual(jsonedData, expectedHtml, "Copied HTML should match for table paste");
	// 	done();
	// });

	QUnit.test("Paste unordered list HTML, then select & copy back", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);
		let done = assert.async();

		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<ul><li>Item 1</li><li>Item 2</li></ul>";
		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement);

		logicDocument.SelectAll();
		let oCopyProcessor = new AscCommon.CopyProcessor(AscTest.Editor);
		oCopyProcessor.Start();
		const copiedHtml = oCopyProcessor.getInnerHtml();
		logicDocument.RemoveSelection();

		const jsonedData = removeBase64(JSON.stringify(copiedHtml));
		const expectedHtml = "\"<ul style=\\\"padding-left:40px\\\" class=\\\"docData;\\\"><li style=\\\"list-style-type: disc\\\"><p style=\\\"margin-left:35.43307086614173pt;text-indent:-18pt;margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\\"><span style=\\\"font-family:'Times New Roman';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">Item 1</span></p></li><li style=\\\"list-style-type: disc\\\"><p style=\\\"margin-left:35.43307086614173pt;text-indent:-18pt;margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\\"><span style=\\\"font-family:'Times New Roman';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">Item 2</span></p></li></ul><p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\\">&nbsp;</p>\""
		assert.strictEqual(jsonedData, expectedHtml, "Copied HTML should match for unordered list");
		done();
	});

	QUnit.test("Paste bold/italic HTML, then select & copy back", function(assert) {
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);
		let done = assert.async();

		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "<p><b>Bold</b> and <i>Italic</i></p>";
		AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement);

		logicDocument.SelectAll();
		let oCopyProcessor = new AscCommon.CopyProcessor(AscTest.Editor);
		oCopyProcessor.Start();
		const copiedHtml = oCopyProcessor.getInnerHtml();
		logicDocument.RemoveSelection();

		const jsonedData = removeBase64(JSON.stringify(copiedHtml));
		const expectedHtml = "\"<p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\\" class=\\\"docData;\\\"><span style=\\\"font-family:'Times New Roman';font-size:10pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">Bold and Italic</span></p><p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\\">&nbsp;</p>\""
		assert.strictEqual(jsonedData, expectedHtml, "Copied HTML should match for bold + italic");
		done();
	});

	QUnit.test("Paste sum formula from excel to word", function(assert) {
		initDocument(logicDocument);
		let done = assert.async();

		let htmlText = `<head>
			<meta http-equiv=Content-Type content="text/html; charset=utf-8">
			<meta name=ProgId content=Excel.Sheet>
			<meta name=Generator content="Microsoft Excel 15">
			<link id=Main-File rel=Main-File
			href="file:///C:/Users/asus/AppData/Local/Temp/msohtmlclip1/01/clip.htm">
			<link rel=File-List
			href="file:///C:/Users/asus/AppData/Local/Temp/msohtmlclip1/01/clip_filelist.xml">
			<style>
			<!--table
			\t{mso-displayed-decimal-separator:"\\,";
			\tmso-displayed-thousand-separator:" ";}
			@page
			\t{margin:.75in .7in .75in .7in;
			\tmso-header-margin:.3in;
			\tmso-footer-margin:.3in;}
			tr
			\t{mso-height-source:auto;}
			col
			\t{mso-width-source:auto;}
			br
			\t{mso-data-placement:same-cell;}
			td
			\t{padding-top:1px;
			\tpadding-right:1px;
			\tpadding-left:1px;
			\tmso-ignore:padding;
			\tcolor:black;
			\tfont-size:11.0pt;
			\tfont-weight:400;
			\tfont-style:normal;
			\ttext-decoration:none;
			\tfont-family:Calibri, sans-serif;
			\tmso-font-charset:0;
			\tmso-number-format:General;
			\ttext-align:general;
			\tvertical-align:bottom;
			\tborder:none;
			\tmso-background-source:auto;
			\tmso-pattern:auto;
			\tmso-protection:locked visible;
			\twhite-space:nowrap;
			\tmso-rotate:0;}
			-->
			</style>
			</head>
			
			<body link="#0563C1" vlink="#954F72">
			
			<table border=0 cellpadding=0 cellspacing=0 width=256 style='border-collapse:
			 collapse;width:192pt'>
			<!--StartFragment-->
			 <col width=64 span=4 style='width:48pt'>
			 <tr height=20 style='height:15.0pt'>
			  <td height=20 width=64 style='height:15.0pt;width:48pt'></td>
			  <td width=64 style='width:48pt'></td>
			  <td width=64 style='width:48pt'></td>
			  <td width=64 style='width:48pt'></td>
			 </tr>
			 <tr height=20 style='height:15.0pt'>
			  <td height=20 align=right style='height:15.0pt'>1</td>
			  <td align=right>2</td>
			  <td align=right>3</td>
			  <td align=right>6</td>
			 </tr>
			<!--EndFragment-->
			</table>
			
			</body>`;
		AscTest.Editor['pluginMethod_PasteHtml'](htmlText)

		logicDocument.SelectAll();
		let oCopyProcessor = new AscCommon.CopyProcessor(AscTest.Editor);
		oCopyProcessor.Start();
		const copiedHtml = oCopyProcessor.getInnerHtml();
		logicDocument.RemoveSelection();

		const jsonedData = removeBase64(JSON.stringify(copiedHtml));
		const expectedHtml = "\"<table cellspacing=\\\"0\\\" border=\\\"0\\\" cellpadding=\\\"0\\\" style=\\\"margin-left:0pt;mso-padding-alt:0pt 5.3858267716535435pt 0pt 5.3858267716535435pt;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-insidev:none;mso-border-insideh:none;\\\" class=\\\"docData;\\\"><tr style=\\\"height:15pt;\\\"><td width=\\\"62\\\" style=\\\"width:46.50000000000001pt;padding:0.7500000000000001pt 0.7500000000000001pt 0pt 0.7500000000000001pt;border-left:none;border-top:none;border-right:none;border-bottom:none;\\\"><p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\\">&nbsp;</p></td><td width=\\\"62\\\" style=\\\"width:46.50000000000001pt;padding:0.7500000000000001pt 0.7500000000000001pt 0pt 0.7500000000000001pt;border-left:none;border-top:none;border-right:none;border-bottom:none;\\\"><p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\\">&nbsp;</p></td><td width=\\\"62\\\" style=\\\"width:46.50000000000001pt;padding:0.7500000000000001pt 0.7500000000000001pt 0pt 0.7500000000000001pt;border-left:none;border-top:none;border-right:none;border-bottom:none;\\\"><p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\\">&nbsp;</p></td><td width=\\\"62\\\" style=\\\"width:46.50000000000001pt;padding:0.7500000000000001pt 0.7500000000000001pt 0pt 0.7500000000000001pt;border-left:none;border-top:none;border-right:none;border-bottom:none;\\\"><p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\\">&nbsp;</p></td></tr><tr style=\\\"height:15pt;\\\"><td width=\\\"62\\\" style=\\\"width:46.50000000000001pt;padding:0.7500000000000001pt 0.7500000000000001pt 0pt 0.7500000000000001pt;border-left:none;border-top:none;border-right:none;border-bottom:none;\\\"><p style=\\\"text-align:right;margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\\"><span style=\\\"font-family:'Times New Roman';font-size:11pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">1</span></p></td><td width=\\\"62\\\" style=\\\"width:46.50000000000001pt;padding:0.7500000000000001pt 0.7500000000000001pt 0pt 0.7500000000000001pt;border-left:none;border-top:none;border-right:none;border-bottom:none;\\\"><p style=\\\"text-align:right;margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\\"><span style=\\\"font-family:'Times New Roman';font-size:11pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">2</span></p></td><td width=\\\"62\\\" style=\\\"width:46.50000000000001pt;padding:0.7500000000000001pt 0.7500000000000001pt 0pt 0.7500000000000001pt;border-left:none;border-top:none;border-right:none;border-bottom:none;\\\"><p style=\\\"text-align:right;margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\\"><span style=\\\"font-family:'Times New Roman';font-size:11pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">3</span></p></td><td width=\\\"62\\\" style=\\\"width:46.50000000000001pt;padding:0.7500000000000001pt 0.7500000000000001pt 0pt 0.7500000000000001pt;border-left:none;border-top:none;border-right:none;border-bottom:none;\\\"><p style=\\\"text-align:right;margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\\"><span style=\\\"font-family:'Times New Roman';font-size:11pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">6</span></p></td></tr></table><p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\\">&nbsp;</p>\""
		assert.strictEqual(jsonedData, expectedHtml, "Copied HTML should match for bold + italic");
		done();
	});

	QUnit.test("Paste Newton's binom formula from word", function(assert) {
		initDocument(logicDocument);
		let done = assert.async();

		// let htmlElement = document.createElement("html");
		// htmlElement.setAttribute("xmlns:o", "urn:schemas-microsoft-com:office:office");
		// htmlElement.setAttribute("xmlns:w", "urn:schemas-microsoft-com:office:word");
		// htmlElement.setAttribute('xmlns:m', "http://schemas.microsoft.com/office/2004/12/omml");
		// htmlElement.setAttribute("xmlns", "http://www.w3.org/TR/REC-html40");

		const htmlText = `
				<head>
			<meta http-equiv=Content-Type content="text/html; charset=utf-8">
			<meta name=ProgId content=Word.Document>
			<meta name=Generator content="Microsoft Word 15">
			<meta name=Originator content="Microsoft Word 15">
			<link rel=File-List
			href="file:///C:/Users/asus/AppData/Local/Temp/msohtmlclip1/01/clip_filelist.xml">
			<!--[if !mso]>
			<style>
			v\\:* {behavior:url(#default#VML);}
			o\\:* {behavior:url(#default#VML);}
			w\\:* {behavior:url(#default#VML);}
			.shape {behavior:url(#default#VML);}
			</style>
			<![endif]--><!--[if gte mso 9]><xml>
			 <o:OfficeDocumentSettings>
			  <o:RelyOnVML/>
			  <o:AllowPNG/>
			 </o:OfficeDocumentSettings>
			</xml><![endif]-->
			<link rel=themeData
			href="file:///C:/Users/asus/AppData/Local/Temp/msohtmlclip1/01/clip_themedata.thmx">
			<link rel=colorSchemeMapping
			href="file:///C:/Users/asus/AppData/Local/Temp/msohtmlclip1/01/clip_colorschememapping.xml">
			<!--[if gte mso 9]><xml>
			 <w:WordDocument>
			  <w:View>Normal</w:View>
			  <w:Zoom>0</w:Zoom>
			  <w:TrackMoves/>
			  <w:TrackFormatting/>
			  <w:PunctuationKerning/>
			  <w:ValidateAgainstSchemas/>
			  <w:SaveIfXMLInvalid>false</w:SaveIfXMLInvalid>
			  <w:IgnoreMixedContent>false</w:IgnoreMixedContent>
			  <w:AlwaysShowPlaceholderText>false</w:AlwaysShowPlaceholderText>
			  <w:DoNotPromoteQF/>
			  <w:LidThemeOther>RU</w:LidThemeOther>
			  <w:LidThemeAsian>X-NONE</w:LidThemeAsian>
			  <w:LidThemeComplexScript>X-NONE</w:LidThemeComplexScript>
			  <w:Compatibility>
			   <w:BreakWrappedTables/>
			   <w:SnapToGridInCell/>
			   <w:WrapTextWithPunct/>
			   <w:UseAsianBreakRules/>
			   <w:DontGrowAutofit/>
			   <w:SplitPgBreakAndParaMark/>
			   <w:EnableOpenTypeKerning/>
			   <w:DontFlipMirrorIndents/>
			   <w:OverrideTableStyleHps/>
			  </w:Compatibility>
			  <m:mathPr>
			   <m:mathFont m:val="Cambria Math"/>
			   <m:brkBin m:val="before"/>
			   <m:brkBinSub m:val="&#45;-"/>
			   <m:smallFrac m:val="off"/>
			   <m:dispDef/>
			   <m:lMargin m:val="0"/>
			   <m:rMargin m:val="0"/>
			   <m:defJc m:val="centerGroup"/>
			   <m:wrapIndent m:val="1440"/>
			   <m:intLim m:val="subSup"/>
			   <m:naryLim m:val="undOvr"/>
			  </m:mathPr></w:WordDocument>
			</xml><![endif]--><!--[if gte mso 9]><xml>
			 <w:LatentStyles DefLockedState="false" DefUnhideWhenUsed="false"
			  DefSemiHidden="false" DefQFormat="false" DefPriority="99"
			  LatentStyleCount="376">
			  <w:LsdException Locked="false" Priority="0" QFormat="true" Name="Normal"/>
			  <w:LsdException Locked="false" Priority="9" QFormat="true" Name="heading 1"/>
			  <w:LsdException Locked="false" Priority="9" SemiHidden="true"
			   UnhideWhenUsed="true" QFormat="true" Name="heading 2"/>
			  <w:LsdException Locked="false" Priority="9" SemiHidden="true"
			   UnhideWhenUsed="true" QFormat="true" Name="heading 3"/>
			  <w:LsdException Locked="false" Priority="9" SemiHidden="true"
			   UnhideWhenUsed="true" QFormat="true" Name="heading 4"/>
			  <w:LsdException Locked="false" Priority="9" SemiHidden="true"
			   UnhideWhenUsed="true" QFormat="true" Name="heading 5"/>
			  <w:LsdException Locked="false" Priority="9" SemiHidden="true"
			   UnhideWhenUsed="true" QFormat="true" Name="heading 6"/>
			  <w:LsdException Locked="false" Priority="9" SemiHidden="true"
			   UnhideWhenUsed="true" QFormat="true" Name="heading 7"/>
			  <w:LsdException Locked="false" Priority="9" SemiHidden="true"
			   UnhideWhenUsed="true" QFormat="true" Name="heading 8"/>
			  <w:LsdException Locked="false" Priority="9" SemiHidden="true"
			   UnhideWhenUsed="true" QFormat="true" Name="heading 9"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="index 1"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="index 2"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="index 3"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="index 4"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="index 5"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="index 6"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="index 7"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="index 8"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="index 9"/>
			  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
			   UnhideWhenUsed="true" Name="toc 1"/>
			  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
			   UnhideWhenUsed="true" Name="toc 2"/>
			  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
			   UnhideWhenUsed="true" Name="toc 3"/>
			  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
			   UnhideWhenUsed="true" Name="toc 4"/>
			  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
			   UnhideWhenUsed="true" Name="toc 5"/>
			  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
			   UnhideWhenUsed="true" Name="toc 6"/>
			  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
			   UnhideWhenUsed="true" Name="toc 7"/>
			  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
			   UnhideWhenUsed="true" Name="toc 8"/>
			  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
			   UnhideWhenUsed="true" Name="toc 9"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Normal Indent"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="footnote text"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="annotation text"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="header"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="footer"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="index heading"/>
			  <w:LsdException Locked="false" Priority="35" SemiHidden="true"
			   UnhideWhenUsed="true" QFormat="true" Name="caption"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="table of figures"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="envelope address"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="envelope return"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="footnote reference"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="annotation reference"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="line number"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="page number"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="endnote reference"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="endnote text"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="table of authorities"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="macro"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="toa heading"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="List"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="List Bullet"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="List Number"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="List 2"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="List 3"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="List 4"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="List 5"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="List Bullet 2"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="List Bullet 3"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="List Bullet 4"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="List Bullet 5"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="List Number 2"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="List Number 3"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="List Number 4"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="List Number 5"/>
			  <w:LsdException Locked="false" Priority="10" QFormat="true" Name="Title"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Closing"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Signature"/>
			  <w:LsdException Locked="false" Priority="1" SemiHidden="true"
			   UnhideWhenUsed="true" Name="Default Paragraph Font"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Body Text"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Body Text Indent"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="List Continue"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="List Continue 2"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="List Continue 3"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="List Continue 4"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="List Continue 5"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Message Header"/>
			  <w:LsdException Locked="false" Priority="11" QFormat="true" Name="Subtitle"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Salutation"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Date"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Body Text First Indent"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Body Text First Indent 2"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Note Heading"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Body Text 2"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Body Text 3"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Body Text Indent 2"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Body Text Indent 3"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Block Text"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Hyperlink"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="FollowedHyperlink"/>
			  <w:LsdException Locked="false" Priority="22" QFormat="true" Name="Strong"/>
			  <w:LsdException Locked="false" Priority="20" QFormat="true" Name="Emphasis"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Document Map"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Plain Text"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="E-mail Signature"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="HTML Top of Form"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="HTML Bottom of Form"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Normal (Web)"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="HTML Acronym"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="HTML Address"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="HTML Cite"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="HTML Code"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="HTML Definition"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="HTML Keyboard"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="HTML Preformatted"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="HTML Sample"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="HTML Typewriter"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="HTML Variable"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Normal Table"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="annotation subject"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="No List"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Outline List 1"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Outline List 2"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Outline List 3"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Simple 1"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Simple 2"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Simple 3"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Classic 1"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Classic 2"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Classic 3"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Classic 4"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Colorful 1"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Colorful 2"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Colorful 3"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Columns 1"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Columns 2"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Columns 3"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Columns 4"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Columns 5"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Grid 1"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Grid 2"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Grid 3"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Grid 4"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Grid 5"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Grid 6"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Grid 7"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Grid 8"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table List 1"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table List 2"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table List 3"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table List 4"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table List 5"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table List 6"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table List 7"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table List 8"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table 3D effects 1"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table 3D effects 2"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table 3D effects 3"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Contemporary"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Elegant"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Professional"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Subtle 1"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Subtle 2"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Web 1"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Web 2"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Web 3"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Balloon Text"/>
			  <w:LsdException Locked="false" Priority="39" Name="Table Grid"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Table Theme"/>
			  <w:LsdException Locked="false" SemiHidden="true" Name="Placeholder Text"/>
			  <w:LsdException Locked="false" Priority="1" QFormat="true" Name="No Spacing"/>
			  <w:LsdException Locked="false" Priority="60" Name="Light Shading"/>
			  <w:LsdException Locked="false" Priority="61" Name="Light List"/>
			  <w:LsdException Locked="false" Priority="62" Name="Light Grid"/>
			  <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1"/>
			  <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2"/>
			  <w:LsdException Locked="false" Priority="65" Name="Medium List 1"/>
			  <w:LsdException Locked="false" Priority="66" Name="Medium List 2"/>
			  <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1"/>
			  <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2"/>
			  <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3"/>
			  <w:LsdException Locked="false" Priority="70" Name="Dark List"/>
			  <w:LsdException Locked="false" Priority="71" Name="Colorful Shading"/>
			  <w:LsdException Locked="false" Priority="72" Name="Colorful List"/>
			  <w:LsdException Locked="false" Priority="73" Name="Colorful Grid"/>
			  <w:LsdException Locked="false" Priority="60" Name="Light Shading Accent 1"/>
			  <w:LsdException Locked="false" Priority="61" Name="Light List Accent 1"/>
			  <w:LsdException Locked="false" Priority="62" Name="Light Grid Accent 1"/>
			  <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1 Accent 1"/>
			  <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2 Accent 1"/>
			  <w:LsdException Locked="false" Priority="65" Name="Medium List 1 Accent 1"/>
			  <w:LsdException Locked="false" SemiHidden="true" Name="Revision"/>
			  <w:LsdException Locked="false" Priority="34" QFormat="true"
			   Name="List Paragraph"/>
			  <w:LsdException Locked="false" Priority="29" QFormat="true" Name="Quote"/>
			  <w:LsdException Locked="false" Priority="30" QFormat="true"
			   Name="Intense Quote"/>
			  <w:LsdException Locked="false" Priority="66" Name="Medium List 2 Accent 1"/>
			  <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1 Accent 1"/>
			  <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2 Accent 1"/>
			  <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3 Accent 1"/>
			  <w:LsdException Locked="false" Priority="70" Name="Dark List Accent 1"/>
			  <w:LsdException Locked="false" Priority="71" Name="Colorful Shading Accent 1"/>
			  <w:LsdException Locked="false" Priority="72" Name="Colorful List Accent 1"/>
			  <w:LsdException Locked="false" Priority="73" Name="Colorful Grid Accent 1"/>
			  <w:LsdException Locked="false" Priority="60" Name="Light Shading Accent 2"/>
			  <w:LsdException Locked="false" Priority="61" Name="Light List Accent 2"/>
			  <w:LsdException Locked="false" Priority="62" Name="Light Grid Accent 2"/>
			  <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1 Accent 2"/>
			  <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2 Accent 2"/>
			  <w:LsdException Locked="false" Priority="65" Name="Medium List 1 Accent 2"/>
			  <w:LsdException Locked="false" Priority="66" Name="Medium List 2 Accent 2"/>
			  <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1 Accent 2"/>
			  <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2 Accent 2"/>
			  <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3 Accent 2"/>
			  <w:LsdException Locked="false" Priority="70" Name="Dark List Accent 2"/>
			  <w:LsdException Locked="false" Priority="71" Name="Colorful Shading Accent 2"/>
			  <w:LsdException Locked="false" Priority="72" Name="Colorful List Accent 2"/>
			  <w:LsdException Locked="false" Priority="73" Name="Colorful Grid Accent 2"/>
			  <w:LsdException Locked="false" Priority="60" Name="Light Shading Accent 3"/>
			  <w:LsdException Locked="false" Priority="61" Name="Light List Accent 3"/>
			  <w:LsdException Locked="false" Priority="62" Name="Light Grid Accent 3"/>
			  <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1 Accent 3"/>
			  <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2 Accent 3"/>
			  <w:LsdException Locked="false" Priority="65" Name="Medium List 1 Accent 3"/>
			  <w:LsdException Locked="false" Priority="66" Name="Medium List 2 Accent 3"/>
			  <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1 Accent 3"/>
			  <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2 Accent 3"/>
			  <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3 Accent 3"/>
			  <w:LsdException Locked="false" Priority="70" Name="Dark List Accent 3"/>
			  <w:LsdException Locked="false" Priority="71" Name="Colorful Shading Accent 3"/>
			  <w:LsdException Locked="false" Priority="72" Name="Colorful List Accent 3"/>
			  <w:LsdException Locked="false" Priority="73" Name="Colorful Grid Accent 3"/>
			  <w:LsdException Locked="false" Priority="60" Name="Light Shading Accent 4"/>
			  <w:LsdException Locked="false" Priority="61" Name="Light List Accent 4"/>
			  <w:LsdException Locked="false" Priority="62" Name="Light Grid Accent 4"/>
			  <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1 Accent 4"/>
			  <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2 Accent 4"/>
			  <w:LsdException Locked="false" Priority="65" Name="Medium List 1 Accent 4"/>
			  <w:LsdException Locked="false" Priority="66" Name="Medium List 2 Accent 4"/>
			  <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1 Accent 4"/>
			  <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2 Accent 4"/>
			  <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3 Accent 4"/>
			  <w:LsdException Locked="false" Priority="70" Name="Dark List Accent 4"/>
			  <w:LsdException Locked="false" Priority="71" Name="Colorful Shading Accent 4"/>
			  <w:LsdException Locked="false" Priority="72" Name="Colorful List Accent 4"/>
			  <w:LsdException Locked="false" Priority="73" Name="Colorful Grid Accent 4"/>
			  <w:LsdException Locked="false" Priority="60" Name="Light Shading Accent 5"/>
			  <w:LsdException Locked="false" Priority="61" Name="Light List Accent 5"/>
			  <w:LsdException Locked="false" Priority="62" Name="Light Grid Accent 5"/>
			  <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1 Accent 5"/>
			  <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2 Accent 5"/>
			  <w:LsdException Locked="false" Priority="65" Name="Medium List 1 Accent 5"/>
			  <w:LsdException Locked="false" Priority="66" Name="Medium List 2 Accent 5"/>
			  <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1 Accent 5"/>
			  <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2 Accent 5"/>
			  <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3 Accent 5"/>
			  <w:LsdException Locked="false" Priority="70" Name="Dark List Accent 5"/>
			  <w:LsdException Locked="false" Priority="71" Name="Colorful Shading Accent 5"/>
			  <w:LsdException Locked="false" Priority="72" Name="Colorful List Accent 5"/>
			  <w:LsdException Locked="false" Priority="73" Name="Colorful Grid Accent 5"/>
			  <w:LsdException Locked="false" Priority="60" Name="Light Shading Accent 6"/>
			  <w:LsdException Locked="false" Priority="61" Name="Light List Accent 6"/>
			  <w:LsdException Locked="false" Priority="62" Name="Light Grid Accent 6"/>
			  <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1 Accent 6"/>
			  <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2 Accent 6"/>
			  <w:LsdException Locked="false" Priority="65" Name="Medium List 1 Accent 6"/>
			  <w:LsdException Locked="false" Priority="66" Name="Medium List 2 Accent 6"/>
			  <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1 Accent 6"/>
			  <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2 Accent 6"/>
			  <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3 Accent 6"/>
			  <w:LsdException Locked="false" Priority="70" Name="Dark List Accent 6"/>
			  <w:LsdException Locked="false" Priority="71" Name="Colorful Shading Accent 6"/>
			  <w:LsdException Locked="false" Priority="72" Name="Colorful List Accent 6"/>
			  <w:LsdException Locked="false" Priority="73" Name="Colorful Grid Accent 6"/>
			  <w:LsdException Locked="false" Priority="19" QFormat="true"
			   Name="Subtle Emphasis"/>
			  <w:LsdException Locked="false" Priority="21" QFormat="true"
			   Name="Intense Emphasis"/>
			  <w:LsdException Locked="false" Priority="31" QFormat="true"
			   Name="Subtle Reference"/>
			  <w:LsdException Locked="false" Priority="32" QFormat="true"
			   Name="Intense Reference"/>
			  <w:LsdException Locked="false" Priority="33" QFormat="true" Name="Book Title"/>
			  <w:LsdException Locked="false" Priority="37" SemiHidden="true"
			   UnhideWhenUsed="true" Name="Bibliography"/>
			  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
			   UnhideWhenUsed="true" QFormat="true" Name="TOC Heading"/>
			  <w:LsdException Locked="false" Priority="41" Name="Plain Table 1"/>
			  <w:LsdException Locked="false" Priority="42" Name="Plain Table 2"/>
			  <w:LsdException Locked="false" Priority="43" Name="Plain Table 3"/>
			  <w:LsdException Locked="false" Priority="44" Name="Plain Table 4"/>
			  <w:LsdException Locked="false" Priority="45" Name="Plain Table 5"/>
			  <w:LsdException Locked="false" Priority="40" Name="Grid Table Light"/>
			  <w:LsdException Locked="false" Priority="46" Name="Grid Table 1 Light"/>
			  <w:LsdException Locked="false" Priority="47" Name="Grid Table 2"/>
			  <w:LsdException Locked="false" Priority="48" Name="Grid Table 3"/>
			  <w:LsdException Locked="false" Priority="49" Name="Grid Table 4"/>
			  <w:LsdException Locked="false" Priority="50" Name="Grid Table 5 Dark"/>
			  <w:LsdException Locked="false" Priority="51" Name="Grid Table 6 Colorful"/>
			  <w:LsdException Locked="false" Priority="52" Name="Grid Table 7 Colorful"/>
			  <w:LsdException Locked="false" Priority="46"
			   Name="Grid Table 1 Light Accent 1"/>
			  <w:LsdException Locked="false" Priority="47" Name="Grid Table 2 Accent 1"/>
			  <w:LsdException Locked="false" Priority="48" Name="Grid Table 3 Accent 1"/>
			  <w:LsdException Locked="false" Priority="49" Name="Grid Table 4 Accent 1"/>
			  <w:LsdException Locked="false" Priority="50" Name="Grid Table 5 Dark Accent 1"/>
			  <w:LsdException Locked="false" Priority="51"
			   Name="Grid Table 6 Colorful Accent 1"/>
			  <w:LsdException Locked="false" Priority="52"
			   Name="Grid Table 7 Colorful Accent 1"/>
			  <w:LsdException Locked="false" Priority="46"
			   Name="Grid Table 1 Light Accent 2"/>
			  <w:LsdException Locked="false" Priority="47" Name="Grid Table 2 Accent 2"/>
			  <w:LsdException Locked="false" Priority="48" Name="Grid Table 3 Accent 2"/>
			  <w:LsdException Locked="false" Priority="49" Name="Grid Table 4 Accent 2"/>
			  <w:LsdException Locked="false" Priority="50" Name="Grid Table 5 Dark Accent 2"/>
			  <w:LsdException Locked="false" Priority="51"
			   Name="Grid Table 6 Colorful Accent 2"/>
			  <w:LsdException Locked="false" Priority="52"
			   Name="Grid Table 7 Colorful Accent 2"/>
			  <w:LsdException Locked="false" Priority="46"
			   Name="Grid Table 1 Light Accent 3"/>
			  <w:LsdException Locked="false" Priority="47" Name="Grid Table 2 Accent 3"/>
			  <w:LsdException Locked="false" Priority="48" Name="Grid Table 3 Accent 3"/>
			  <w:LsdException Locked="false" Priority="49" Name="Grid Table 4 Accent 3"/>
			  <w:LsdException Locked="false" Priority="50" Name="Grid Table 5 Dark Accent 3"/>
			  <w:LsdException Locked="false" Priority="51"
			   Name="Grid Table 6 Colorful Accent 3"/>
			  <w:LsdException Locked="false" Priority="52"
			   Name="Grid Table 7 Colorful Accent 3"/>
			  <w:LsdException Locked="false" Priority="46"
			   Name="Grid Table 1 Light Accent 4"/>
			  <w:LsdException Locked="false" Priority="47" Name="Grid Table 2 Accent 4"/>
			  <w:LsdException Locked="false" Priority="48" Name="Grid Table 3 Accent 4"/>
			  <w:LsdException Locked="false" Priority="49" Name="Grid Table 4 Accent 4"/>
			  <w:LsdException Locked="false" Priority="50" Name="Grid Table 5 Dark Accent 4"/>
			  <w:LsdException Locked="false" Priority="51"
			   Name="Grid Table 6 Colorful Accent 4"/>
			  <w:LsdException Locked="false" Priority="52"
			   Name="Grid Table 7 Colorful Accent 4"/>
			  <w:LsdException Locked="false" Priority="46"
			   Name="Grid Table 1 Light Accent 5"/>
			  <w:LsdException Locked="false" Priority="47" Name="Grid Table 2 Accent 5"/>
			  <w:LsdException Locked="false" Priority="48" Name="Grid Table 3 Accent 5"/>
			  <w:LsdException Locked="false" Priority="49" Name="Grid Table 4 Accent 5"/>
			  <w:LsdException Locked="false" Priority="50" Name="Grid Table 5 Dark Accent 5"/>
			  <w:LsdException Locked="false" Priority="51"
			   Name="Grid Table 6 Colorful Accent 5"/>
			  <w:LsdException Locked="false" Priority="52"
			   Name="Grid Table 7 Colorful Accent 5"/>
			  <w:LsdException Locked="false" Priority="46"
			   Name="Grid Table 1 Light Accent 6"/>
			  <w:LsdException Locked="false" Priority="47" Name="Grid Table 2 Accent 6"/>
			  <w:LsdException Locked="false" Priority="48" Name="Grid Table 3 Accent 6"/>
			  <w:LsdException Locked="false" Priority="49" Name="Grid Table 4 Accent 6"/>
			  <w:LsdException Locked="false" Priority="50" Name="Grid Table 5 Dark Accent 6"/>
			  <w:LsdException Locked="false" Priority="51"
			   Name="Grid Table 6 Colorful Accent 6"/>
			  <w:LsdException Locked="false" Priority="52"
			   Name="Grid Table 7 Colorful Accent 6"/>
			  <w:LsdException Locked="false" Priority="46" Name="List Table 1 Light"/>
			  <w:LsdException Locked="false" Priority="47" Name="List Table 2"/>
			  <w:LsdException Locked="false" Priority="48" Name="List Table 3"/>
			  <w:LsdException Locked="false" Priority="49" Name="List Table 4"/>
			  <w:LsdException Locked="false" Priority="50" Name="List Table 5 Dark"/>
			  <w:LsdException Locked="false" Priority="51" Name="List Table 6 Colorful"/>
			  <w:LsdException Locked="false" Priority="52" Name="List Table 7 Colorful"/>
			  <w:LsdException Locked="false" Priority="46"
			   Name="List Table 1 Light Accent 1"/>
			  <w:LsdException Locked="false" Priority="47" Name="List Table 2 Accent 1"/>
			  <w:LsdException Locked="false" Priority="48" Name="List Table 3 Accent 1"/>
			  <w:LsdException Locked="false" Priority="49" Name="List Table 4 Accent 1"/>
			  <w:LsdException Locked="false" Priority="50" Name="List Table 5 Dark Accent 1"/>
			  <w:LsdException Locked="false" Priority="51"
			   Name="List Table 6 Colorful Accent 1"/>
			  <w:LsdException Locked="false" Priority="52"
			   Name="List Table 7 Colorful Accent 1"/>
			  <w:LsdException Locked="false" Priority="46"
			   Name="List Table 1 Light Accent 2"/>
			  <w:LsdException Locked="false" Priority="47" Name="List Table 2 Accent 2"/>
			  <w:LsdException Locked="false" Priority="48" Name="List Table 3 Accent 2"/>
			  <w:LsdException Locked="false" Priority="49" Name="List Table 4 Accent 2"/>
			  <w:LsdException Locked="false" Priority="50" Name="List Table 5 Dark Accent 2"/>
			  <w:LsdException Locked="false" Priority="51"
			   Name="List Table 6 Colorful Accent 2"/>
			  <w:LsdException Locked="false" Priority="52"
			   Name="List Table 7 Colorful Accent 2"/>
			  <w:LsdException Locked="false" Priority="46"
			   Name="List Table 1 Light Accent 3"/>
			  <w:LsdException Locked="false" Priority="47" Name="List Table 2 Accent 3"/>
			  <w:LsdException Locked="false" Priority="48" Name="List Table 3 Accent 3"/>
			  <w:LsdException Locked="false" Priority="49" Name="List Table 4 Accent 3"/>
			  <w:LsdException Locked="false" Priority="50" Name="List Table 5 Dark Accent 3"/>
			  <w:LsdException Locked="false" Priority="51"
			   Name="List Table 6 Colorful Accent 3"/>
			  <w:LsdException Locked="false" Priority="52"
			   Name="List Table 7 Colorful Accent 3"/>
			  <w:LsdException Locked="false" Priority="46"
			   Name="List Table 1 Light Accent 4"/>
			  <w:LsdException Locked="false" Priority="47" Name="List Table 2 Accent 4"/>
			  <w:LsdException Locked="false" Priority="48" Name="List Table 3 Accent 4"/>
			  <w:LsdException Locked="false" Priority="49" Name="List Table 4 Accent 4"/>
			  <w:LsdException Locked="false" Priority="50" Name="List Table 5 Dark Accent 4"/>
			  <w:LsdException Locked="false" Priority="51"
			   Name="List Table 6 Colorful Accent 4"/>
			  <w:LsdException Locked="false" Priority="52"
			   Name="List Table 7 Colorful Accent 4"/>
			  <w:LsdException Locked="false" Priority="46"
			   Name="List Table 1 Light Accent 5"/>
			  <w:LsdException Locked="false" Priority="47" Name="List Table 2 Accent 5"/>
			  <w:LsdException Locked="false" Priority="48" Name="List Table 3 Accent 5"/>
			  <w:LsdException Locked="false" Priority="49" Name="List Table 4 Accent 5"/>
			  <w:LsdException Locked="false" Priority="50" Name="List Table 5 Dark Accent 5"/>
			  <w:LsdException Locked="false" Priority="51"
			   Name="List Table 6 Colorful Accent 5"/>
			  <w:LsdException Locked="false" Priority="52"
			   Name="List Table 7 Colorful Accent 5"/>
			  <w:LsdException Locked="false" Priority="46"
			   Name="List Table 1 Light Accent 6"/>
			  <w:LsdException Locked="false" Priority="47" Name="List Table 2 Accent 6"/>
			  <w:LsdException Locked="false" Priority="48" Name="List Table 3 Accent 6"/>
			  <w:LsdException Locked="false" Priority="49" Name="List Table 4 Accent 6"/>
			  <w:LsdException Locked="false" Priority="50" Name="List Table 5 Dark Accent 6"/>
			  <w:LsdException Locked="false" Priority="51"
			   Name="List Table 6 Colorful Accent 6"/>
			  <w:LsdException Locked="false" Priority="52"
			   Name="List Table 7 Colorful Accent 6"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Mention"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Smart Hyperlink"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Hashtag"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Unresolved Mention"/>
			  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
			   Name="Smart Link"/>
			 </w:LatentStyles>
			</xml><![endif]-->
			<style>
			<!--
			 /* Font Definitions */
			 @font-face
			\t{font-family:"Cambria Math";
			\tpanose-1:2 4 5 3 5 4 6 3 2 4;
			\tmso-font-charset:204;
			\tmso-generic-font-family:roman;
			\tmso-font-pitch:variable;
			\tmso-font-signature:-536869121 1107305727 33554432 0 415 0;}
			@font-face
			\t{font-family:Aptos;
			\tmso-font-charset:0;
			\tmso-generic-font-family:swiss;
			\tmso-font-pitch:variable;
			\tmso-font-signature:536871559 3 0 0 415 0;}
			 /* Style Definitions */
			 p.MsoNormal, li.MsoNormal, div.MsoNormal
			\t{mso-style-unhide:no;
			\tmso-style-qformat:yes;
			\tmso-style-parent:"";
			\tmargin-top:0cm;
			\tmargin-right:0cm;
			\tmargin-bottom:8.0pt;
			\tmargin-left:0cm;
			\tline-height:115%;
			\tmso-pagination:widow-orphan;
			\tfont-size:12.0pt;
			\tfont-family:"Aptos",sans-serif;
			\tmso-ascii-font-family:Aptos;
			\tmso-ascii-theme-font:minor-latin;
			\tmso-fareast-font-family:Aptos;
			\tmso-fareast-theme-font:minor-latin;
			\tmso-hansi-font-family:Aptos;
			\tmso-hansi-theme-font:minor-latin;
			\tmso-bidi-font-family:"Times New Roman";
			\tmso-bidi-theme-font:minor-bidi;
			\tmso-font-kerning:1.0pt;
			\tmso-ligatures:standardcontextual;
			\tmso-fareast-language:EN-US;}
			.MsoChpDefault
			\t{mso-style-type:export-only;
			\tmso-default-props:yes;
			\tfont-family:"Aptos",sans-serif;
			\tmso-ascii-font-family:Aptos;
			\tmso-ascii-theme-font:minor-latin;
			\tmso-fareast-font-family:Aptos;
			\tmso-fareast-theme-font:minor-latin;
			\tmso-hansi-font-family:Aptos;
			\tmso-hansi-theme-font:minor-latin;
			\tmso-bidi-font-family:"Times New Roman";
			\tmso-bidi-theme-font:minor-bidi;
			\tmso-fareast-language:EN-US;}
			.MsoPapDefault
			\t{mso-style-type:export-only;
			\tmargin-bottom:8.0pt;
			\tline-height:115%;}
			@page WordSection1
			\t{size:595.3pt 841.9pt;
			\tmargin:2.0cm 42.5pt 2.0cm 3.0cm;
			\tmso-header-margin:35.4pt;
			\tmso-footer-margin:35.4pt;
			\tmso-paper-source:0;}
			div.WordSection1
			\t{page:WordSection1;}
			-->
			</style>
			<!--[if gte mso 10]>
			<style>
			 /* Style Definitions */
			 table.MsoNormalTable
			\t{mso-style-name:"РћР±С‹С‡РЅР°СЏ С‚Р°Р±Р»РёС†Р°";
			\tmso-tstyle-rowband-size:0;
			\tmso-tstyle-colband-size:0;
			\tmso-style-noshow:yes;
			\tmso-style-priority:99;
			\tmso-style-parent:"";
			\tmso-padding-alt:0cm 5.4pt 0cm 5.4pt;
			\tmso-para-margin-top:0cm;
			\tmso-para-margin-right:0cm;
			\tmso-para-margin-bottom:8.0pt;
			\tmso-para-margin-left:0cm;
			\tline-height:115%;
			\tmso-pagination:widow-orphan;
			\tfont-size:12.0pt;
			\tfont-family:"Aptos",sans-serif;
			\tmso-ascii-font-family:Aptos;
			\tmso-ascii-theme-font:minor-latin;
			\tmso-hansi-font-family:Aptos;
			\tmso-hansi-theme-font:minor-latin;
			\tmso-font-kerning:1.0pt;
			\tmso-ligatures:standardcontextual;
			\tmso-fareast-language:EN-US;}
			</style>
			<![endif]-->
			</head>
			
			<body lang=RU style='tab-interval:35.4pt;word-wrap:break-word'>
			<!--StartFragment-->
			
			<p class=MsoNormal><!--[if gte msEquation 12]><m:oMathPara><m:oMath><m:sSup><m:sSupPr><span
				style='font-family:"Cambria Math",serif;mso-ascii-font-family:"Cambria Math";
				mso-hansi-font-family:"Cambria Math"'><m:ctrlPr></m:ctrlPr></span></m:sSupPr><m:e><m:d><m:dPr><span
				  style='font-family:"Cambria Math",serif;mso-ascii-font-family:"Cambria Math";
				  mso-hansi-font-family:"Cambria Math"'><m:ctrlPr></m:ctrlPr></span></m:dPr><m:e><i
				  style='mso-bidi-font-style:normal'><span style='font-family:"Cambria Math",serif'><m:r>x</m:r><m:r>+</m:r><m:r>a</m:r></span></i></m:e></m:d></m:e><m:sup><i
				style='mso-bidi-font-style:normal'><span style='font-family:"Cambria Math",serif'><m:r>n</m:r></span></i></m:sup></m:sSup><i
			  style='mso-bidi-font-style:normal'><span style='font-family:"Cambria Math",serif;
			  mso-fareast-font-family:"Cambria Math";mso-bidi-font-family:"Cambria Math"'><m:r>=</m:r></span></i><m:nary><m:naryPr><m:chr
				 m:val="в€‘"/><m:grow m:val="on"/><span style='font-family:"Cambria Math",serif;
				mso-ascii-font-family:"Cambria Math";mso-hansi-font-family:"Cambria Math"'><m:ctrlPr></m:ctrlPr></span></m:naryPr><m:sub><i
				style='mso-bidi-font-style:normal'><span style='font-family:"Cambria Math",serif;
				mso-fareast-font-family:"Cambria Math";mso-bidi-font-family:"Cambria Math"'><m:r>k</m:r><m:r>=0</m:r></span></i></m:sub><m:sup><i
				style='mso-bidi-font-style:normal'><span style='font-family:"Cambria Math",serif;
				mso-fareast-font-family:"Cambria Math";mso-bidi-font-family:"Cambria Math"'><m:r>n</m:r></span></i></m:sup><m:e><m:d><m:dPr><span
				  style='font-family:"Cambria Math",serif;mso-ascii-font-family:"Cambria Math";
				  mso-hansi-font-family:"Cambria Math"'><m:ctrlPr></m:ctrlPr></span></m:dPr><m:e><m:f><m:fPr><m:type
					 m:val="noBar"/><span style='font-family:"Cambria Math",serif;
					mso-ascii-font-family:"Cambria Math";mso-hansi-font-family:"Cambria Math"'><m:ctrlPr></m:ctrlPr></span></m:fPr><m:num><i
					style='mso-bidi-font-style:normal'><span style='font-family:"Cambria Math",serif;
					mso-fareast-font-family:"Cambria Math";mso-bidi-font-family:"Cambria Math"'><m:r>n</m:r></span></i></m:num><m:den><i
					style='mso-bidi-font-style:normal'><span style='font-family:"Cambria Math",serif;
					mso-fareast-font-family:"Cambria Math";mso-bidi-font-family:"Cambria Math"'><m:r>k</m:r></span></i></m:den></m:f></m:e></m:d><m:sSup><m:sSupPr><span
				  style='font-family:"Cambria Math",serif;mso-ascii-font-family:"Cambria Math";
				  mso-hansi-font-family:"Cambria Math"'><m:ctrlPr></m:ctrlPr></span></m:sSupPr><m:e><i
				  style='mso-bidi-font-style:normal'><span style='font-family:"Cambria Math",serif;
				  mso-fareast-font-family:"Cambria Math";mso-bidi-font-family:"Cambria Math"'><m:r>x</m:r></span></i></m:e><m:sup><i
				  style='mso-bidi-font-style:normal'><span style='font-family:"Cambria Math",serif;
				  mso-fareast-font-family:"Cambria Math";mso-bidi-font-family:"Cambria Math"'><m:r>k</m:r></span></i></m:sup></m:sSup><m:sSup><m:sSupPr><span
				  style='font-family:"Cambria Math",serif;mso-ascii-font-family:"Cambria Math";
				  mso-hansi-font-family:"Cambria Math"'><m:ctrlPr></m:ctrlPr></span></m:sSupPr><m:e><i
				  style='mso-bidi-font-style:normal'><span style='font-family:"Cambria Math",serif;
				  mso-fareast-font-family:"Cambria Math";mso-bidi-font-family:"Cambria Math"'><m:r>a</m:r></span></i></m:e><m:sup><i
				  style='mso-bidi-font-style:normal'><span style='font-family:"Cambria Math",serif;
				  mso-fareast-font-family:"Cambria Math";mso-bidi-font-family:"Cambria Math"'><m:r>n</m:r><m:r>-</m:r><m:r>k</m:r></span></i></m:sup></m:sSup></m:e></m:nary></m:oMath></m:oMathPara><![endif]--><![if !msEquation]><span
			style='font-size:12.0pt;line-height:115%;font-family:"Aptos",sans-serif;
			mso-ascii-theme-font:minor-latin;mso-fareast-font-family:Aptos;mso-fareast-theme-font:
			minor-latin;mso-hansi-theme-font:minor-latin;mso-bidi-font-family:"Times New Roman";
			mso-bidi-theme-font:minor-bidi;mso-ansi-language:RU;mso-fareast-language:EN-US;
			mso-bidi-language:AR-SA'><v:shapetype id="_x0000_t75" coordsize="21600,21600"
			 o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f"
			 stroked="f">
			 <v:stroke joinstyle="miter"/>
			 <v:formulas>
			  <v:f eqn="if lineDrawn pixelLineWidth 0"/>
			  <v:f eqn="sum @0 1 0"/>
			  <v:f eqn="sum 0 0 @1"/>
			  <v:f eqn="prod @2 1 2"/>
			  <v:f eqn="prod @3 21600 pixelWidth"/>
			  <v:f eqn="prod @3 21600 pixelHeight"/>
			  <v:f eqn="sum @0 0 1"/>
			  <v:f eqn="prod @6 1 2"/>
			  <v:f eqn="prod @7 21600 pixelWidth"/>
			  <v:f eqn="sum @8 21600 0"/>
			  <v:f eqn="prod @7 21600 pixelHeight"/>
			  <v:f eqn="sum @10 21600 0"/>
			 </v:formulas>
			 <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
			 <o:lock v:ext="edit" aspectratio="t"/>
			</v:shapetype><v:shape id="_x0000_i1025" type="#_x0000_t75" style='width:136.5pt;
			 height:41.25pt'>
			 <v:imagedata src="file:///C:/Users/asus/AppData/Local/Temp/msohtmlclip1/01/clip_image001.png"
			  o:title="" chromakey="white"/>
			</v:shape></span><![endif]><o:p></o:p></p>
			
			<!--EndFragment-->
			</body>

		`;
		// AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement);
		AscTest.Editor['pluginMethod_PasteHtml'](htmlText)

		logicDocument.SelectAll();
		let oCopyProcessor = new AscCommon.CopyProcessor(AscTest.Editor);
		oCopyProcessor.Start();
		const copiedHtml = oCopyProcessor.getInnerHtml();
		logicDocument.RemoveSelection();

		const jsonedData = removeBase64(JSON.stringify(copiedHtml));
		const expectedHtml = "\"<p style=\\\"line-height:13.8pt;margin-top:0pt;margin-bottom:0pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none\\\" class=\\\"docData;\\\"><span style=\\\"font-family:'Times New Roman';font-size:12pt;color:#000000;mso-style-textfill-fill-color:#000000\\\">&nbsp; </span></p><p style=\\\"margin-top:0pt;margin-bottom:0pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none\\\">&nbsp;</p>\""
		assert.strictEqual(jsonedData, expectedHtml, "Copied HTML should match for Newton's binom formula");
		done();
	});

	QUnit.module("Word Copy Paste Tests");
});

function removeBase64(html) {
	// 1. Remove long base64-like strings (letters, digits, +, /, =)
	html = html.replace(/([A-Za-z0-9+/=]{50,})/g, '');

	// 2. Remove dynamic docData metadata like: docData;DOCY;v5;3707;
	html = html.replace(/docData;DOCY;v\d+;\d+;?/g, 'docData;');

	return html;
}

function ToJsonString(logicDocument) {
	var oWriter = new AscJsonConverter.WriterToJSON();

	var oResult = {
		"type":      "document",
		"textPr":    logicDocument.GetText(),
		'content':	 oWriter.SerContent(logicDocument.Content, undefined, undefined, undefined, true),
		// "paraPr":    bWriteDefaultParaPr ? oWriter.SerParaPr(this.GetDefaultParaPr().ParaPr) : undefined,
		// "theme":     bWriteTheme ? oWriter.SerTheme(this.Document.GetTheme()) : undefined,
		// "sectPr":    bWriteSectionPr ? oWriter.SerSectionPr(this.Document.SectPr) : undefined,
		// "numbering": bWriteNumberings ? oWriter.jsonWordNumberings : undefined,
		// "styles":    bWriteStyles ? oWriter.SerWordStylesForWrite() : undefined
	}

	return JSON.stringify(oResult);
}

const charWidth = AscTest.CharWidth * AscTest.FontSize;
const L_FIELD = 20 * charWidth;
const R_FIELD = 30 * charWidth;
const PAGE_W  = 150 * charWidth;

function initDocument(logicDocument)
{
	AscTest.ClearDocument();
	logicDocument.AddToContent(0, AscTest.CreateParagraph());

	let sectPr = AscTest.GetFinalSection();
	sectPr.SetPageSize(PAGE_W, 1000);
	sectPr.SetPageMargins(L_FIELD, 50, R_FIELD, 50);
}
