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

	QUnit.test("Test: \"paste html, select text, copy html, check htmls\"", function (assert) {
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
		 <col width=64 span=4 style='width:48pt'>
		 <tr height=20 style='height:15.0pt'>
		<!--StartFragment-->
		  <td height=20 align=right width=64 style='height:15.0pt;width:48pt'>1</td>
		  <td align=right width=64 style='width:48pt'>2</td>
		  <td align=right width=64 style='width:48pt'>3</td>
		  <td align=right width=64 style='width:48pt'>6</td>
		<!--EndFragment-->
		 </tr>
		</table>

		</body>`;
		AscTest.Editor['pluginMethod_PasteHtml'](htmlText)

		logicDocument.SelectAll();
		let oCopyProcessor = new AscCommon.CopyProcessor(AscTest.Editor);
		oCopyProcessor.Start();
		const copiedHtml = oCopyProcessor.getInnerHtml();
		console.log(copiedHtml)
		logicDocument.RemoveSelection();

		const jsonedData = removeBase64(JSON.stringify(copiedHtml));
		const expectedHtml = ""
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
	<link rel=themeData
	href="file:///C:/Users/asus/AppData/Local/Temp/msohtmlclip1/01/clip_themedata.thmx">
	<link rel=colorSchemeMapping
	href="file:///C:/Users/asus/AppData/Local/Temp/msohtmlclip1/01/clip_colorschememapping.xml">
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
	
		</head>
		
		<body lang=RU style='tab-interval:35.4pt;word-wrap:break-word'>
		<!--StartFragment-->
		
		<p class=MsoNormal><span
			style='font-size:12.0pt;line-height:115%;font-family:"Aptos",sans-serif;
			mso-ascii-theme-font:minor-latin;mso-fareast-font-family:Aptos;mso-fareast-theme-font:
			minor-latin;mso-hansi-theme-font:minor-latin;mso-bidi-font-family:"Times New Roman";
			mso-bidi-theme-font:minor-bidi;mso-ansi-language:RU;mso-fareast-language:EN-US;
			mso-bidi-language:AR-SA'><img width=182 height=55
			src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAALYAAAA3CAMAAACWwkBLAAAAAXNSR0IArs4c6QAAALpQTFRFAAAAAAAAAAA6AABmADpmADqQAGa2OgAAOgA6OgBmOjo6OjpmOjqQOmaQOma2OpCQOpDbZgAAZgA6ZgBmZjoAZjo6ZjqQZmYAZmaQZpDbZraQZrbbZrb/kDoAkDo6kGY6kGZmkGaQkJC2kLbbkNvbkNv/tmYAtmY6tmaQtpA6tpBmtpCQtpC2ttv/tv/btv//25A625CQ27Zm27aQ27a229uQ2//b2////7Zm/7aQ/9uQ/9u2//+2///b3VjZygAAAAF0Uk5TAEDm2GYAAAAJcEhZcwAADsQAAA7EAZUrDhsAAAAZdEVYdFNvZnR3YXJlAE1pY3Jvc29mdCBPZmZpY2V/7TVxAAADn0lEQVRoQ+1Zi1baQBDdKBLbWgqWPiFqW4nVljZVC+mS/f/f6sw+ktlHMNgHyznMOeoBlszduzN3ZkfG9rZnYDsMrMb92/RouR3nj/d6//N4Ur3eOdis6LNy8Phtb+ub+ZTlk6+7Rnd1OmPXB7Ntkbb3u2dgz4BhoEg8m+4AOyJLDr8TnPDagi3Ok+lqnAxj2wpPkz7FxFMC+8f8Zp6/+LgsoPD/mkQFHcKEcimyGra4gg9EBrvK4Ue+isjyxCo0UDC1IVjGj2cAHd8TWVS4q1ESbP5KGfT4mz+R4Y87iMjc8FbQNLlIeXm0+ID9iqQ/HoPw9lVPUSzBl2mvhe78oUaGLOCvuh1W8XzekZrcVkH5rQC1fnRXL6l6Brw1C4qTrg0mf4r763B1CYR3MAFzNwn4wBZ6z1e9QOeHs7MimOUyq7pcXUpbBTH/qH4bZypLiRXD1cUn8lr7gmwBwwAyC1pkKAxbLe5ydfHCuwxFrfemF12uL7OgRYXCsBmWN9bl6uIWeVa4xEoJdI6gOr3Mhisas46veoE+p9UY2LhqGLFgQxNx8P4tPg1X06sLoNPmRimEtyVvQdjVSCpO7ZyfLEtVY7VL95pUL5D8MfF5kQ2vSJdAYfN0sKxG8mkbVIjSVhMv+/Bx6qmeczyG2mVATjBQDU2FaSU8BtGjbi02gM1TK6vXwEYUTk9IXD4A285qwrY8yZZ63B4kIrPDZj1sR1KoyzBs8wVxbqU6gS0Z1l49xQo/FPlzUnBtSjrOVSwGN2r86WMXZ5DDvAluCjudii9v+vwdDalWuPoDT6bXwXado8LULsOelBTnvRnE16CRHpqS10lvfis/7NxrViO3S1qj255zxhqXbfyEz31NlXyIaPw891Ta0A/C2KRq8Ai6PH+Tg+8ufwFq64OiYvRHnevd3+4AHe1T9BmlJT2QrjYd6f3Xy1ztU/4My3WhUDU3HivCvX6uYhr+rEZqhX4nDuTBFhXpltcbDAz+TN43wi3zljbhal8TFNCEYqxc3iihJaOILUGlbl3tI2pRAF6ctsm2T5zFNN8pndvvXUoKz/0cSvYCBzzQmna9mP6Po1DXJ8tUJmrDVrU4/CbnDfFY0xE2yC3YKIM87cXEdDzs7ZHEzYAzDoMU3Yl/q6ppl5nHiLNZXP1S25nb4zCs31F1Hm2w7XEYzDVwiBW/qWmXCRJku2XAFdVWnHEYBHZcHVMLWc04TC0od0RJojr7yMD8BlNbdQ3ULx22AAAAAElFTkSuQmCC"
			v:shapes="_x0000_i1025"></span><o:p></o:p></p>
	
		
		<!--EndFragment-->
		</body>

		`;
		// AscTest.Editor.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement);
		AscTest.Editor['pluginMethod_PasteHtml'](htmlText)

		logicDocument.SelectAll();
		let oCopyProcessor = new AscCommon.CopyProcessor(AscTest.Editor);
		oCopyProcessor.Start();
		console.log(oCopyProcessor);
		const copiedHtml = oCopyProcessor.getInnerHtml();
		logicDocument.RemoveSelection();

		const jsonedData = removeBase64(JSON.stringify(copiedHtml));
		const expectedHtml = ``
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
