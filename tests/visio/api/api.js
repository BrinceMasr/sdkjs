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


(function(window, document)
{
	/**
	 * Works like nullish coalescing operator but sets null if no default
	 * @param argumentValue - value from argument
	 * @param defaultValue - default value
	 */
	function getValueSafe(argumentValue, defaultValue) {
		if (typeof defaultValue === 'undefined') {
			defaultValue = null;
		}
		if (typeof argumentValue === 'undefined' || argumentValue === null) {
			return defaultValue;
		} else {
			return argumentValue;
		}
	}

	// function initProperty(thisArgument, argumentsObject, propName, defaultValue) {
	// 	thisArgument[propName] = getValueSafe(argumentsObject[propName], defaultValue);
	// }
	//
	// function Shape(argumentsObject) {
	// 	let initThisProperty = initProperty.bind(this, argumentsObject);
	// 	initThisProperty("width");
	// 	initThisProperty("height");
	// 	initThisProperty("locPinX", this.width / 2);
	// 	initThisProperty("locPinY", this.height / 2);
	// 	initThisProperty("angle", 0);
	// 	initThisProperty("fillForeground", 0);
	// 	return this;
	// }

	function createCell(n, v, u, e, f) {
		let cell = new AscVisio.Cell_Type();
		cell.n = getValueSafe(n, null);
		cell.v = getValueSafe(v, null);
		cell.u = getValueSafe(u, null);
		cell.e = getValueSafe(e, null);
		cell.f = getValueSafe(f, null);
		cell.refBy = [];
		cell.textContent = null;
		return cell;
	}

	function createShape(argumentsObject) {
		// add id when add to page
		// TODO accept objects like cell to be able to set F for example
		let width = getValueSafe(argumentsObject.width);
		let height = getValueSafe(argumentsObject.height);
		let locPinX = getValueSafe(argumentsObject.locPinX, width / 2);
		let locPinY = getValueSafe(argumentsObject.locPinY, height / 2);
		let angle = getValueSafe(argumentsObject.angle, 0);
		let fillForegroundColor = getValueSafe(argumentsObject.fillForegroundColor, 0);

		let pointsArray = getValueSafe(argumentsObject.points, []);

		let shape = new AscVisio.ShapeSheet_Type();
		shape.type = AscVisio.SHAPE_TYPES_SHAPE;
		shape.lineStyle = 0;
		shape.fillStyle = 0;
		shape.textStyle = 0;
		let elements = shape.elements;
		elements.push(createCell("Width", width));
		elements.push(createCell("Height", height));
		elements.push(createCell("LocPinX", locPinX));
		elements.push(createCell("LocPinY", locPinY));
		elements.push(createCell("Angle", angle));
		elements.push(createCell("FillForegnd", fillForegroundColor));

		let section = new AscVisio.Section_Type();
		section.n = "Geometry";
		section.ix = 0;

		for (let i = 0; i < pointsArray.length; i++) {
			let row;
			if (i === 0) {
				row = new AscVisio.Row_Type();
				row.t = "RelMoveTo";
				row.ix = 1;
			} else {
				row = new AscVisio.Row_Type();
				row.t = "RelLineTo";
				row.ix = i + 1;
			}
			row.cells.push(createCell("X", pointsArray[i].x));
			row.cells.push(createCell("Y", pointsArray[i].y));
			section.rows.push(row);
		}
		elements.push(section);
		return shape;
	}

	AscVisio.CPageContents.prototype.addShape = function (shape, pinX, pinY) {
		shape.iD = this.shapes.length + 1;
		shape.elements.push(createCell("PinX", pinX));
		shape.elements.push(createCell("PinY", pinY));
		this.shapes.push(shape);
	}

	AscVisio.CVisioDocument.prototype.getPages = function () {
		let pagesArray = this.pages.page;
		return pagesArray;
	}

	AscVisio.CVisioDocument.prototype.getPageContents = function (i) {
		let pageContents = this.pageContents[i];
		return pageContents;
	}

	window['AscVisio'].createShape = createShape;
})(window, window.document);
