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

var AscTest = AscTest || {};

(function(window)
{
	const Letter = {
		f : 102,
		i : 105,

		x : 120,
		y : 121,
		z : 122
	};

	AscCommon.g_oTableId = {
		map : {},
		GetId : function()
		{
			return "-1";
		},
		Get_Id : function()
		{
			return this.GetId();
		},
		Add : function(c, id)
		{
			this.map[id] = c;
		},
		Get_ById : function(id)
		{
			if (!this.map[id])
				return null;

			return this.map[id];
		},
		GetById : function(id)
		{
			return this.Get_ById(id);
		},
		TurnOff : function(){},
		TurnOn : function(){},
		IsOn : function()
		{
			return true
		}
	};

	AscCommon.g_oIdCounter.m_bLoad = false;
	AscCommon.g_oIdCounter.m_bRead = false;

	function AddTextToInlineSdt(control, text)
	{
		AscWord.TextToRunElements(text, function(runElement)
		{
			control.Add(runElement);
		});
	}

	function GetBinaryWriter()
	{
		return new AscCommon.CMemory();
	}
	function GetBinaryReader(binaryWriter)
	{
		return new AscCommon.FT_Stream2(binaryWriter.GetData(), binaryWriter.GetCurPosition());
	}

	//--------------------------------------------------------export----------------------------------------------------
	AscTest.Letter = Letter;

	AscTest.AddTextToInlineSdt = AddTextToInlineSdt;
	AscTest.GetBinaryWriter    = GetBinaryWriter;
	AscTest.GetBinaryReader    = GetBinaryReader;

})(window);

if (QUnit && !QUnit.assert.close)
{
	if (!QUnit.assert.close)
	{
		QUnit.assert.close = function(number, expected, maxDifference, message)
		{
			if (undefined === maxDifference || null === maxDifference || 0 === maxDifference)
				maxDifference = 0.00001;
			
			QUnit.assert.pushResult({
				result   : Math.abs(number - expected) < maxDifference,
				actual   : number,
				expected : expected,
				message  : message
			});
		}
	}
	
	if (!QUnit.assert.equalRgb)
	{
		QUnit.assert.equalRgb = function(actualRgb, expectedRgb, message)
		{
			QUnit.assert.strictEqual(actualRgb.r, expectedRgb.r, (message ? message : "") + " check r component");
			QUnit.assert.strictEqual(actualRgb.g, expectedRgb.g, (message ? message : "") + " check g component");
			QUnit.assert.strictEqual(actualRgb.b, expectedRgb.b, (message ? message : "") + " check b component");
		};
	}
	
	if (!QUnit.assert.equalRgba)
	{
		QUnit.assert.equalRgba = function(actualRgba, expectedRgba, message)
		{
			QUnit.assert.strictEqual(actualRgba.r, expectedRgba.r, (message ? message : "") + " check r component");
			QUnit.assert.strictEqual(actualRgba.g, expectedRgba.g, (message ? message : "") + " check g component");
			QUnit.assert.strictEqual(actualRgba.b, expectedRgba.b, (message ? message : "") + " check b component");
			QUnit.assert.strictEqual(actualRgba.a, expectedRgba.a, (message ? message : "") + " check a component");
		};
	}
}

