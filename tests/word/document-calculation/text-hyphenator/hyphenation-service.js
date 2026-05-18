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

const AscHyphenation = {};

(function(window)
{
	// These are test functions that should be replaced with regular ones
	let BUFFER_STRING = "";
	
	AscHyphenation.addCodePoint = function(codePoint)
	{
		BUFFER_STRING += String.fromCodePoint(codePoint);
	};
	AscHyphenation.setLang = function(langCode)
	{
		return true;
	};
	AscHyphenation.hyphenate = function()
	{
		let checkString = BUFFER_STRING.toLowerCase();
		
		if ("abcd" === checkString)
			return [2];
		else if ("abb" === checkString)
			return [1];
		else if ("abbb" === checkString)
			return [1];
		else if ("aaabbb" === checkString)
			return [3];
		else if ("aabbb" === checkString)
			return [2];
		else if ("aabbbb" === checkString)
			return [2];
		else if ("abbb" === checkString)
			return [1];
		else if ("abbbb" === checkString)
			return [1];
		else if ("aaaaabbb" === checkString)
			return [5];
		else if ("testtest" === checkString)
			return [4];
		else if ("aabbbcccdddd" === checkString)
			return [2, 5, 8];
		else if ("ccdddd" === checkString)
			return [3];
		else if ("zz½www" === checkString)
			return [2];
		
		return [];
	};
	AscHyphenation.clear = function()
	{
		BUFFER_STRING = "";
	};
	
})(window);
