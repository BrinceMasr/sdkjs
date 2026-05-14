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

$(function () {

	QUnit.module("Test api for the all editors");


	QUnit.test("Test asc_getUrlType", function (assert)
	{
		const editor = new AscCommon.baseEditorsApi({});

		let test = [
			["http://foo.com/blah_blah", AscCommon.c_oAscUrlType.Http],
			["http://foo.com/blah_blah_(wikipedia)_(again)", AscCommon.c_oAscUrlType.Http],
			["https://www.example.com/foo/?bar=baz&inga=42&quux", AscCommon.c_oAscUrlType.Http],
			["http://userid:password@example.com:8080", AscCommon.c_oAscUrlType.Http],
			["http://userid@example.com:8080/", AscCommon.c_oAscUrlType.Http],
			["http://142.42.1.1", AscCommon.c_oAscUrlType.Http],
			["http://142.42.1.1:8080/", AscCommon.c_oAscUrlType.Http],
			["http://foo.com/blah_(wikipedia)_blah#cite-1", AscCommon.c_oAscUrlType.Http],
			["http://foo.bar/?q=Test%20URL-encoded%20stuff", AscCommon.c_oAscUrlType.Http],
			["http://a.b-c.de", AscCommon.c_oAscUrlType.Http],
			["ftp://public.ftp-servers.example.com/mydirectory/myfile.txt", AscCommon.c_oAscUrlType.Http],
			["ftp://user001:secretpassword@private.ftp-servers.example.com/mydirectory/myfile.txt", AscCommon.c_oAscUrlType.Http],
			["ftps://user001:secretpassword@private.ftp-servers.example.com/mydirectory/myfile.txt", AscCommon.c_oAscUrlType.Http],
			["http://a.b-c.de", AscCommon.c_oAscUrlType.Http],
			["http://مثال.إختبار", AscCommon.c_oAscUrlType.Unsafe],//todo Http
			["http://фывап.ролдж", AscCommon.c_oAscUrlType.Http],
			["http://", AscCommon.c_oAscUrlType.Unsafe],//todo Invalid ?
			["http:///a", AscCommon.c_oAscUrlType.Unsafe],
			["http://.www.foo.bar/", AscCommon.c_oAscUrlType.Unsafe],

			["mysite@ourearth.com", AscCommon.c_oAscUrlType.Http],//todo Email
			["my.ownsite@ourearth.org", AscCommon.c_oAscUrlType.Email],
			["mysite@you.me.net", AscCommon.c_oAscUrlType.Http],//todo Email
			["mysite@.com.my", AscCommon.c_oAscUrlType.Email],//todo Invalid
			["@you.me.net", AscCommon.c_oAscUrlType.Http],//todo Invalid
			[".mysite@mysite.org", AscCommon.c_oAscUrlType.Email],//todo Invalid
			["mysite()*@gmail.com", AscCommon.c_oAscUrlType.Invalid],

			["smb://192.168.56.1/e/Testfolder/TestFile.docx", AscCommon.c_oAscUrlType.Unsafe],

			["tessa://tessaclient.EPD/?Action=OpenCard&ID=c40076f5-daa9-4929-8f66-d3fd6ae2dcb1", AscCommon.c_oAscUrlType.Unsafe],
			["joplin://x-callback-url/openFolder?id=1234", AscCommon.c_oAscUrlType.Unsafe],

			["file://localhost/etc/fstab", AscCommon.c_oAscUrlType.Unsafe],
			["file:///etc/fstab", AscCommon.c_oAscUrlType.Unsafe],
			["file://localhost/c:/WINDOWS/clock.avi", AscCommon.c_oAscUrlType.Unsafe],
			["file:///c:/WINDOWS/clock.avi", AscCommon.c_oAscUrlType.Unsafe],
			["file://\"C:\\Users\\User\\Documents\\About.pdf\"", AscCommon.c_oAscUrlType.Invalid],
			["file://'C:\\Users\\User\\Documents\\About.pdf'", AscCommon.c_oAscUrlType.Invalid],

			["/home/user/123.txt", AscCommon.c_oAscUrlType.Invalid],
			["123.txt", AscCommon.c_oAscUrlType.Http],
			["../../123.txt", AscCommon.c_oAscUrlType.Invalid],
		];
		for(let i = 0; i < test.length; ++i) {
			assert.strictEqual(editor.asc_getUrlType(test[i][0]), test[i][1], "Check " + test[i][0]);
		}
	});

	QUnit.test("Test asc_getUrlType desktop", function (assert)
	{
		let oldAscDesktopEditor = window["AscDesktopEditor"];
		window["AscDesktopEditor"] = {"IsLocalFile": function(){return true;}};

		const editor = new AscCommon.baseEditorsApi({});

		let test = [
			["http://foo.com/blah_blah", AscCommon.c_oAscUrlType.Http],
			["http://foo.com/blah_blah_(wikipedia)_(again)", AscCommon.c_oAscUrlType.Http],
			["https://www.example.com/foo/?bar=baz&inga=42&quux", AscCommon.c_oAscUrlType.Http],
			["http://userid:password@example.com:8080", AscCommon.c_oAscUrlType.Http],
			["http://userid@example.com:8080/", AscCommon.c_oAscUrlType.Http],
			["http://142.42.1.1", AscCommon.c_oAscUrlType.Http],
			["http://142.42.1.1:8080/", AscCommon.c_oAscUrlType.Http],
			["http://foo.com/blah_(wikipedia)_blah#cite-1", AscCommon.c_oAscUrlType.Http],
			["http://foo.bar/?q=Test%20URL-encoded%20stuff", AscCommon.c_oAscUrlType.Http],
			["http://a.b-c.de", AscCommon.c_oAscUrlType.Http],
			["ftp://public.ftp-servers.example.com/mydirectory/myfile.txt", AscCommon.c_oAscUrlType.Http],
			["ftp://user001:secretpassword@private.ftp-servers.example.com/mydirectory/myfile.txt", AscCommon.c_oAscUrlType.Http],
			["ftps://user001:secretpassword@private.ftp-servers.example.com/mydirectory/myfile.txt", AscCommon.c_oAscUrlType.Http],
			["http://a.b-c.de", AscCommon.c_oAscUrlType.Http],
			["http://مثال.إختبار", AscCommon.c_oAscUrlType.Unsafe],//todo Http
			["http://фывап.ролдж", AscCommon.c_oAscUrlType.Http],
			["http://", AscCommon.c_oAscUrlType.Unsafe],//todo Invalid ?
			["http:///a", AscCommon.c_oAscUrlType.Unsafe],
			["http://.www.foo.bar/", AscCommon.c_oAscUrlType.Unsafe],

			["mysite@ourearth.com", AscCommon.c_oAscUrlType.Http],//todo Email
			["my.ownsite@ourearth.org", AscCommon.c_oAscUrlType.Email],
			["mysite@you.me.net", AscCommon.c_oAscUrlType.Http],//todo Email
			["mysite@.com.my", AscCommon.c_oAscUrlType.Email],//todo Invalid
			["@you.me.net", AscCommon.c_oAscUrlType.Http],//todo Invalid
			[".mysite@mysite.org", AscCommon.c_oAscUrlType.Email],//todo Invalid
			["mysite()*@gmail.com", AscCommon.c_oAscUrlType.Invalid],

			["smb://192.168.56.1/e/Testfolder/TestFile.docx", AscCommon.c_oAscUrlType.Unsafe],

			["tessa://tessaclient.EPD/?Action=OpenCard&ID=c40076f5-daa9-4929-8f66-d3fd6ae2dcb1", AscCommon.c_oAscUrlType.Unsafe],

			["file://localhost/etc/fstab", AscCommon.c_oAscUrlType.Unsafe],
			["file:///etc/fstab", AscCommon.c_oAscUrlType.Unsafe],
			["file://localhost/c:/WINDOWS/clock.avi", AscCommon.c_oAscUrlType.Unsafe],
			["file:///c:/WINDOWS/clock.avi", AscCommon.c_oAscUrlType.Unsafe],
			["file://\"C:\\Users\\User\\Documents\\About.pdf\"", AscCommon.c_oAscUrlType.Invalid],
			["file://'C:\\Users\\User\\Documents\\About.pdf'", AscCommon.c_oAscUrlType.Invalid],

			["/home/user/123.txt", AscCommon.c_oAscUrlType.Unsafe],
			["123.txt", AscCommon.c_oAscUrlType.Unsafe],
			["../../123.txt", AscCommon.c_oAscUrlType.Unsafe],
		];
		for(let i = 0; i < test.length; ++i) {
			assert.strictEqual(editor.asc_getUrlType(test[i][0]), test[i][1], "Check " + test[i][0]);
		}
		window["AscDesktopEditor"] = oldAscDesktopEditor;
	});
});
