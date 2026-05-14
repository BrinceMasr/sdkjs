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

// Example: node runSingle.js D:/Git/sdkjs/tests/word/api/api.html

const testPath = process.argv[2];

if (!testPath)
{
	console.error("Usage: node runSingle.js <absolute-path-to-test.html>");
	console.error("Example: node runSingle.js D:/Git/sdkjs/tests/word/api/api.html");
	process.exit(1);
}

const {performance} = require('perf_hooks');

const {
  runQunitPuppeteer,
  printResultSummary,
  printFailedTests
} = require("node-qunit-puppeteer");

(async function()
{
	let startTime = performance.now();

	console.log("\n" + testPath.yellow.bold);

	try
	{
		const result = await runQunitPuppeteer({targetUrl: testPath, timeout: 60000});
		printResultSummary(result, console);

		if (result.stats.failed > 0)
		{
			printFailedTests(result, console);
			console.log("\nFAILED".red.bold);
		}
		else
		{
			console.log("\nPASSED".green.bold);
		}

		console.log("\nElapsed " + (Math.round(( ((performance.now() - startTime) / 1000) + Number.EPSILON) * 1000) / 1000) + "s");
		process.exit(result.stats.failed > 0 ? 1 : 0);
	}
	catch (ex)
	{
		console.error(ex);
		console.log("\nFAILED".red.bold);
		process.exit(1);
	}
})();
