/*
 * (c) Copyright Ascensio System SIA 2010-2024
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
