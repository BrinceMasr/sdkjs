#!/usr/bin/env node
/**
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
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 */

'use strict';

// Replaces the grunt build-develop task (writeScripts function in Gruntfile.js).
//
// Generates develop/sdkjs/{module}/scripts.js for each module (word/cell/slide/visio).
// The file contains a `var sdk_scripts = [...]` array of relative URLs to all
// uncompiled source files, used by the development HTML loader to load SDK
// without a build step.
//
// Mirrors writeScripts() + fixUrl() from Gruntfile.js exactly.
//
// Options (env vars):
//   BUILD_ROOT   if set, writes to $BUILD_ROOT/sdkjs/develop/sdkjs/{module}/scripts.js
//   SDK_PLATFORM '' | 'desktop' | 'mobile'
//   SDK_ADDONS   path.delimiter-separated addon directories
//   COMPILED     set to '1' to reference built sdk-all-min.js instead of source files

const path = require('path');
const fs   = require('fs');
const url  = require('url');
const { loadAllConfigs, getFilesMin, getFilesAll, expandGlobs } = require('../lib/sdk-configs.cjs');
const { parseAddonDirs, resolveBuildRoot } = require('../lib/env.cjs');

const BUILD_DIR = path.resolve(__dirname, '..');
const SRC_ROOT  = path.resolve(BUILD_DIR, '..');

const BUILD_ROOT = resolveBuildRoot(BUILD_DIR);

const DEVELOP_ROOT = process.env.BUILD_ROOT
    ? path.join(process.env.BUILD_ROOT, 'sdkjs', 'develop', 'sdkjs')
    : path.join(BUILD_DIR, '..', 'develop', 'sdkjs');

const platform  = process.env.SDK_PLATFORM || '';
const addonDirs = parseAddonDirs();
const compiled  = process.env.COMPILED === '1';

// ---- writeScripts (exact port of writeScripts() from Gruntfile.js) ---------

function fixUrl(arrPaths, basePath) {
    return arrPaths.map(p => url.resolve(basePath, p));
}

function writeScripts(sdkCfg, name) {
    let files = [
        path.join(SRC_ROOT, 'vendor', 'polyfill.js'),
        path.join(SRC_ROOT, 'common', 'AllFonts.js'),
    ];

    if (compiled) {
        // When process.env.BUILD_ROOT is set (e.g. Docker's /package), BUILD_ROOT
        // resolves to a path outside this checkout's directory tree, so
        // path.relative(BUILD_DIR, ...) below would compute a bogus path escaping
        // out to that external root instead of the sibling compiled bundle.
        // Push an already-relative entry in that case instead — mirrors the old
        // Gruntfile's writeScripts() special case exactly.
        files.push(process.env.BUILD_ROOT
            ? path.join('..', name, 'sdk-all-min.js')
            : path.join(BUILD_ROOT, name, 'sdk-all-min.js'));
    } else {
        files = files.concat(
            [path.join(SRC_ROOT, 'common', 'applyDocumentChanges.js')],
            expandGlobs(getFilesMin(sdkCfg, platform)),
            expandGlobs(getFilesAll(sdkCfg, platform)),
        );
    }

    // Convert absolute paths to relative URL strings anchored at build/
    // (mirrors fixUrl(files, '../../../../sdkjs/build/') from Gruntfile.js)
    files = fixUrl(
        files.map(f => path.isAbsolute(f) ? path.relative(BUILD_DIR, f) : f),
        '../../../../sdkjs/build/',
    );

    const outDir  = path.join(DEVELOP_ROOT, name);
    const outFile = path.join(outDir, 'scripts.js');
    fs.mkdirSync(outDir, { recursive: true });
    fs.writeFileSync(
        outFile,
        'var sdk_scripts = [\n\t"' + files.join('",\n\t"') + '"\n];',
        'utf8',
    );
    process.stdout.write(`build-develop: wrote ${outFile}\n`);
}

// Replaces the grunt copy-standalone task: device_scale.js is loaded directly
// by HTML templates (outside the SDK bundle), so it needs a standalone copy
// into the deploy root alongside develop/, not just inside the min/all bundles.
//
// Only needed when this script runs standalone (`npm run develop`, outside the
// full build-pipeline). The full pipeline already deploys a properly processed
// (terser + license header) copy via deploy-assets.cjs — running this here too
// would clobber it with a raw, unprocessed file. build-pipeline.cjs sets
// SKIP_STANDALONE=1 when invoking this script for that reason.
function copyStandalone() {
    const src  = path.join(SRC_ROOT, 'common', 'device_scale.js');
    const dest = path.join(BUILD_ROOT, 'common', 'device_scale.js');
    fs.mkdirSync(path.dirname(dest), { recursive: true });
    fs.copyFileSync(src, dest);
    process.stdout.write(`build-develop: wrote ${dest}\n`);
}

// ---- main ------------------------------------------------------------------

function main() {
    const configs = loadAllConfigs(SRC_ROOT, addonDirs);
    for (const name of ['word', 'cell', 'slide', 'visio']) {
        if (!configs[name] || !configs[name]['sdk']) {
            throw new Error(`build-develop: missing sdk config for ${name}`);
        }
        writeScripts(configs[name]['sdk'], name);
    }
    if (process.env.SKIP_STANDALONE !== '1') {
        copyStandalone();
    }
}

main();
