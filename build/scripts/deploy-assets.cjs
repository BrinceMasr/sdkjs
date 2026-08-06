#!/usr/bin/env node
/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

'use strict';

// Replaces the grunt copy-other and copy-standalone tasks.
//
// For each entry in the otherFiles list (mirrors Gruntfile.js):
//   - Non-JS files: plain fs.cp (recursive for directories, flat for globs)
//   - JS files: run through terser (WHITESPACE_ONLY equivalent) + license header
//
// JS files ignored (same ignoreFiles list as Gruntfile.js):
//   jquery_native, fonts_ie, spell_ie, engine_ie, zlib_ie, drawingfile_ie, themes

const path  = require('path');
const fs    = require('fs');
const { sync: globSync } = require('glob');
const { minify }   = require('terser');
const { resolveBuildRoot } = require('../lib/env.cjs');

const BUILD_DIR = path.resolve(__dirname, '..');
const SRC_ROOT  = path.resolve(BUILD_DIR, '..');

const BUILD_ROOT = resolveBuildRoot(BUILD_DIR);

const version      = process.env.PRODUCT_VERSION || '0.0.0';
const buildNumber  = process.env.BUILD_NUMBER     || '0';
const appCopyright = process.env.APP_COPYRIGHT
    || `Copyright (C) Ascensio System SIA 2012-2025. All rights reserved; Euro-Office contributors 2026 - ${new Date().getFullYear()}`;
const publisherUrl = process.env.PUBLISHER_URL || 'https://github.com/Euro-Office/';

let licenseText = fs.readFileSync(path.join(BUILD_DIR, 'license.header'), 'utf8');
licenseText = licenseText
    .replace('@@AppCopyright', appCopyright)
    .replace('@@PublisherUrl', publisherUrl)
    .replace('@@Version', version)
    .replace('@@Build', buildNumber)
    // @@license-banner@@ only exists so webpack.sdk.factory.mjs's Terser pass can
    // distinguish this banner from per-file headers — irrelevant here since this
    // script never runs Terser over the banner text, so it must not ship.
    .replace(' @@license-banner@@', '');

// JS files skipped from individual minification (same as ignoreFiles in Gruntfile.js)
const IGNORE_NAMES = new Set([
    'jquery_native', 'fonts_ie', 'spell_ie', 'engine_ie',
    'zlib_ie', 'drawingfile_ie', 'themes',
]);

// Mirrors the otherFiles array in Gruntfile.js
const OTHER_FILES = [
    {
        cwd:  path.join(SRC_ROOT, 'vendor'),
        src:  ['polyfill.js'],
        dest: path.join(BUILD_ROOT, 'vendor'),
    },
    {
        cwd:  path.join(SRC_ROOT, 'common'),
        src:  [
            'device_scale.js',
            'Drawings/Format/path-boolean-min.js',
            'Charts/ChartStyles.js',
            'SmartArts/SmartArtData/*',
            'SmartArts/SmartArtDrawing/*',
            'Images/*',
            'Images/placeholders/*',
            'Images/content_controls/*',
            'Images/cursors/*',
            'Images/reporter/*',
            'Images/icons/*',
            'Native/*.js',
            'libfont/engine/*',
            'spell/spell/*',
            'hash/hash/*',
            'zlib/engine/*',
            'serviceworker/*',
        ],
        dest: path.join(BUILD_ROOT, 'common'),
    },
    {
        cwd:  path.join(SRC_ROOT, 'cell', 'css'),
        src:  ['*.css'],
        dest: path.join(BUILD_ROOT, 'cell', 'css'),
    },
    {
        cwd:  path.join(SRC_ROOT, 'slide', 'themes'),
        src:  ['**/**'],
        dest: path.join(BUILD_ROOT, 'slide', 'themes'),
    },
    {
        cwd:  path.join(SRC_ROOT, 'pdf'),
        src:  [
            'src/engine/*',
            'src/annotations/stamps/*.json',
        ],
        dest: path.join(BUILD_ROOT, 'pdf'),
    },
];

async function deployJsFile(srcPath, destPath) {
    const source  = fs.readFileSync(srcPath, 'utf8');
    const result  = await minify(source, {
        compress: false,
        mangle:   false,
        format:   { comments: false },
    });
    const content = licenseText + '\n' + (result.code != null ? result.code : source);
    fs.mkdirSync(path.dirname(destPath), { recursive: true });
    fs.writeFileSync(destPath, content, 'utf8');
}

function deployFile(srcPath, destPath) {
    fs.mkdirSync(path.dirname(destPath), { recursive: true });
    fs.copyFileSync(srcPath, destPath);
}

async function main() {
    const tasks = [];

    for (const entry of OTHER_FILES) {
        const matches = [];
        for (const pattern of entry.src) {
            const found = globSync(pattern, { cwd: entry.cwd, nodir: true });
            for (const f of found) matches.push(f);
        }

        for (const relFile of matches) {
            const ext      = path.extname(relFile);
            const baseName = path.parse(relFile).name;
            const srcPath  = path.join(entry.cwd, relFile);
            const destPath = path.join(entry.dest, relFile);

            if (ext === '.js' && !IGNORE_NAMES.has(baseName)) {
                tasks.push(deployJsFile(srcPath, destPath));
            } else {
                deployFile(srcPath, destPath);
            }
        }
    }

    await Promise.all(tasks);
    process.stdout.write(`deploy-assets: ${tasks.length} files deployed to ${BUILD_ROOT}\n`);
}

main().catch(err => {
    process.stderr.write(`deploy-assets FAILED: ${err.message}\n`);
    process.exit(1);
});
