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

const BUILD_DIR = path.resolve(__dirname, '..');
const SRC_ROOT  = path.resolve(BUILD_DIR, '..');

const BUILD_ROOT = process.env.BUILD_ROOT
    ? path.resolve(process.env.BUILD_ROOT, 'sdkjs')
    : path.resolve(BUILD_DIR, '..', 'deploy', 'sdkjs');

const DEVELOP_ROOT = process.env.BUILD_ROOT
    ? path.join(process.env.BUILD_ROOT, 'sdkjs', 'develop', 'sdkjs')
    : path.join(BUILD_DIR, '..', 'develop', 'sdkjs');

const platform  = process.env.SDK_PLATFORM || '';
const addonDirs = process.env.SDK_ADDONS
    ? process.env.SDK_ADDONS.split(path.delimiter).filter(Boolean)
    : [];
const compiled  = process.env.COMPILED === '1';

// ---- Config loading (mirrors CConfig from Gruntfile.js) --------------------

function loadJsonConfig(configsDir, name) {
    const file = path.join(configsDir, name + '.json');
    if (!fs.existsSync(file)) return null;
    return JSON.parse(fs.readFileSync(file, 'utf8'));
}

function fixPath(obj, basePath) {
    if (Array.isArray(obj)) {
        for (let i = 0; i < obj.length; i++) obj[i] = path.join(basePath, obj[i]);
        return;
    }
    for (const k of Object.keys(obj)) fixPath(obj[k], basePath);
}

function mergeConfigs(base, addon) {
    for (const k of Object.keys(addon)) {
        if (Array.isArray(addon[k])) {
            base[k] = Array.isArray(base[k]) ? base[k].concat(addon[k]) : addon[k];
        } else {
            if (!base[k]) base[k] = {};
            mergeConfigs(base[k], addon[k]);
        }
    }
}

function loadAllConfigs() {
    const configs = {};
    const configsDir = path.join(SRC_ROOT, 'configs');
    for (const name of ['word', 'cell', 'slide', 'visio']) {
        const cfg = loadJsonConfig(configsDir, name);
        if (cfg) { fixPath(cfg, SRC_ROOT); configs[name] = cfg; }
    }
    for (const addonDir of addonDirs) {
        for (const name of ['word', 'cell', 'slide', 'visio']) {
            if (!configs[name]) continue;
            const addon = loadJsonConfig(path.join(addonDir, 'configs'), name);
            if (!addon) continue;
            fixPath(addon, addonDir);
            mergeConfigs(configs[name], addon);
        }
    }
    return configs;
}

function getFilesMin(sdkCfg) {
    let files = (sdkCfg['min'] || []).slice();
    if (platform === 'mobile' && sdkCfg['mobile_banners']) {
        files = sdkCfg['mobile_banners']['min'].concat(files);
    }
    if (platform === 'desktop' && sdkCfg['desktop']) {
        files = files.concat(sdkCfg['desktop']['min']);
    }
    return files;
}

function getFilesAll(sdkCfg) {
    let files = (sdkCfg['common'] || []).slice();
    if (platform === 'mobile') {
        if (sdkCfg['mobile_banners']) {
            files = sdkCfg['mobile_banners']['common'].concat(files);
        }
        const exclude = sdkCfg['exclude_mobile'] || [];
        files = files.filter(f => !exclude.includes(f));
        files = files.concat(sdkCfg['mobile'] || []);
    }
    if (platform === 'desktop' && sdkCfg['desktop']) {
        files = files.concat(sdkCfg['desktop']['common']);
    }
    return files;
}

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
        if (process.env.BUILD_ROOT) {
            files.push(path.join('..', name, 'sdk-all-min.js'));
        } else {
            files.push(path.join(BUILD_ROOT, name, 'sdk-all-min.js'));
        }
    } else {
        files = files.concat(
            [path.join(SRC_ROOT, 'common', 'applyDocumentChanges.js')],
            getFilesMin(sdkCfg),
            getFilesAll(sdkCfg),
        );
    }

    // Convert absolute paths to relative URL strings anchored at build/
    // (mirrors fixUrl(files, '../../../../sdkjs/build/') from Gruntfile.js)
    files = fixUrl(
        files.map(f => path.relative(BUILD_DIR, f)),
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

// ---- main ------------------------------------------------------------------

function main() {
    const configs = loadAllConfigs();
    for (const name of ['word', 'cell', 'slide', 'visio']) {
        if (!configs[name]) {
            process.stderr.write(`build-develop: no config for ${name}, skipping\n`);
            continue;
        }
        writeScripts(configs[name]['sdk'], name);
    }
}

main();
