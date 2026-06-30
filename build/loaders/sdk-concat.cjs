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
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 */

/**
 * sdk-concat-loader
 *
 * Reads the ordered SDK JSON configs and returns ALL source files for a given
 * module/chunk as a single concatenated module. This is the only correct way to
 * bundle sdkjs under webpack: 69+ files across the word SDK alone use bare
 * top-level `var` declarations (no IIFE) that communicate across file boundaries
 * via concatenated scope. Putting all files into ONE webpack module preserves
 * that scope — every bare `var` is visible to every other file in the same chunk.
 *
 * Mirrors the CConfig + getFilesMin/getFilesAll logic from the original Gruntfile.js.
 *
 * Options (webpack loader options object):
 *   module    {string}   'word' | 'cell' | 'slide' | 'visio'           required
 *   chunk     {string}   'min' | 'all'                                  required
 *   platform  {string}   '' | 'desktop' | 'mobile'                      default ''
 *   srcRoot   {string}   absolute path to sdkjs root (one level above build/)
 *   addonDirs {string[]} absolute paths to addon directories
 */

'use strict';

const path = require('path');
const fs   = require('fs');

// ---------------------------------------------------------------------------
// Config loading — exact port of CConfig.prototype.append from Gruntfile.js
// ---------------------------------------------------------------------------

function loadJsonConfig(configsDir, name) {
    const file = path.join(configsDir, name + '.json');
    if (!fs.existsSync(file)) return null;
    try {
        return JSON.parse(fs.readFileSync(file, 'utf8'));
    } catch (e) {
        throw new Error(`sdk-concat-loader: failed to parse ${file}: ${e.message}`);
    }
}

function fixPath(obj, basePath) {
    if (Array.isArray(obj)) {
        for (let i = 0; i < obj.length; i++) {
            obj[i] = path.join(basePath, obj[i]);
        }
        return;
    }
    for (const k of Object.keys(obj)) {
        fixPath(obj[k], basePath);
    }
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

function loadAllConfigs(srcRoot, addonDirs) {
    const configs = {};
    const configsDir = path.join(srcRoot, 'configs');

    for (const name of ['word', 'cell', 'slide', 'visio']) {
        const cfg = loadJsonConfig(configsDir, name);
        if (cfg) {
            fixPath(cfg, srcRoot);
            configs[name] = cfg;
        }
    }

    for (const addonDir of (addonDirs || [])) {
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

// ---------------------------------------------------------------------------
// File list helpers — exact port of getFilesMin/getFilesAll from Gruntfile.js
// ---------------------------------------------------------------------------

function getFilesMin(sdkCfg, platform) {
    let files = (sdkCfg['min'] || []).slice();
    if (platform === 'mobile' && sdkCfg['mobile_banners']) {
        files = sdkCfg['mobile_banners']['min'].concat(files);
    }
    if (platform === 'desktop' && sdkCfg['desktop']) {
        files = files.concat(sdkCfg['desktop']['min'] || []);
    }
    return files;
}

function getFilesAll(sdkCfg, platform) {
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
        files = files.concat(sdkCfg['desktop']['common'] || []);
    }
    return files;
}

// ---------------------------------------------------------------------------
// Loader
// ---------------------------------------------------------------------------

module.exports = function sdkConcatLoader() {
    // this.resourcePath is dummy.js — its content is irrelevant; we ignore it.
    const opts      = this.getOptions();
    const srcRoot   = path.resolve(opts.srcRoot   || path.join(this.context, '..'));
    const platform  = opts.platform  || '';
    const addonDirs = opts.addonDirs || [];

    const configs = loadAllConfigs(srcRoot, addonDirs);
    const sdkCfg  = configs[opts.module] && configs[opts.module]['sdk'];

    if (!sdkCfg) {
        this.emitError(new Error(`sdk-concat-loader: no config found for module "${opts.module}" at ${srcRoot}`));
        return '';
    }

    const files = opts.chunk === 'min'
        ? getFilesMin(sdkCfg, platform)
        : getFilesAll(sdkCfg, platform);

    // Register every source file as a webpack dependency so watch mode works.
    for (const f of files) {
        this.addDependency(path.resolve(f));
    }
    // Watch the config file for this module so a config change triggers a rebuild.
    this.addDependency(path.join(srcRoot, 'configs', opts.module + '.json'));

    const parts = [];
    for (const f of files) {
        try {
            parts.push(fs.readFileSync(f, 'utf8'));
        } catch (e) {
            this.emitError(new Error(`sdk-concat-loader: cannot read ${f}: ${e.message}`));
            parts.push('');
        }
    }

    const content = parts.join('\n');

    // sdk-all.js: wrap in (function(window, undefined){...})(window) to match
    // the original Closure Compiler --chunk_wrapper for the sdk-all chunk.
    // sdk-all-min.js: no wrapper — it exposes bootstrap globals consumed by sdk-all.js.
    return opts.chunk === 'all'
        ? `(function(window, undefined) {\n${content}\n})(window);`
        : content;
};

module.exports.schema = {
    type: 'object',
    properties: {
        module:    { type: 'string', enum: ['word', 'cell', 'slide', 'visio'] },
        chunk:     { type: 'string', enum: ['min', 'all'] },
        platform:  { type: 'string', enum: ['', 'desktop', 'mobile'] },
        srcRoot:   { type: 'string' },
        addonDirs: { type: 'array', items: { type: 'string' } },
    },
    required: ['module', 'chunk'],
    additionalProperties: false,
};
