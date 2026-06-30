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
 * Shared webpack 5 config factory for all SDK modules (word, cell, slide, visio).
 *
 * sdkjs has no module system — all source files are global IIFE or bare-var scripts
 * that communicate via window.AscCommon / window.AscWord etc. The sdk-concat-loader
 * reads the ordered JSON configs and returns all files as ONE concatenated module,
 * preserving the shared scope that bare `var` declarations depend on.
 *
 * Each call to sdkConfig() returns TWO webpack compiler configs:
 *   [0]  sdk-all-min  — bootstrap files (device_scale, browser, skin, API defs …)
 *   [1]  sdk-all      — full feature set, wrapped in (function(window,undefined){…})(window)
 *
 * Environment variables (all optional, mirror the original Gruntfile.js):
 *   BUILD_ROOT        override deploy root; defaults to ../deploy/sdkjs
 *   SDK_PLATFORM      '' | 'desktop' | 'mobile'
 *   SDK_ADDONS        path.delimiter-separated list of addon directories
 *   COMPANY_NAME      default 'onlyoffice'
 *   PRODUCT_VERSION   default '0.0.0'
 *   BUILD_NUMBER      default '0'
 *   BETA              default 'false'
 *   APP_COPYRIGHT     default 'Copyright (C) Ascensio System SIA …'
 *   PUBLISHER_URL     default 'https://www.onlyoffice.com/'
 *   NODE_ENV          'production' (default) | 'development'
 */

import webpack    from 'webpack';
import TerserPlugin from 'terser-webpack-plugin';
import path       from 'path';
import fs         from 'fs';
import { fileURLToPath } from 'url';

const __dirname     = path.dirname(fileURLToPath(import.meta.url));
const CONCAT_LOADER = path.join(__dirname, 'loaders', 'sdk-concat.cjs');
const DUMMY_ENTRY   = path.join(__dirname, 'dummy.js');

/**
 * @param {string} moduleName  'word' | 'cell' | 'slide' | 'visio'
 * @returns {object[]}  Two webpack compiler configs: [sdk-all-min, sdk-all]
 */
export function sdkConfig(moduleName) {
    const env = process.env.NODE_ENV || 'production';

    const BUILD_ROOT = process.env.BUILD_ROOT
        ? path.resolve(process.env.BUILD_ROOT, 'sdkjs')
        : path.resolve(__dirname, '..', 'deploy', 'sdkjs');

    const SRC_ROOT  = path.resolve(__dirname, '..');
    const OUT_DIR   = path.join(BUILD_ROOT, moduleName);
    const platform  = process.env.SDK_PLATFORM || '';
    const addonDirs = process.env.SDK_ADDONS
        ? process.env.SDK_ADDONS.split(path.delimiter).filter(Boolean)
        : [];

    const companyName  = process.env.COMPANY_NAME     || 'onlyoffice';
    const version      = process.env.PRODUCT_VERSION  || '0.0.0';
    const buildNumber  = process.env.BUILD_NUMBER      || '0';
    const beta         = process.env.BETA              || 'false';

    const appCopyright = process.env.APP_COPYRIGHT
        || `Copyright (C) Ascensio System SIA 2012-${new Date().getFullYear()}. All rights reserved`;
    const publisherUrl = process.env.PUBLISHER_URL || 'https://www.onlyoffice.com/';

    let licenseText = fs.readFileSync(path.join(__dirname, 'license.header'), 'utf8');
    licenseText = licenseText
        .replace('@@AppCopyright', appCopyright)
        .replace('@@PublisherUrl', publisherUrl)
        .replace('@@Version', version)
        .replace('@@Build', buildNumber);

    function chunkConfig(chunk, outName) {
        return {
            name: `${moduleName}:${chunk}`,
            mode: env,

            entry: {
                [outName]: DUMMY_ENTRY,
            },

            output: {
                path: OUT_DIR,
                filename: '[name].js',
                // iife:false — we control wrapping via the loader:
                //   sdk-all-min: no wrapper
                //   sdk-all:     (function(window, undefined){…})(window)
                // Letting webpack add its own ()=>{} on top would still work
                // (code sets window.xxx), but iife:false gives a cleaner output.
                iife: false,
                // Multiple chunk configs share OUT_DIR; do not wipe sibling output.
                clean: false,
            },

            module: {
                rules: [
                    {
                        // Match only our dummy entry, not real source files.
                        test: /[/\\]dummy\.js$/,
                        use: [
                            {
                                loader: CONCAT_LOADER,
                                options: {
                                    module:    moduleName,
                                    chunk,
                                    platform,
                                    srcRoot:   SRC_ROOT,
                                    addonDirs,
                                },
                            },
                        ],
                    },
                ],
            },

            plugins: [
                new webpack.BannerPlugin({
                    banner: licenseText,
                    raw: true,
                    entryOnly: true,
                }),

                // Replaces Closure Compiler's --define= flags.
                // webpack DefinePlugin performs AST-level identifier replacement
                // so dead-code branches (if (g_cIsBeta === 'true') …) are
                // eliminated by TerserPlugin in the same pass.
                new webpack.DefinePlugin({
                    'AscCommon.g_cCompanyName':    JSON.stringify(companyName),
                    'AscCommon.g_cProductVersion': JSON.stringify(version),
                    'AscCommon.g_cBuildNumber':    JSON.stringify(buildNumber),
                    'AscCommon.g_cIsBeta':         JSON.stringify(beta),
                }),
            ],

            optimization: {
                minimize: env === 'production',
                minimizer: [
                    new TerserPlugin({
                        extractComments: false,
                        terserOptions: {
                            format: {
                                // Preserve the license header injected by BannerPlugin.
                                comments: /AGPL|Copyright|Ascensio|License/i,
                            },
                            compress: {
                                drop_console: env === 'production',
                            },
                            // mangle:false is load-bearing — same reason as web-apps:
                            // sdkjs files communicate via window.AscCommon.xxx and bare
                            // top-level var declarations shared across concatenated scope.
                            // Mangling property names would silently corrupt those references.
                            mangle: false,
                        },
                    }),
                ],
            },

            devtool: env === 'production' ? false : 'source-map',
        };
    }

    return [
        chunkConfig('min', 'sdk-all-min'),
        chunkConfig('all', 'sdk-all'),
    ];
}
