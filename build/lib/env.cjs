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
 * Shared env-var parsing used identically by webpack.sdk.factory.mjs,
 * scripts/build-develop.cjs, scripts/build-pipeline.cjs, and
 * scripts/deploy-assets.cjs — extracted so these copies can't silently
 * diverge (see build/lib/sdk-configs.cjs for the same rationale applied
 * to config loading).
 */

'use strict';

const path = require('path');

// process.env.SDK_ADDONS: path.delimiter-separated list of addon directories.
function parseAddonDirs(env) {
    env = env || process.env;
    return env.SDK_ADDONS
        ? env.SDK_ADDONS.split(path.delimiter).filter(Boolean)
        : [];
}

// process.env.BUILD_ROOT, resolved to the sdkjs-specific deploy dir.
// buildDir is the caller's build/ directory (path.resolve(__dirname, ...)),
// since the default falls back to <sdkjs-root>/deploy/sdkjs.
function resolveBuildRoot(buildDir, env) {
    env = env || process.env;
    return env.BUILD_ROOT
        ? path.resolve(env.BUILD_ROOT, 'sdkjs')
        : path.resolve(buildDir, '..', 'deploy', 'sdkjs');
}

module.exports = { parseAddonDirs, resolveBuildRoot };
