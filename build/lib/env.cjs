/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
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
