/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

/**
 * Snapshot/assertion coverage for sdkConfig()'s per-platform Terser options
 * (build/webpack.sdk.factory.mjs's chunkConfig()). These exact knobs were
 * what commit f787b365b8 fixed to solve real bundle-bloat and duplicate
 * license-header bugs, but had zero test coverage — a future refactor could
 * silently invert the desktop/mobile `compress` branch (re-enabling dead-code
 * elimination on platforms that intentionally disable it) or drop the
 * `mangle:false` invariant sdkjs's bare-global/shared-scope model depends on,
 * with nothing here to catch it.
 */

'use strict';

const test   = require('node:test');
const assert = require('node:assert/strict');
const path   = require('node:path');
const url    = require('node:url');

async function loadSdkConfig() {
    const mod = await import(url.pathToFileURL(path.join(__dirname, '..', 'webpack.sdk.factory.mjs')));
    return mod.sdkConfig;
}

// sdkConfig() reads SDK_PLATFORM from process.env at call time (not at module
// load time), so tests can drive it directly as long as they restore it
// afterwards for any other test relying on the default ('').
function withPlatform(platform, fn) {
    const prev = process.env.SDK_PLATFORM;
    if (platform) {
        process.env.SDK_PLATFORM = platform;
    } else {
        delete process.env.SDK_PLATFORM;
    }
    try {
        return fn();
    } finally {
        if (prev === undefined) {
            delete process.env.SDK_PLATFORM;
        } else {
            process.env.SDK_PLATFORM = prev;
        }
    }
}

function terserOptionsOf(chunkConfigs) {
    // chunkConfigs is [minConfig, allConfig] from sdkConfig(); both chunks
    // share the same platform-derived terserOptions, so either is representative.
    return chunkConfigs[0].optimization.minimizer[0].options.minimizer.options;
}

test('sdkConfig: web platform (SDK_PLATFORM unset) enables Terser compress', async () => {
    const sdkConfig = await loadSdkConfig();
    const terserOptions = withPlatform('', () => terserOptionsOf(sdkConfig('word')));

    assert.notEqual(terserOptions.compress, false);
    assert.equal(typeof terserOptions.compress, 'object');
    assert.equal(terserOptions.mangle, false);
});

test('sdkConfig: desktop platform disables Terser compress (restores legacy WHITESPACE_ONLY-equivalent behavior)', async () => {
    const sdkConfig = await loadSdkConfig();
    const terserOptions = withPlatform('desktop', () => terserOptionsOf(sdkConfig('word')));

    assert.equal(terserOptions.compress, false);
    assert.equal(terserOptions.mangle, false);
});

test('sdkConfig: mobile platform disables Terser compress (same as desktop)', async () => {
    const sdkConfig = await loadSdkConfig();
    const terserOptions = withPlatform('mobile', () => terserOptionsOf(sdkConfig('word')));

    assert.equal(terserOptions.compress, false);
    assert.equal(terserOptions.mangle, false);
});

test('sdkConfig: comments format only preserves the license-banner sentinel, not every per-file AGPL header', async () => {
    const sdkConfig = await loadSdkConfig();
    const terserOptions = withPlatform('', () => terserOptionsOf(sdkConfig('word')));

    const commentsRegex = terserOptions.format.comments;
    assert.equal(commentsRegex.test('@@license-banner@@'), true);
    // A per-file AGPL header (repeated ~400+ times pre-minification) must NOT
    // match, or the duplicate-header bloat commit f787b365b8 fixed would regress.
    assert.equal(commentsRegex.test('This program is a free software product'), false);
});

test('sdkConfig: DROP_CONSOLE is opt-in, not the default, for non-desktop/mobile platforms', async () => {
    const sdkConfig = await loadSdkConfig();
    const prevDropConsole = process.env.DROP_CONSOLE;
    delete process.env.DROP_CONSOLE;
    try {
        const terserOptions = withPlatform('', () => terserOptionsOf(sdkConfig('word')));
        assert.equal(terserOptions.compress.drop_console, false);
    } finally {
        if (prevDropConsole === undefined) {
            delete process.env.DROP_CONSOLE;
        } else {
            process.env.DROP_CONSOLE = prevDropConsole;
        }
    }
});
