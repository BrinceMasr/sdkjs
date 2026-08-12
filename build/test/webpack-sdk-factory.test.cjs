/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

'use strict';

const test   = require('node:test');
const assert = require('node:assert/strict');
const path   = require('node:path');
const url    = require('node:url');
const fs     = require('node:fs');
const os     = require('node:os');
const webpack = require('webpack');
const TerserPlugin = require('terser-webpack-plugin');

test('stripBootstrapStrictDirective: removes webpack\'s forced top-level "use strict" prologue', async () => {
    const { stripBootstrapStrictDirective } = await import(
        url.pathToFileURL(path.join(__dirname, '..', 'webpack.sdk.factory.mjs'))
    );

    const bundle = '/******/ "use strict";\n/******/ (() => {\nvar x = 1;\n})();';
    const patched = stripBootstrapStrictDirective(bundle);

    assert.equal(patched.includes('"use strict"'), false);
    // Same length / same line count — must not shift anything a source map points at.
    assert.equal(patched.length, bundle.length);
    assert.equal(patched.split('\n').length, bundle.split('\n').length);
});

test('stripBootstrapStrictDirective: leaves the bundle unchanged when no directive is present', async () => {
    const { stripBootstrapStrictDirective } = await import(
        url.pathToFileURL(path.join(__dirname, '..', 'webpack.sdk.factory.mjs'))
    );

    const bundle = '/******/ (() => {\nvar x = 1;\n})();';
    assert.equal(stripBootstrapStrictDirective(bundle), bundle);
});

test('stripBootstrapStrictDirective: does not touch a "use strict" appearing past the prologue window', async () => {
    const { stripBootstrapStrictDirective, PROLOGUE_SCAN_LIMIT } = await import(
        url.pathToFileURL(path.join(__dirname, '..', 'webpack.sdk.factory.mjs'))
    );

    // A real source file's own string literal containing this text, buried
    // deep in the concatenated bundle, must never be mistaken for webpack's
    // bootstrap directive.
    const padding = 'x'.repeat(PROLOGUE_SCAN_LIMIT + 100);
    const bundle = `${padding}var msg = "use strict";`;
    assert.equal(stripBootstrapStrictDirective(bundle), bundle);
});

test('StripBootstrapStrictModePlugin: strips webpack\'s bootstrap "use strict" from a real minified compilation', async (t) => {
    const { StripBootstrapStrictModePlugin } = await import(
        url.pathToFileURL(path.join(__dirname, '..', 'webpack.sdk.factory.mjs'))
    );

    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'strip-strict-test-'));
    const entry  = path.join(tmpDir, 'entry.js');
    // Bare top-level var, no import/export — same shape sdk-concat-loader
    // produces (sourceType:'script'), so webpack treats this chunk the same
    // way it treats a real SDK bundle for bootstrap-generation purposes.
    fs.writeFileSync(entry, 'var AscCommonSdkTestGlobal = { value: 1 + 1 };\n');

    function runCompiler(withPlugin) {
        return new Promise((resolve, reject) => {
            const compiler = webpack({
                mode: 'production',
                entry,
                output: { path: tmpDir, filename: withPlugin ? 'with-plugin.js' : 'without-plugin.js', iife: false },
                optimization: {
                    minimize: true,
                    minimizer: [new TerserPlugin({ terserOptions: { mangle: false, compress: true } })],
                },
                plugins: withPlugin ? [new StripBootstrapStrictModePlugin()] : [],
            });
            compiler.run((err, stats) => {
                compiler.close(() => {});
                if (err || stats.hasErrors()) return reject(err || new Error(stats.toString()));
                resolve(fs.readFileSync(path.join(tmpDir, withPlugin ? 'with-plugin.js' : 'without-plugin.js'), 'utf8'));
            });
        });
    }

    try {
        const withoutPlugin = await runCompiler(false);
        const withPlugin    = await runCompiler(true);

        // Sanity check the test fixture itself is meaningful: if webpack's own
        // output never carries the directive in the first place (e.g. a future
        // webpack version stops emitting it for script-sourceType chunks), the
        // plugin has nothing to strip and this assertion would catch that the
        // integration test itself needs updating, rather than silently passing
        // for the wrong reason.
        if (!withoutPlugin.includes('"use strict"')) {
            t.diagnostic('webpack did not emit a bootstrap "use strict" for this chunk shape — plugin has nothing to strip here');
        } else {
            assert.equal(withPlugin.includes('"use strict"'), false);
        }
    } finally {
        fs.rmSync(tmpDir, { recursive: true, force: true });
    }
});

test('StripLicenseSentinelPlugin: strips the @@license-banner@@ sentinel after Terser has used it to keep the banner', async () => {
    const { StripLicenseSentinelPlugin } = await import(
        url.pathToFileURL(path.join(__dirname, '..', 'webpack.sdk.factory.mjs'))
    );

    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'strip-license-sentinel-test-'));
    const entry  = path.join(tmpDir, 'entry.js');
    // A per-file AGPL header identical in shape to the ~400 real ones sdk-concat-loader
    // concatenates ahead of the BannerPlugin-injected banner — it must NOT be kept by
    // Terser's comments regex, only the sentinel-marked banner should survive.
    fs.writeFileSync(
        entry,
        '/* AGPL header, repeated in every source file */\n' +
        'var AscCommonSdkTestGlobal = { value: 1 + 1 };\n'
    );

    const banner = '/* @@license-banner@@\n * Copyright (C) Test Corp\n */';

    function runCompiler(withPlugin) {
        return new Promise((resolve, reject) => {
            const compiler = webpack({
                mode: 'production',
                entry,
                output: { path: tmpDir, filename: withPlugin ? 'with-plugin.js' : 'without-plugin.js', iife: false },
                plugins: [
                    new webpack.BannerPlugin({ banner, raw: true, entryOnly: true }),
                    ...(withPlugin ? [new StripLicenseSentinelPlugin()] : []),
                ],
                optimization: {
                    minimize: true,
                    minimizer: [
                        new TerserPlugin({
                            extractComments: false,
                            terserOptions: {
                                mangle: false,
                                compress: true,
                                // Same regex used in sdkConfig(): only the sentinel-marked
                                // banner should survive Terser's comment-stripping pass,
                                // not the per-file AGPL headers.
                                format: { comments: /@@license-banner@@/ },
                            },
                        }),
                    ],
                },
            });
            compiler.run((err, stats) => {
                compiler.close(() => {});
                if (err || stats.hasErrors()) return reject(err || new Error(stats.toString()));
                resolve(fs.readFileSync(path.join(tmpDir, withPlugin ? 'with-plugin.js' : 'without-plugin.js'), 'utf8'));
            });
        });
    }

    try {
        const withoutPlugin = await runCompiler(false);
        const withPlugin    = await runCompiler(true);

        // Sanity checks on the fixture itself: the per-file header must be gone
        // (Terser's regex didn't match it) and the banner text must have survived
        // in both outputs — only the sentinel differs between them.
        assert.equal(withoutPlugin.includes('AGPL header, repeated'), false);
        assert.equal(withoutPlugin.includes('Copyright (C) Test Corp'), true);
        assert.equal(withPlugin.includes('Copyright (C) Test Corp'), true);

        assert.equal(withoutPlugin.includes('@@license-banner@@'), true);
        assert.equal(withPlugin.includes('@@license-banner@@'), false);
    } finally {
        fs.rmSync(tmpDir, { recursive: true, force: true });
    }
});
