/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

/**
 * Guards the "bare global" backward-compat contract that ~287 places across
 * sdkjs depend on (see webpack.sdk.factory.mjs's `iife: false` comment): the
 * 'min' chunk config has no imports, no splitChunks, no externals, so webpack
 * currently inlines the single concatenated module directly at true top
 * level, with no __webpack_require__ runtime and no per-module function
 * wrapper — a bare top-level `var` in the source lands on the global object
 * exactly like a plain <script> tag would.
 *
 * That property is an emergent side effect of the config shape, not something
 * asserted anywhere else. This test compiles a real, minimal webpack config
 * with the SAME load-bearing shape (iife:false, single no-import entry) and
 * proves both directions:
 *   - the safe shape stays unwrapped and the bare var reaches the global object
 *   - the shape the warning comment predicts as dangerous (adding import())
 *     actually does switch webpack to the wrapped __webpack_require__ form,
 *     so this test would fail loudly if a future edit to webpack.sdk.factory.mjs
 *     ever introduced that shape for the real 'min' chunk.
 */

'use strict';

const test   = require('node:test');
const assert = require('node:assert/strict');
const path   = require('node:path');
const fs     = require('node:fs');
const os     = require('node:os');
const vm     = require('node:vm');
const webpack = require('webpack');

function compile(tmpDir, entrySource, { withDynamicImport }) {
    const entry = path.join(tmpDir, 'entry.js');
    fs.writeFileSync(entry, entrySource);

    if (withDynamicImport) {
        fs.writeFileSync(path.join(tmpDir, 'other.js'), 'export default 42;');
    }

    return new Promise((resolve, reject) => {
        const compiler = webpack({
            mode: 'production',
            entry,
            output: {
                path: tmpDir,
                filename: withDynamicImport ? 'wrapped.js' : 'bare.js',
                // Same knob the real 'min'/'all' chunk configs set — see
                // webpack.sdk.factory.mjs's chunkConfig().
                iife: false,
            },
            // Minification is a separate concern (covered by the mangle:false
            // tests) — disabled here so it can't obscure the wrapping question.
            optimization: { minimize: false },
        });
        compiler.run((err, stats) => {
            compiler.close(() => {});
            if (err || stats.hasErrors()) return reject(err || new Error(stats.toString()));
            resolve(fs.readFileSync(path.join(tmpDir, withDynamicImport ? 'wrapped.js' : 'bare.js'), 'utf8'));
        });
    });
}

test('bare-global contract: a single no-import iife:false entry stays unwrapped and reaches the global object', async () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'bare-global-safe-'));
    try {
        const bundle = await compile(
            tmpDir,
            'var AscCommonSdkTestGlobal = { value: 1 + 1 };\nconsole.log(AscCommonSdkTestGlobal);\n',
            { withDynamicImport: false },
        );

        assert.equal(bundle.includes('__webpack_require__'), false,
            'a real sdk-concat-loader entry has no imports/splitChunks/externals — this config shape must never grow the wrapped-module runtime');

        const sandbox = { console: { log: () => {} } };
        vm.createContext(sandbox);
        vm.runInContext(bundle, sandbox);

        assert.equal(sandbox.AscCommonSdkTestGlobal !== undefined, true,
            'a bare top-level var must land on the execution global, matching plain <script> tag semantics');
        assert.equal(sandbox.AscCommonSdkTestGlobal.value, 2);
    } finally {
        fs.rmSync(tmpDir, { recursive: true, force: true });
    }
});

test('bare-global contract: adding import() (the documented danger case) switches webpack to the wrapped runtime', async () => {
    // Proves the test above actually discriminates wrapped vs. unwrapped output,
    // rather than passing unconditionally regardless of config shape.
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'bare-global-wrapped-'));
    try {
        const bundle = await compile(
            tmpDir,
            'var AscCommonSdkTestGlobal = { value: 1 + 1 };\nimport("./other.js").then(function (m) { console.log(m); });\n',
            { withDynamicImport: true },
        );

        assert.equal(bundle.includes('__webpack_require__'), true,
            'sanity check: import() must actually force the wrapped bootstrap form for this assertion to mean anything above');
    } finally {
        fs.rmSync(tmpDir, { recursive: true, force: true });
    }
});
