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

'use strict';

const test   = require('node:test');
const assert = require('node:assert/strict');
const fs     = require('node:fs');
const os     = require('node:os');
const path   = require('node:path');

const sdkConcatLoader = require('../loaders/sdk-concat.cjs');
const { transpileToES5, stripDuplicateHelpers, cacheKeyFor } = sdkConcatLoader;

// Minimal on-disk fixture satisfying loadAllConfigs()'s expectations (a real
// configs/word.json + two small source files), so the loader can run
// end-to-end without a real sdkjs checkout.
function makeFixtureSrcRoot() {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), 'sdk-concat-test-'));
    fs.mkdirSync(path.join(root, 'configs'));
    fs.mkdirSync(path.join(root, 'src'));
    fs.mkdirSync(path.join(root, 'vendor'));

    fs.writeFileSync(path.join(root, 'src', 'a.js'), 'let a = (x) => x + 1;\n');
    fs.writeFileSync(path.join(root, 'src', 'b.js'), 'const b = 2;\n');
    fs.writeFileSync(path.join(root, 'vendor', 'polyfill.js'), 'var $jscomp_polyfill_marker = 1;\n');
    fs.writeFileSync(
        path.join(root, 'configs', 'word.json'),
        JSON.stringify({ sdk: { min: ['src/a.js', 'src/b.js'] } })
    );

    return root;
}

// Runs the loader with a minimal mocked webpack loader context, mirroring
// what webpack itself provides (this.async/getOptions/context/addDependency/
// sourceMap), and resolves with { content, map }.
function runLoader(srcRoot, opts) {
    return new Promise((resolve, reject) => {
        const context = {
            async: () => (err, content, map) => (err ? reject(err) : resolve({ content, map })),
            getOptions: () => opts,
            context: srcRoot,
            sourceMap: true,
            addDependency: () => {},
            addContextDependency: () => {},
        };
        sdkConcatLoader.call(context);
    });
}

test('transpileToES5: downlevels let/const/arrow functions to ES5 syntax', () => {
    const { code } = transpileToES5('let f = (x) => x + 1;', 'a.js', false);
    assert.equal(/\blet\b|=>/.test(code), false);
    assert.match(code, /var f = function/);
});

test('transpileToES5: does not wrap output in "use strict" or a module wrapper', () => {
    const { code } = transpileToES5('var x = 1;', 'a.js', false);
    assert.equal(code.includes('use strict'), false);
});

test('transpileToES5: returns no source map when needSourceMap is false', () => {
    const { map } = transpileToES5('var x = 1;', 'a.js', false);
    assert.equal(map, null);
});

test('transpileToES5: returns a source map when needSourceMap is true', () => {
    const { map } = transpileToES5('let x = 1;', 'a.js', true);
    assert.ok(map);
    assert.ok(Array.isArray(map.sources));
});

test('stripDuplicateHelpers: removes a second copy of an already-emitted helper', () => {
    const emitted = new Set(['_typeof']);
    const helper =
        'function _typeof(obj) {\n' +
        '  "@babel/helpers - typeof";\n' +
        '  return typeof obj;\n' +
        '}\n';
    const result = stripDuplicateHelpers(helper, emitted);

    assert.equal(result.includes('return typeof obj'), false);
    // Line count must be unchanged — replaced with blank lines, not deleted,
    // so per-file source maps built against the pre-strip output stay valid.
    assert.equal(result.split('\n').length, helper.split('\n').length);
});

test('stripDuplicateHelpers: keeps the first occurrence of a helper', () => {
    const emitted = new Set();
    const helper =
        'function _typeof(obj) {\n' +
        '  "@babel/helpers - typeof";\n' +
        '  return typeof obj;\n' +
        '}\n';
    const result = stripDuplicateHelpers(helper, emitted);

    assert.equal(result, helper);
    assert.ok(emitted.has('_typeof'));
});

test('stripDuplicateHelpers: leaves ordinary (non-helper) functions untouched even if seen before', () => {
    const emitted = new Set(['doStuff']);
    const code = 'function doStuff() {\n  return 1;\n}\n';
    const result = stripDuplicateHelpers(code, emitted);
    assert.equal(result, code);
});

test('cacheKeyFor: same content + same needSourceMap produces the same key', () => {
    assert.equal(cacheKeyFor('var x = 1;', true), cacheKeyFor('var x = 1;', true));
});

test('cacheKeyFor: needSourceMap is folded into the key (map vs no-map differ)', () => {
    assert.notEqual(cacheKeyFor('var x = 1;', true), cacheKeyFor('var x = 1;', false));
});

test('cacheKeyFor: different content produces a different key', () => {
    assert.notEqual(cacheKeyFor('var x = 1;', false), cacheKeyFor('var x = 2;', false));
});

test('loader: min chunk prepends vendor/polyfill.js so sdk-all-min.js is self-contained', async () => {
    const srcRoot = makeFixtureSrcRoot();
    try {
        const { content } = await runLoader(srcRoot, { module: 'word', chunk: 'min', srcRoot });
        assert.match(content, /^var \$jscomp_polyfill_marker = 1;/);
    } finally {
        fs.rmSync(srcRoot, { recursive: true, force: true });
    }
});

test('loader: fails the build if vendor/polyfill.js is missing for a min chunk', async () => {
    const srcRoot = makeFixtureSrcRoot();
    try {
        fs.rmSync(path.join(srcRoot, 'vendor', 'polyfill.js'));
        await assert.rejects(
            runLoader(srcRoot, { module: 'word', chunk: 'min', srcRoot }),
            /cannot read required polyfill file/
        );
    } finally {
        fs.rmSync(srcRoot, { recursive: true, force: true });
    }
});

test('loader: source map sourcesContent holds the original file content, not transpiled/dedup\'d output', async () => {
    const srcRoot = makeFixtureSrcRoot();
    try {
        const { map } = await runLoader(srcRoot, { module: 'word', chunk: 'min', srcRoot });

        assert.ok(map, 'expected a source map (this.sourceMap was true)');
        const aIndex = map.sources.findIndex(s => s.endsWith(path.join('src', 'a.js')));
        assert.notEqual(aIndex, -1);

        const originalA = fs.readFileSync(path.join(srcRoot, 'src', 'a.js'), 'utf8');
        assert.equal(map.sourcesContent[aIndex], originalA);
        // Specifically must NOT be the ES5-transpiled form.
        assert.equal(/=>/.test(map.sourcesContent[aIndex]), true);
    } finally {
        fs.rmSync(srcRoot, { recursive: true, force: true });
    }
});
