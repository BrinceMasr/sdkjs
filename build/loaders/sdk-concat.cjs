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
 *   buildMeta {object}   { companyName, version, buildNumber, beta } — patched into
 *                        window.AscCommon.g_cXxx after commonDefines.js ('min' chunk only)
 */

'use strict';

const path   = require('path');
const fs     = require('fs');
const crypto = require('crypto');
const babel = require('@babel/core');
const { loadAllConfigs, getFilesMin, getFilesAll, expandGlobs } = require('../lib/sdk-configs.cjs');

// Lazy-load source-map-js (optional dep — absent means no source maps, build still works).
let SourceMapGenerator = null;
let SourceMapConsumer  = null;
try {
    const sourceMapJs = require('source-map-js');
    SourceMapGenerator = sourceMapJs.SourceMapGenerator;
    SourceMapConsumer  = sourceMapJs.SourceMapConsumer;
} catch (_) {}

// On-disk cache for the per-file Babel transpile below, keyed on file content.
// Every loader invalidation (a single changed file in --watch mode, or each of
// the 4 separate webpack-cli processes the full pipeline spawns) otherwise
// re-transpiles every file in every chunk that includes it from scratch, even
// though the transpiled output only depends on the file's own content.
const CACHE_DIR = path.join(__dirname, '..', '.webpack-cache', 'babel');

// The actual babel preset/options object transpileToES5 passes — folded into
// the cache key below (not a manually-maintained version integer) so a future
// change to these options (e.g. the 'ie: 11' target) can't silently leave
// stale cache entries from a previous options set undetected.
const BABEL_OPTIONS_KEY = JSON.stringify({
    sourceType: 'script',
    presetEnvTargets: { ie: '11' },
    presetEnvModules: false,
});

// needSourceMap is folded into the key: a cache entry produced without a map
// (devtool:false runs) must never be handed back to a caller that needs one.
function cacheKeyFor(content, needSourceMap) {
    return crypto.createHash('sha1')
        .update(BABEL_OPTIONS_KEY)
        .update(needSourceMap ? 'map' : 'nomap')
        .update(content)
        .digest('hex');
}

function readCache(key) {
    try {
        return JSON.parse(fs.readFileSync(path.join(CACHE_DIR, key + '.json'), 'utf8'));
    } catch (_) {
        return null;
    }
}

function writeCache(key, entry) {
    // build-pipeline.cjs runs 4 webpack-cli processes in parallel against this
    // same cache dir, and aborts siblings by SIGTERM on the first failure —
    // writing straight to the final path risks a sibling reading that exact
    // key mid-write (a torn, unparseable JSON file) if it's killed at the
    // wrong moment. Write to a per-process temp file and rename into place:
    // rename is atomic on the same filesystem, so readers only ever see the
    // old complete file or the new complete file, never a partial one.
    const finalPath = path.join(CACHE_DIR, key + '.json');
    const tmpPath   = finalPath + '.' + process.pid + '.tmp';
    try {
        fs.mkdirSync(CACHE_DIR, { recursive: true });
        fs.writeFileSync(tmpPath, JSON.stringify(entry), 'utf8');
        fs.renameSync(tmpPath, finalPath);
    } catch (_) {
        // Cache is a pure optimization — a write failure (e.g. read-only fs) must
        // not fail the build.
        try { fs.unlinkSync(tmpPath); } catch (_) {}
    }
}

// The original Closure Compiler build ran with --language_out=ECMASCRIPT5:
// sdkjs source files use let/const/arrow functions/classes/etc, and Closure
// downleveled them to ES5 in the compiled output. webpack+Terser do not
// transpile syntax (Terser only avoids *introducing* new-ES syntax while
// minifying — it doesn't rewrite existing let/const/=> etc down to ES5), so
// that guarantee has to be reproduced explicitly here, per file, before
// concatenation. sourceType:'script' (not 'module') avoids Babel injecting
// "use strict" or wrapping content — required for the bare-var-across-files
// shared-scope model this loader depends on.
//
// Class/generator/destructuring transforms routinely change a file's line
// count (retainLines is best-effort, not a guarantee, and real sdkjs sources
// hit that on real inputs — see git history). So instead of relying on line
// counts staying stable, we ask Babel for its own generated↔original mapping
// (sourceMaps: needSourceMap) and stitch that into the combined map below via
// SourceMapConsumer, rather than assuming line N of the output is line N of
// the input. Skipping sourceMaps entirely when not needed (this.sourceMap is
// false) avoids the extra work across the 4 parallel webpack-cli processes ×
// 2 chunks the full pipeline spawns.
// Returns { code, map }.
function transpileToES5(content, filename, needSourceMap) {
    const result = babel.transformSync(content, {
        filename,
        babelrc:     false,
        configFile:  false,
        sourceType:  'script',
        sourceMaps:  needSourceMap,
        compact:     false,
        presets: [
            [require.resolve('@babel/preset-env'), { targets: { ie: '11' }, modules: false }],
        ],
    });

    return { code: result.code, map: result.map || null };
}

// Each per-file transformSync() call above independently inlines its own copy
// of any Babel helper it needs (_typeof, _classCallCheck, _inherits, …). That's
// fine standalone, but once concatenated into one shared-scope module, two
// files needing the same helper produce two top-level `function _typeof(){…}`
// declarations — a SyntaxError under strict/module parsing ("Identifier has
// already been declared"). Keep only the first occurrence of each helper
// across the whole chunk; babel tags helper functions with a recognizable
// `"@babel/helpers - name"` directive as their first statement, so detection
// doesn't depend on guessing every possible helper name.
// Removed text is replaced with a matching run of blank lines (not deleted
// outright) so every line number after the removed span is unchanged — the
// per-file source map built from Babel's own (pre-strip) output stays valid
// without needing to be recomputed here.
// code is always re-parsed here rather than reusing the AST transformSync
// already built: Babel's helper injection creates brand-new synthetic nodes
// with no source position (node.start/end are undefined on them), so slicing
// against a reused post-transform AST silently corrupts output instead of
// removing anything. A fresh parseSync of the actual generated text gives
// every node real, code-accurate offsets.
function stripDuplicateHelpers(code, emittedHelpers) {
    let ast;
    try {
        ast = babel.parseSync(code, { sourceType: 'script', babelrc: false, configFile: false });
    } catch (err) {
        // Shouldn't happen — code babel itself just emitted failing to
        // re-parse. Log so a future occurrence is diagnosable instead of
        // surfacing only as a confusing duplicate-declaration SyntaxError
        // at the concatenated-bundle level with nothing pointing back here.
        console.warn(`sdk-concat-loader: failed to re-parse transpiled output for helper dedup (${err.message}); leaving file un-deduplicated`);
        return code;
    }

    const removals = [];
    for (const node of ast.program.body) {
        if (node.type !== 'FunctionDeclaration' || !node.id) continue;
        // A leading string-literal statement in a function body is parsed as a
        // Directive (like "use strict"), not a regular ExpressionStatement —
        // that's how babel tags its helper functions.
        const directive = node.body.directives && node.body.directives[0];
        const isHelper = directive && directive.value.value.startsWith('@babel/helpers');
        if (!isHelper) continue;

        if (emittedHelpers.has(node.id.name)) {
            removals.push([node.start, node.end]);
        } else {
            emittedHelpers.add(node.id.name);
        }
    }

    if (!removals.length) return code;
    removals.sort((a, b) => b[0] - a[0]);
    for (const [start, end] of removals) {
        const removedNewlines = (code.slice(start, end).match(/\n/g) || []).length;
        code = code.slice(0, start) + '\n'.repeat(removedNewlines) + code.slice(end);
    }
    return code;
}

// ---------------------------------------------------------------------------
// Loader
// ---------------------------------------------------------------------------

module.exports = function sdkConcatLoader() {
    // this.resourcePath is dummy.js — its content is irrelevant; we ignore it.
    const callback  = this.async();
    const opts      = this.getOptions();
    const srcRoot   = path.resolve(opts.srcRoot   || path.join(this.context, '..'));
    const platform  = opts.platform  || '';
    const addonDirs = opts.addonDirs || [];

    const configs = loadAllConfigs(srcRoot, addonDirs);
    const sdkCfg  = configs[opts.module] && configs[opts.module]['sdk'];

    if (!sdkCfg) {
        callback(new Error(`sdk-concat-loader: no config found for module "${opts.module}" at ${srcRoot}`));
        return;
    }

    const rawFiles = opts.chunk === 'min'
        ? getFilesMin(sdkCfg, platform)
        : getFilesAll(sdkCfg, platform);

    const addCtx = this.addContextDependency.bind(this);
    const files  = expandGlobs(rawFiles, addCtx);

    // Register every source file as a webpack dependency so watch mode works.
    for (const f of files) {
        this.addDependency(path.resolve(f));
    }
    // Watch the config file(s) so a config change triggers a rebuild — including
    // addon configs, which loadAllConfigs() merges in but which webpack would
    // otherwise never see, leaving --watch/persistent-cache builds stale.
    this.addDependency(path.join(srcRoot, 'configs', opts.module + '.json'));
    for (const addonDir of addonDirs) {
        this.addDependency(path.join(addonDir, 'configs', opts.module + '.json'));
    }

    // this.sourceMap reflects webpack's own devtool setting — building the
    // full per-file map (and embedding every source's content) is wasted
    // CPU/memory when devtool:false, and this loader runs across 4 parallel
    // webpack-cli processes × 2 chunks each in the full pipeline.
    const needSourceMap = !!(this.sourceMap && SourceMapGenerator && SourceMapConsumer);

    // The old Closure Compiler build ran with --rewrite_polyfills=true, which
    // injected runtime polyfills (Promise, Map, WeakMap, Array.prototype.includes,
    // etc.) directly into the compiled SDK bundle for whatever ES6+ APIs the
    // source actually used. webpack+Babel's per-file syntax downlevel above
    // does NOT do this — preset-env here has no useBuiltIns/corejs configured,
    // so consumers that load only sdk-all-min.js (e.g. embed pages, which do
    // not load sdkjs/vendor/polyfill.js nearby) would regress on older
    // browsers. Prepend the same polyfill file the old dev-mode writeScripts()
    // path already loads ahead of sdk-all-min.js, so the 'min' chunk stays
    // self-contained exactly like the Closure output was.
    let polyfillContent = '';
    if (opts.chunk === 'min') {
        const polyfillPath = path.join(srcRoot, 'vendor', 'polyfill.js');
        this.addDependency(polyfillPath);
        try {
            polyfillContent = fs.readFileSync(polyfillPath, 'utf8');
        } catch (err) {
            callback(new Error(`sdk-concat-loader: cannot read required polyfill file ${polyfillPath}: ${err.message}`));
            return;
        }
        if (!polyfillContent.endsWith('\n')) polyfillContent += '\n';
    }

    // Read all source files in parallel. A missing/unreadable file must fail
    // the build outright: silently substituting '' would concatenate a chunk
    // with a piece of its bare-var scope missing, producing a build that looks
    // green but throws a confusing ReferenceError deep inside the SDK at runtime.
    Promise.all(files.map(f =>
        fs.promises.readFile(f, 'utf8').catch(err => {
            throw new Error(`sdk-concat-loader: cannot read ${f}: ${err.message}`);
        }).then(content => {
            const key    = cacheKeyFor(content, needSourceMap);
            const cached = readCache(key);
            // original is the pre-transpile file content — kept alongside the
            // transpiled code (not persisted to the on-disk cache, which is
            // keyed on it and can just re-read it from `content` here) so the
            // source map's sourcesContent can show real original source
            // instead of the transpiled/dedup'd output under the original
            // file's name.
            if (cached !== null) return { code: cached.code, map: cached.map, original: content };

            try {
                const transpiled = transpileToES5(content, f, needSourceMap);
                writeCache(key, { code: transpiled.code, map: transpiled.map });
                return { ...transpiled, original: content };
            } catch (err) {
                throw new Error(`sdk-concat-loader: failed to transpile ${f} to ES5: ${err.message}`);
            }
        })
    )).then(transpiled => {
        // Sequential (not part of the parallel read/transpile above) so "first
        // occurrence" tracking follows the files' configured concatenation
        // order, not whichever file's transpile happened to resolve first.
        const emittedHelpers = new Set();
        const contents = transpiled.map(t => stripDuplicateHelpers(t.code, emittedHelpers));

        const isAll     = opts.chunk === 'all';
        const prefix    = isAll ? '(function(window, undefined) {\n' : '';
        const suffix    = isAll ? '\n})(window);' : '';

        // --- Build output + optional per-file source map ---
        const bundleName = opts.chunk === 'min' ? 'sdk-all-min.js' : 'sdk-all.js';
        const gen = needSourceMap ? new SourceMapGenerator({ file: bundleName }) : null;

        let result  = prefix + polyfillContent;
        // Generated line cursor: prefix occupies line 1 (the wrapper open), content starts at line 2.
        // polyfillContent (min chunk only) is opaque vendor code — not remapped — so it
        // just shifts where the first real source file's mappings begin.
        let genLine = (isAll ? 2 : 1) + (polyfillContent.match(/\n/g) || []).length;

        for (let i = 0; i < files.length; i++) {
            const content  = contents[i];
            const nlCount  = (content.match(/\n/g) || []).length;
            const hasTrail = content.endsWith('\n');

            if (gen) {
                if (transpiled[i].map) {
                    // Use Babel's own generated↔original mapping for this file
                    // (transform may change line count) rather than assuming a
                    // 1:1 line correspondence. Offset every generated line by
                    // where this file's content starts in the combined output.
                    const consumer = new SourceMapConsumer(transpiled[i].map);
                    consumer.eachMapping(m => {
                        if (m.originalLine == null) return;
                        gen.addMapping({
                            generated: { line: genLine + (m.generatedLine - 1), column: m.generatedColumn },
                            original:  { line: m.originalLine, column: m.originalColumn },
                            source:    files[i],
                            name:      m.name || undefined,
                        });
                    });
                } else {
                    // Cache hit against a pre-existing no-map entry, or a file
                    // babel left byte-for-byte unchanged: fall back to a
                    // best-effort 1:1 line mapping.
                    const srcLines = content.split('\n');
                    const mapped   = hasTrail ? srcLines.length - 1 : srcLines.length;
                    for (let j = 0; j < mapped; j++) {
                        gen.addMapping({
                            generated: { line: genLine + j, column: 0 },
                            source:    files[i],
                            original:  { line: j + 1, column: 0 },
                        });
                    }
                }
                // Original (pre-transpile) source, not the transformed `content`
                // above — a debugger must show the real file it's labeled as.
                gen.setSourceContent(files[i], transpiled[i].original);
            }

            result += content;

            if (hasTrail) {
                // Trailing \n already provides the inter-file separator.
                genLine += nlCount;
            } else {
                result  += '\n';
                genLine += nlCount + 1;
            }
        }

        // commonDefines.js (part of the 'min' chunk) hardcodes
        // window.AscCommon.g_cXxx to placeholder values — DefinePlugin above
        // can't touch that assignment's LHS (see the comment at its call site
        // in webpack.sdk.factory.mjs). Patch the real build metadata in right
        // after, so the runtime-visible globals match the folded call-sites.
        if (opts.chunk === 'min' && opts.buildMeta) {
            const { companyName, version, buildNumber, beta } = opts.buildMeta;
            result +=
                '\nwindow.AscCommon.g_cCompanyName = '    + JSON.stringify(companyName) + ';' +
                '\nwindow.AscCommon.g_cProductVersion = ' + JSON.stringify(version)     + ';' +
                '\nwindow.AscCommon.g_cBuildNumber = '    + JSON.stringify(buildNumber) + ';' +
                '\nwindow.AscCommon.g_cIsBeta = '         + JSON.stringify(beta)        + ';\n';
        }

        result += suffix;

        callback(null, result, gen ? gen.toJSON() : undefined);
    }).catch(callback);
};

// Exposed for unit testing only (build/test/sdk-concat.test.cjs) — these are
// pure enough to test directly without spinning up a full webpack build.
module.exports.transpileToES5 = transpileToES5;
module.exports.stripDuplicateHelpers = stripDuplicateHelpers;
module.exports.cacheKeyFor = cacheKeyFor;

module.exports.schema = {
    type: 'object',
    properties: {
        module:    { type: 'string', enum: ['word', 'cell', 'slide', 'visio'] },
        chunk:     { type: 'string', enum: ['min', 'all'] },
        platform:  { type: 'string', enum: ['', 'desktop', 'mobile'] },
        srcRoot:   { type: 'string' },
        addonDirs: { type: 'array', items: { type: 'string' } },
        buildMeta: {
            type: 'object',
            properties: {
                companyName: { type: 'string' },
                version:     { type: 'string' },
                buildNumber: { type: 'string' },
                beta:        { type: 'string' },
            },
        },
    },
    required: ['module', 'chunk'],
    additionalProperties: false,
};
