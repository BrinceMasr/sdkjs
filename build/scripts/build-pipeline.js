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

// Full grunt-free build pipeline for Euro Office sdkjs.
//
// Usage (from sdkjs/build/):
//   PRODUCT_VERSION=9.2.1 BUILD_ROOT=/path/to/deploy node scripts/build-pipeline.js
//
// Options (env vars):
//   PRODUCT_VERSION   default '0.0.0'
//   BUILD_ROOT        default ../deploy/sdkjs
//   BUILD_NUMBER      default '0'
//   COMPANY_NAME      default 'onlyoffice'
//   SDK_PLATFORM      '' | 'desktop' | 'mobile' — passed through to webpack configs
//   SDK_ADDONS        path.delimiter-separated addon directories
//   SKIP_DEVELOP      set to '1' to skip develop scripts generation
//
// Phase layout (wall-clock optimised):
//   Phase 1 — parallel: deploy-assets + webpack ×4 (word, cell, slide, visio)
//             Each webpack config runs 2 compiler configs (min + all chunk) in parallel.
//   Phase 2 — sequential: build-develop (writes develop/sdkjs/{module}/scripts.js)

const { spawn } = require('child_process');
const path      = require('path');
const fs        = require('fs');

const BUILD_DIR = path.resolve(__dirname, '..');

const BUILD_ROOT = process.env.BUILD_ROOT
    ? path.resolve(process.env.BUILD_ROOT, 'sdkjs')
    : path.resolve(BUILD_DIR, '..', 'deploy', 'sdkjs');

const PRODUCT_VERSION = process.env.PRODUCT_VERSION || '0.0.0';
const BUILD_NUMBER    = String(process.env.BUILD_NUMBER || process.env.GITHUB_RUN_NUMBER || '0');
const SKIP_DEVELOP    = process.env.SKIP_DEVELOP === '1';

const CHILD_ENV = {
    ...process.env,
    PRODUCT_VERSION,
    BUILD_NUMBER,
    BUILD_ROOT: process.env.BUILD_ROOT || path.resolve(BUILD_DIR, '..', 'deploy'),
};

// ---- output helpers (mirrors web-apps/build/scripts/build-pipeline.js) ----

const BOLD  = s => `\x1b[1m${s}\x1b[0m`;
const DIM   = s => `\x1b[2m${s}\x1b[0m`;
const GREEN = s => `\x1b[32m${s}\x1b[0m`;
const RED   = s => `\x1b[31m${s}\x1b[0m`;
const CYAN  = s => `\x1b[36m${s}\x1b[0m`;
const PAD   = 20;

function elapsed(ms) {
    return ms < 1000 ? `${ms}ms` : `${(ms / 1000).toFixed(1)}s`;
}

function banner(msg) {
    process.stdout.write(`\n${BOLD(CYAN('▶ ' + msg))}\n`);
}

// ---- task runner -----------------------------------------------------------

function task(label, cmd, args = [], opts = {}) {
    return { label, cmd, args, opts };
}

function runTask({ label, cmd, args, opts = {} }) {
    let child = null;
    const promise = new Promise(resolve => {
        const start = Date.now();
        const paddedLabel = label.padEnd(PAD);
        const stderrBuf = [];

        child = spawn(cmd, args, {
            env:   { ...CHILD_ENV, ...(opts.env || {}) },
            cwd:   opts.cwd || BUILD_DIR,
            stdio: ['ignore', 'pipe', 'pipe'],
        });

        child.stdout.on('data', chunk => {
            for (const line of chunk.toString().split('\n')) {
                if (line.trim()) process.stdout.write(`  ${DIM('[' + label + ']')} ${line}\n`);
            }
        });

        child.stderr.on('data', chunk => { stderrBuf.push(chunk.toString()); });

        child.on('error', err => {
            const ms = Date.now() - start;
            process.stdout.write(`  ${RED('✗')} ${paddedLabel} ${RED('FAILED')} ${DIM(elapsed(ms))}\n`);
            process.stderr.write(`  spawn error: ${err.message}\n`);
            resolve({ label, ms, code: 1 });
        });

        child.on('exit', (code, signal) => {
            const ms = Date.now() - start;
            if (signal) {
                process.stdout.write(`  ${DIM('○')} ${paddedLabel} ${DIM('killed ' + elapsed(ms))}\n`);
                if (stderrBuf.length) process.stderr.write(stderrBuf.join(''));
                resolve({ label, ms, code: -1 });
            } else if (code === 0) {
                process.stdout.write(`  ${GREEN('✓')} ${paddedLabel} ${DIM(elapsed(ms))}\n`);
                resolve({ label, ms, code: 0 });
            } else {
                process.stdout.write(`  ${RED('✗')} ${paddedLabel} ${RED('FAILED')} ${DIM(elapsed(ms))}\n`);
                if (stderrBuf.length) process.stderr.write(stderrBuf.join(''));
                resolve({ label, ms, code });
            }
        });
    });
    return { promise, kill: () => child && child.kill('SIGTERM'), label };
}

async function phase(title, taskSpecs) {
    const count = taskSpecs.length;
    banner(`${title} — ${count} task${count !== 1 ? 's' : ''}`);

    const running = taskSpecs.map(runTask);
    let aborted = false;

    const results = await Promise.all(
        running.map(t =>
            t.promise.then(r => {
                if (r.code > 0 && !aborted) {
                    aborted = true;
                    running.forEach(o => { try { o.kill(); } catch (_) {} });
                }
                return r;
            })
        )
    );

    const failed = results.filter(r => r.code > 0);
    if (failed.length) {
        process.stderr.write(RED(`\n✗ ${failed.map(r => r.label).join(', ')} failed — aborting\n`));
        process.exit(1);
    }
    return results;
}

// ---- pipeline --------------------------------------------------------------

const node = process.execPath;
const wp   = path.join(BUILD_DIR, 'node_modules', '.bin', 'webpack');

const WEBPACK_CONFIGS = [
    'webpack.word.mjs',
    'webpack.cell.mjs',
    'webpack.slide.mjs',
    'webpack.visio.mjs',
];

async function main() {
    const wallStart = Date.now();

    process.stdout.write([
        BOLD('Euro Office sdkjs build pipeline'),
        `  BUILD_ROOT        ${BUILD_ROOT}`,
        `  PRODUCT_VERSION   ${PRODUCT_VERSION}`,
        `  BUILD_NUMBER      ${BUILD_NUMBER}`,
        `  SDK_PLATFORM      ${process.env.SDK_PLATFORM || '(default)'}`,
        `  SKIP_DEVELOP      ${SKIP_DEVELOP}`,
        '',
    ].join('\n'));

    // Clean deploy directory before building.
    if (fs.existsSync(BUILD_ROOT)) {
        fs.rmSync(BUILD_ROOT, { recursive: true, force: true });
    }

    // Phase 1: all independent work in parallel.
    //   - deploy-assets: copies CSS, fonts, images, themes, native JS (WHITESPACE compiled)
    //   - webpack ×4: each produces sdk-all-min.js + sdk-all.js for its module
    const phase1Tasks = [
        task('deploy-assets', node, ['scripts/deploy-assets.js']),
        ...WEBPACK_CONFIGS.map(cfg => {
            const name = cfg.replace('webpack.', '').replace('.mjs', '');
            return task(`webpack:${name}`, wp, ['--config', cfg]);
        }),
    ];

    const p1 = await phase('Phase 1 — parallel', phase1Tasks);

    // Phase 2: develop scripts (fast, sequential is fine).
    let p2 = [];
    if (!SKIP_DEVELOP) {
        p2 = await phase('Phase 2 — develop', [
            task('build-develop', node, ['scripts/build-develop.js']),
        ]);
    }

    // Summary
    const all = [...p1, ...p2];
    const wallMs = Date.now() - wallStart;
    const longestLabel = Math.max(...all.map(r => r.label.length));

    process.stdout.write([
        '',
        BOLD('Summary'),
        ...all.map(r => {
            const mark = r.code === 0 ? GREEN('✓') : r.code < 0 ? DIM('○') : RED('✗');
            return `  ${mark} ${r.label.padEnd(longestLabel + 2)} ${DIM(elapsed(r.ms))}`;
        }),
        '',
        `  Wall clock: ${BOLD(elapsed(wallMs))}`,
        '',
    ].join('\n'));
}

main().catch(err => {
    process.stderr.write(RED(`\nFatal: ${err.message || err}\n`));
    process.exit(1);
});
