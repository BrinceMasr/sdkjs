/*
 * (c) Copyright Ascensio System SIA 2010-2024
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation.
 */

"use strict";

// drawingfile.js uses window["AscViewer"] for callbacks and expects window === globalThis in browser.
// In a Worker, we bridge this by aliasing window to self.
self.window = self;

// Namespace must exist before the engine script runs
self["AscViewer"] = self["AscViewer"] || {};

// ------- Fake AscFonts -------------------------------------------------------
// The WASM engine calls AscFonts.pickFont() from within _CheckStreamId.
// In a Worker, the real AscFonts is not available, so we provide a thin shim
// that forwards font requests to the main thread via postMessage.

var g_fakeFonts   = {};  // fontKey -> FakeAscFont
var g_fontStreams  = {};  // fontKey -> { data: Uint8Array, size: number }

function FakeAscFont(name, style) {
    this._name   = name;
    this._style  = style;
    // Use the same "<name>:<style>" key format so fontStreams tracking is consistent.
    this._id     = name + ":" + style;
    this._status = 1;          // 1 = not loaded (same semantics as real AscFont.GetStatus)
    this.externalCallback = undefined;
    this._loadStarted = false;
}
FakeAscFont.prototype.GetID          = function() { return this._id; };
FakeAscFont.prototype.GetStatus      = function() { return this._status; };
FakeAscFont.prototype.GetStreamIndex = function() { return this._id; };  // key into g_fontStreams
FakeAscFont.prototype.LoadFontAsync  = function(basePath, param) {
    if (this._loadStarted) return;
    this._loadStarted = true;
    // Ask the main thread (proxy) to load the font binary and send it back.
    self.postMessage({ type: "needFont", name: this._name, style: this._style, key: this._id });
};

self["AscFonts"] = {
    pickFont: function(name, style) {
        var key = name + ":" + style;
        if (!g_fakeFonts[key])
            g_fakeFonts[key] = new FakeAscFont(name, style);
        return g_fakeFonts[key];
    },
    getFontStream: function(index) {
        return g_fontStreams[index] || null;
    },
    getSymbolRanges: function() { return []; }
};

// ------- Engine readiness ----------------------------------------------------
// drawingfile.js fires window["AscViewer"]["onLoadModule"] from __ATPOSTRUN__.
// Because window === self, we set this callback before importScripts.

var g_engineReady = false;
var g_messages    = [];      // pending command queue (spell.js pattern)

self["AscViewer"]["onLoadModule"] = function() {
    g_engineReady = true;
    _checkMessages();
};

// ------- Message handling (spell.js queue pattern) ---------------------------

self.onmessage = function(e) {
    var data = e.data;

    // Font responses arrive out-of-band and are handled immediately.
    if (data.type === "fontReady") {
        _onFontReady(data);
        return;
    }

    // For render requests: evict any earlier pending render for the same page.
    if (data.type === "render") {
        for (var i = g_messages.length - 1; i >= 0; i--) {
            if (g_messages[i].type === "render" && g_messages[i].pageIndex === data.pageIndex) {
                self.postMessage({ type: "render", id: g_messages[i].id, buffer: null });
                g_messages.splice(i, 1);
            }
        }
    }

    g_messages.push(data);
    if (g_messages.length > 1)
        return;  // already processing; _checkMessages will loop
    _checkMessages();
};

function _checkMessages() {
    if (!g_messages.length || !g_engineReady) return;
    _processMessage(g_messages[0]);
}

function _processMessage(m) {
    switch (m.type) {
        case "open":   _doOpen(m);   break;
        case "render": _doRender(m); break;
        default: break;
    }
    g_messages.shift();
    setTimeout(_checkMessages, 1);
}

// ------- Commands ------------------------------------------------------------

function _doOpen(m) {
    var CDrawingFile = self["AscViewer"]["CDrawingFile"];
    if (!CDrawingFile) {
        self.postMessage({ type: "open", id: m.id, error: -1 });
        return;
    }

    var df   = new CDrawingFile();
    var error = m.password
        ? df["loadFromDataWithPassword"](m.password)
        : df["loadFromData"](m.buffer);

    if (error === 0 || error === 4) {
        // Wire repaint callbacks so font-load completions propagate to main thread.
        df["onRepaintPages"] = function(pages) {
            self.postMessage({ type: "repaintPages", pages: pages });
        };
        df["onRepaintAnnotations"] = function(pages) {
            self.postMessage({ type: "repaintAnnotations", pages: pages });
        };
        df["onRepaintForms"] = function(pages) {
            self.postMessage({ type: "repaintForms", pages: pages });
        };
    }

    self.postMessage({
        type  : "open",
        id    : m.id,
        error : error,
        pages : (error === 0) ? df["getPages"]() : null
    });
}

function _doRender(m) {
    var df = self.drawingFile;
    if (!df) {
        self.postMessage({ type: "render", id: m.id, buffer: null });
        return;
    }

    var ptr = df["getPagePixmap"](m.pageIndex, m.width, m.height, m.backgroundColor);
    if (!ptr) {
        // null means fonts are still loading; repaintPages will retrigger.
        self.postMessage({ type: "render", id: m.id, buffer: null, needFonts: true });
        return;
    }

    // Copy pixel data out of WASM heap into a transferable ArrayBuffer.
    var srcPixels  = df["getUint8ClampedArray"](ptr, 4 * m.width * m.height);
    var outBuffer  = new Uint8ClampedArray(srcPixels.length);
    outBuffer.set(srcPixels);
    df["free"](ptr);

	console.log('printed');
	
    self.postMessage(
        { type: "render", id: m.id, buffer: outBuffer.buffer, width: m.width, height: m.height },
        [outBuffer.buffer]
    );
}

// ------- Font arrival --------------------------------------------------------

function _onFontReady(data) {
    var fakeFont = g_fakeFonts[data.key];
    if (!fakeFont) return;

    if (data.fontData) {
        g_fontStreams[data.key] = {
            data : new Uint8Array(data.fontData),
            size : data.fontData.byteLength
        };
        fakeFont._status = 0;  // mark as loaded
    }

    // externalCallback was set by _CheckStreamId inside drawingfile.js.
    // Calling it triggers fontToMemory() (inside the IIFE, has access to Module)
    // which writes the binary into WASM memory and fires onRepaintPages.
    if (fakeFont.externalCallback)
        fakeFont.externalCallback(fakeFont);

    delete g_fakeFonts[data.key];
}

// ------- Load WASM engine ----------------------------------------------------
// Must be last: importScripts is synchronous, runs the IIFE, which sets up
// CDrawingFile and starts the async WASM fetch.  onLoadModule fires when ready.
importScripts("./drawingfile.js");
