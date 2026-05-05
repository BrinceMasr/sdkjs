/*
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
 *
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

"use strict";

(function (window, undefined) {
	var DEFAULT_MAX_ROW = (window['AscCommon'] && window['AscCommon'].gc_nMaxRow0 != null)
		? window['AscCommon'].gc_nMaxRow0
		: 1048575;

	// Stage 2 skeleton for the per-column direct cell style storage on Worksheet.
	// All operations only touch `ws.cellStylesByCol[col]` (a CRangeAttrArray) and
	// never call into Cell, SheetMemory, or History. Nothing in the existing
	// code path calls these methods yet — they exist so Stage 3 (shadow mode)
	// has a stable surface to write through.

	function _getCRangeAttrArrayCtor() {
		return window['AscCommonExcel'].CRangeAttrArray;
	}

	// Normalize an input value into an xfIndex (a non-negative integer).
	// Returns 0 for "no direct cell style". This matches Cell.loadContent,
	// which only restores style when the saved index is > 0.
	function _toXfIndex(xfIndexOrXfs) {
		if (xfIndexOrXfs == null) {
			return 0;
		}
		if (typeof xfIndexOrXfs === 'number') {
			return xfIndexOrXfs > 0 ? (xfIndexOrXfs | 0) : 0;
		}
		if (typeof xfIndexOrXfs.getIndexNumber === 'function') {
			var idx = xfIndexOrXfs.getIndexNumber();
			return (typeof idx === 'number' && idx > 0) ? (idx | 0) : 0;
		}
		// Raw style object — funnel through g_StyleCache to deduplicate.
		var cache = window['AscCommonExcel'].g_StyleCache;
		if (cache && typeof cache.addXf === 'function') {
			var cached = cache.addXf(xfIndexOrXfs);
			if (cached && typeof cached.getIndexNumber === 'function') {
				var idx2 = cached.getIndexNumber();
				return (typeof idx2 === 'number' && idx2 > 0) ? (idx2 | 0) : 0;
			}
		}
		return 0;
	}

	function getCellStyleStore(ws, col, opt_create) {
		if (col < 0) {
			return null;
		}
		var store = ws.cellStylesByCol[col];
		if (!store && opt_create) {
			var Ctor = _getCRangeAttrArrayCtor();
			store = new Ctor(DEFAULT_MAX_ROW);
			ws.cellStylesByCol[col] = store;
		}
		return store || null;
	}

	// Returns 0 for "no direct cell style" so callers can compare with the
	// existing Cell.saveContent / Cell.loadContent semantics directly.
	function getCellXf(ws, row, col) {
		if (col < 0) {
			return 0;
		}
		var store = ws.cellStylesByCol[col];
		if (!store) {
			return 0;
		}
		var idx = store.get(row);
		return (idx == null) ? 0 : idx;
	}

	function setCellXf(ws, row, col, xfIndexOrXfs, opt_options) {
		var idx = _toXfIndex(xfIndexOrXfs);
		if (idx === 0) {
			var existing = ws.cellStylesByCol[col];
			if (existing) {
				existing.clearRange(row, row);
			}
			return;
		}
		var store = getCellStyleStore(ws, col, true);
		store.setRange(row, row, idx);
	}

	function setCellXfRange(ws, r1, c1, r2, c2, xfIndexOrXfs, opt_options) {
		if (c2 < c1 || r2 < r1) {
			return;
		}
		var idx = _toXfIndex(xfIndexOrXfs);
		if (idx === 0) {
			for (var c = c1; c <= c2; c++) {
				var existing = ws.cellStylesByCol[c];
				if (existing) {
					existing.clearRange(r1, r2);
				}
			}
			return;
		}
		for (var c2col = c1; c2col <= c2; c2col++) {
			var store = getCellStyleStore(ws, c2col, true);
			store.setRange(r1, r2, idx);
		}
	}

	function clearCellXfRange(ws, r1, c1, r2, c2, opt_options) {
		if (c2 < c1 || r2 < r1) {
			return;
		}
		for (var c = c1; c <= c2; c++) {
			var existing = ws.cellStylesByCol[c];
			if (existing) {
				existing.clearRange(r1, r2);
			}
		}
	}

	// Copy a 2D range of direct cell styles between worksheets (or within one).
	// fromBBox and toBBox are expected to have matching dimensions; if they
	// differ, the smaller dimension is used so the call is bounded.
	function copyCellXfRange(wsTo, wsFrom, fromBBox, toBBox, opt_options) {
		var fw = fromBBox.c2 - fromBBox.c1 + 1;
		var fh = fromBBox.r2 - fromBBox.r1 + 1;
		var tw = toBBox.c2 - toBBox.c1 + 1;
		var th = toBBox.r2 - toBBox.r1 + 1;
		var w = fw < tw ? fw : tw;
		var h = fh < th ? fh : th;
		if (w <= 0 || h <= 0) {
			return;
		}
		for (var dc = 0; dc < w; dc++) {
			var srcCol = fromBBox.c1 + dc;
			var dstCol = toBBox.c1 + dc;
			var src = wsFrom.cellStylesByCol[srcCol];
			var dst = wsTo.cellStylesByCol[dstCol];
			if (!src) {
				if (dst) {
					dst.clearRange(toBBox.r1, toBBox.r1 + h - 1);
				}
				continue;
			}
			if (!dst) {
				dst = getCellStyleStore(wsTo, dstCol, true);
			}
			dst.copyFrom(src, fromBBox.r1, toBBox.r1, h);
		}
	}

	// Shift cell styles inside / through `bbox` in one of four directions.
	// `offset` is currently informational; `mode` drives the behavior.
	// Modes:
	//   'up'    — delete the bbox rows in cols [c1..c2]; rows below shift up.
	//   'down'  — insert empty rows of bbox height in cols [c1..c2] at r1.
	//   'left'  — within rows [r1..r2], pull style from columns to the right of c2 into [c1..]; clear the gap at the far right.
	//   'right' — within rows [r1..r2], shift style in cols starting at c1 by +width; clear the inserted block at [c1..c1+width-1].
	function shiftCellXfs(ws, bbox, offset, mode) {
		var c1 = bbox.c1;
		var c2 = bbox.c2;
		var r1 = bbox.r1;
		var r2 = bbox.r2;
		if (c2 < c1 || r2 < r1) {
			return;
		}
		var w = c2 - c1 + 1;
		var h = r2 - r1 + 1;

		if (mode === 'up') {
			for (var c = c1; c <= c2; c++) {
				var sUp = ws.cellStylesByCol[c];
				if (sUp) {
					sUp.deleteRows(r1, h);
				}
			}
			return;
		}
		if (mode === 'down') {
			for (var cd = c1; cd <= c2; cd++) {
				var sDown = ws.cellStylesByCol[cd];
				if (sDown) {
					sDown.insertRows(r1, h);
				}
			}
			return;
		}
		if (mode === 'left') {
			var L = ws.cellStylesByCol.length;
			for (var cl = c1; cl < L; cl++) {
				var srcL = ws.cellStylesByCol[cl + w];
				var dstL = ws.cellStylesByCol[cl];
				if (!srcL) {
					if (dstL) {
						dstL.clearRange(r1, r2);
					}
					continue;
				}
				if (!dstL) {
					dstL = getCellStyleStore(ws, cl, true);
				}
				dstL.copyFrom(srcL, r1, r1, h);
			}
			return;
		}
		if (mode === 'right') {
			// The rightmost source index is Lr-1; after shifting, it lands at Lr-1+w,
			// so iterate the destination range up to that point. Indices past the
			// current array length are materialized via getCellStyleStore.
			var Lr = ws.cellStylesByCol.length;
			for (var cr = Lr - 1 + w; cr >= c1 + w; cr--) {
				var srcR = ws.cellStylesByCol[cr - w];
				var dstR = ws.cellStylesByCol[cr];
				if (!srcR) {
					if (dstR) {
						dstR.clearRange(r1, r2);
					}
					continue;
				}
				if (!dstR) {
					dstR = getCellStyleStore(ws, cr, true);
				}
				dstR.copyFrom(srcR, r1, r1, h);
			}
			for (var cg = c1; cg < c1 + w; cg++) {
				var dstGap = ws.cellStylesByCol[cg];
				if (dstGap) {
					dstGap.clearRange(r1, r2);
				}
			}
			return;
		}
	}

	// Iterate every style run intersecting bbox in column-major order.
	// callback(loRow, hiRow, col, xfIndex). Return false from the callback to stop.
	function iterCellXfs(ws, bbox, callback) {
		var c1 = bbox.c1;
		var c2 = bbox.c2;
		if (c2 < c1) {
			return;
		}
		for (var c = c1; c <= c2; c++) {
			var store = ws.cellStylesByCol[c];
			if (!store) {
				continue;
			}
			var stopped = false;
			var colCaptured = c;
			store.iter(bbox.r1, bbox.r2, function (lo, hi, v) {
				if (callback(lo, hi, colCaptured, v) === false) {
					stopped = true;
					return false;
				}
			});
			if (stopped) {
				return;
			}
		}
	}

	// Stage 3 shadow mirror entry point.
	// Called from Cell.setStyleInternal (and Cell.clearDataKeepXf, which bypasses
	// setStyleInternal). Mirrors the in-memory `cell.xfs` into the worksheet's
	// per-column range storage so cellStylesByCol stays in sync with the
	// authoritative SheetMemory path. Skipped for transient cells, for cells
	// without a real (row, col), and for cells without a Worksheet host.
	function mirrorCellStyle(cell) {
		if (!cell || cell._isTransient) {
			return;
		}
		if (cell.nRow < 0 || cell.nCol < 0) {
			return;
		}
		var ws = cell.ws;
		if (!ws || !ws.cellStylesByCol) {
			return;
		}
		setCellXf(ws, cell.nRow, cell.nCol, cell.xfs);
	}

	// Stage 4 writer entry point.
	// Resolves the direct-cell xfIndex for a serializer in the order:
	//   1. Worksheet.cellStylesByCol — the new primary read source.
	//   2. cell.xfs (from SheetMemory via loadContent) — the safety shadow
	//      that is still being kept in sync until Stage 8.
	// During Stages 4–7 both stores are equivalent for any cell that has been
	// touched by setStyleInternal; the fallback covers cells that were loaded
	// from disk and never mutated, where cellStylesByCol has not been
	// populated by the shadow mirror yet.
	function getWriterCellXfIndex(ws, row, col, cell) {
		if (ws && row >= 0 && col >= 0) {
			var idx = getCellXf(ws, row, col);
			if (idx > 0) {
				return idx;
			}
		}
		return (cell && cell.xfs) ? cell.xfs.getIndexNumber() : 0;
	}

	// Same as getWriterCellXfIndex but returns the CellXfs object suitable
	// for stylesForWrite.add(...). Returns null when there is no direct style.
	function getWriterCellXfs(ws, row, col, cell) {
		var idx = getWriterCellXfIndex(ws, row, col, cell);
		if (idx <= 0) {
			return null;
		}
		var cache = window['AscCommonExcel'].g_StyleCache;
		return cache ? cache.getXf(idx) : null;
	}

	function installOnWorksheet(WorksheetCtor) {
		WorksheetCtor.prototype.getCellStyleStore = function (col, opt_create) {
			return getCellStyleStore(this, col, opt_create);
		};
		WorksheetCtor.prototype.getCellXf = function (row, col) {
			return getCellXf(this, row, col);
		};
		WorksheetCtor.prototype.setCellXf = function (row, col, xfIndexOrXfs, opt_options) {
			return setCellXf(this, row, col, xfIndexOrXfs, opt_options);
		};
		WorksheetCtor.prototype.setCellXfRange = function (r1, c1, r2, c2, xfIndexOrXfs, opt_options) {
			return setCellXfRange(this, r1, c1, r2, c2, xfIndexOrXfs, opt_options);
		};
		WorksheetCtor.prototype.clearCellXfRange = function (r1, c1, r2, c2, opt_options) {
			return clearCellXfRange(this, r1, c1, r2, c2, opt_options);
		};
		WorksheetCtor.prototype.copyCellXfRange = function (wsFrom, fromBBox, toBBox, opt_options) {
			return copyCellXfRange(this, wsFrom, fromBBox, toBBox, opt_options);
		};
		WorksheetCtor.prototype.shiftCellXfs = function (bbox, offset, mode) {
			return shiftCellXfs(this, bbox, offset, mode);
		};
		WorksheetCtor.prototype.iterCellXfs = function (bbox, callback) {
			return iterCellXfs(this, bbox, callback);
		};
	}

	window['AscCommonExcel'] = window['AscCommonExcel'] || {};
	window['AscCommonExcel'].CellStyleStorage = {
		getCellStyleStore: getCellStyleStore,
		getCellXf: getCellXf,
		setCellXf: setCellXf,
		setCellXfRange: setCellXfRange,
		clearCellXfRange: clearCellXfRange,
		copyCellXfRange: copyCellXfRange,
		shiftCellXfs: shiftCellXfs,
		iterCellXfs: iterCellXfs,
		mirrorCellStyle: mirrorCellStyle,
		getWriterCellXfIndex: getWriterCellXfIndex,
		getWriterCellXfs: getWriterCellXfs,
		installOnWorksheet: installOnWorksheet,
		_toXfIndex: _toXfIndex
	};
})(window);
