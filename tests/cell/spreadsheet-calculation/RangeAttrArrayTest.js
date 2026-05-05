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

$(function () {
	let CRangeAttrArray = AscCommonExcel.CRangeAttrArray;

	// Naive reference model: a flat array of length maxRow+1.
	function NaiveModel(maxRow) {
		this.maxRow = maxRow;
		this.data = new Array(maxRow + 1);
		for (let i = 0; i <= maxRow; i++) {
			this.data[i] = null;
		}
	}
	NaiveModel.prototype.get = function (r) {
		if (r < 0 || r > this.maxRow) {
			return null;
		}
		return this.data[r];
	};
	NaiveModel.prototype.setRange = function (r1, r2, v) {
		if (r1 < 0) r1 = 0;
		if (r2 > this.maxRow) r2 = this.maxRow;
		for (let i = r1; i <= r2; i++) {
			this.data[i] = (v === undefined) ? null : v;
		}
	};
	NaiveModel.prototype.clearRange = function (r1, r2) {
		this.setRange(r1, r2, null);
	};
	NaiveModel.prototype.insertRows = function (start, count) {
		if (count <= 0) return;
		for (let i = this.maxRow; i >= start + count; i--) {
			this.data[i] = this.data[i - count];
		}
		for (let i = start; i < start + count && i <= this.maxRow; i++) {
			this.data[i] = null;
		}
	};
	NaiveModel.prototype.deleteRows = function (start, count) {
		if (count <= 0) return;
		for (let i = start; i + count <= this.maxRow; i++) {
			this.data[i] = this.data[i + count];
		}
		for (let i = Math.max(start, this.maxRow + 1 - count); i <= this.maxRow; i++) {
			this.data[i] = null;
		}
	};

	function expectEqualToNaive(assert, arr, naive, label) {
		let mismatches = [];
		for (let i = 0; i <= naive.maxRow; i++) {
			let a = arr.get(i);
			let n = naive.get(i);
			if (a !== n) {
				mismatches.push(i + ": got " + a + ", expected " + n);
				if (mismatches.length >= 5) {
					break;
				}
			}
		}
		assert.deepEqual(mismatches, [], label || "values match naive model");
	}

	QUnit.test("Test: empty array basics", function (assert) {
		let a = new CRangeAttrArray(99);
		assert.strictEqual(a.isEmpty(), true);
		assert.strictEqual(a.firstRow(), -1);
		assert.strictEqual(a.lastRow(), -1);
		assert.strictEqual(a.get(0), null);
		assert.strictEqual(a.get(50), null);
		assert.strictEqual(a.get(99), null);
		assert.strictEqual(a.get(-1), null);
		assert.strictEqual(a.get(100), null);
	});

	QUnit.test("Test: setRange single value, get and clear", function (assert) {
		let a = new CRangeAttrArray(99);
		a.setRange(10, 20, 7);
		assert.strictEqual(a.isEmpty(), false);
		assert.strictEqual(a.firstRow(), 10);
		assert.strictEqual(a.lastRow(), 20);
		assert.strictEqual(a.get(9), null);
		assert.strictEqual(a.get(10), 7);
		assert.strictEqual(a.get(15), 7);
		assert.strictEqual(a.get(20), 7);
		assert.strictEqual(a.get(21), null);
		assert.strictEqual(a.getRangeCount(), 1);

		a.clearRange(15, 18);
		assert.strictEqual(a.get(14), 7);
		assert.strictEqual(a.get(15), null);
		assert.strictEqual(a.get(18), null);
		assert.strictEqual(a.get(19), 7);
		assert.strictEqual(a.getRangeCount(), 2);

		a.clearRange(0, 99);
		assert.strictEqual(a.isEmpty(), true);
	});

	QUnit.test("Test: setRange merges adjacent same-value runs", function (assert) {
		let a = new CRangeAttrArray(99);
		a.setRange(0, 4, 1);
		a.setRange(5, 9, 1);
		assert.strictEqual(a.getRangeCount(), 1, "adjacent same value should merge");
		assert.strictEqual(a.firstRow(), 0);
		assert.strictEqual(a.lastRow(), 9);

		a.setRange(10, 14, 2);
		assert.strictEqual(a.getRangeCount(), 2, "different value should not merge");
		a.setRange(5, 14, 1);
		assert.strictEqual(a.getRangeCount(), 1, "overwrite then merge");
		assert.strictEqual(a.lastRow(), 14);
	});

	QUnit.test("Test: setRange overwrites and splits existing range", function (assert) {
		let a = new CRangeAttrArray(99);
		a.setRange(0, 30, 1);
		a.setRange(10, 20, 2);
		assert.strictEqual(a.getRangeCount(), 3);
		assert.strictEqual(a.get(9), 1);
		assert.strictEqual(a.get(10), 2);
		assert.strictEqual(a.get(20), 2);
		assert.strictEqual(a.get(21), 1);
		assert.strictEqual(a.get(30), 1);

		a.setRange(10, 20, 1);
		assert.strictEqual(a.getRangeCount(), 1, "writing same value back collapses splits");
		assert.strictEqual(a.firstRow(), 0);
		assert.strictEqual(a.lastRow(), 30);
	});

	QUnit.test("Test: clearRange splits and merges correctly", function (assert) {
		let a = new CRangeAttrArray(99);
		a.setRange(0, 30, 1);
		a.clearRange(10, 20);
		assert.strictEqual(a.get(9), 1);
		assert.strictEqual(a.get(10), null);
		assert.strictEqual(a.get(20), null);
		assert.strictEqual(a.get(21), 1);
		assert.strictEqual(a.getRangeCount(), 2);

		// clearing an already-empty middle keeps the split intact
		a.clearRange(12, 18);
		assert.strictEqual(a.getRangeCount(), 2);

		// clear completely
		a.clearRange(0, 99);
		assert.strictEqual(a.isEmpty(), true);
	});

	QUnit.test("Test: setRange clamps to [0..maxRow]", function (assert) {
		let a = new CRangeAttrArray(99);
		a.setRange(-50, 5, 1);
		assert.strictEqual(a.firstRow(), 0);
		assert.strictEqual(a.lastRow(), 5);

		a.setRange(95, 1000, 2);
		assert.strictEqual(a.lastRow(), 99);

		// fully outside
		a.setRange(-10, -5, 9);
		a.setRange(120, 200, 9);
		// should not introduce any range outside bounds
		for (let i = 0; i <= 99; i++) {
			let v = a.get(i);
			if (i <= 5) assert.strictEqual(v, 1);
			else if (i >= 95) assert.strictEqual(v, 2);
			else assert.strictEqual(v, null);
		}
	});

	QUnit.test("Test: insertRows shifts and splits", function (assert) {
		let a = new CRangeAttrArray(99);
		a.setRange(10, 20, 1);
		a.insertRows(15, 3);
		// rows 10..14 keep value, rows 15..17 are newly empty, rows 18..23 hold the shifted tail.
		assert.strictEqual(a.get(10), 1);
		assert.strictEqual(a.get(14), 1);
		assert.strictEqual(a.get(15), null);
		assert.strictEqual(a.get(16), null);
		assert.strictEqual(a.get(17), null);
		assert.strictEqual(a.get(18), 1);
		assert.strictEqual(a.get(23), 1);
		assert.strictEqual(a.get(24), null);
		assert.strictEqual(a.getRangeCount(), 2);
	});

	QUnit.test("Test: insertRows fully before existing range shifts entirely", function (assert) {
		let a = new CRangeAttrArray(99);
		a.setRange(20, 30, 1);
		a.insertRows(0, 5);
		assert.strictEqual(a.firstRow(), 25);
		assert.strictEqual(a.lastRow(), 35);
	});

	QUnit.test("Test: insertRows clamps tail past maxRow", function (assert) {
		let a = new CRangeAttrArray(99);
		a.setRange(80, 90, 1);
		a.insertRows(85, 20);
		// 80..84 stay; 105..110 clamped away; tail beyond 99 dropped.
		assert.strictEqual(a.get(80), 1);
		assert.strictEqual(a.get(84), 1);
		assert.strictEqual(a.lastRow(), 84);
	});

	QUnit.test("Test: deleteRows removes, shifts, and merges across gap", function (assert) {
		let a = new CRangeAttrArray(99);
		a.setRange(10, 30, 1);
		a.deleteRows(15, 5);
		// Range [10..30] becomes [10..14] (kept) + [10..(30-5)=25] (shifted from rows 20..30).
		// Adjacent same-value merge -> single [10..25].
		assert.strictEqual(a.firstRow(), 10);
		assert.strictEqual(a.lastRow(), 25);
		assert.strictEqual(a.getRangeCount(), 1);
	});

	QUnit.test("Test: deleteRows wholly inside range", function (assert) {
		let a = new CRangeAttrArray(99);
		a.setRange(0, 5, 1);
		a.setRange(20, 25, 2);
		a.deleteRows(10, 5);
		// First range untouched; second range shifts left by 5.
		assert.strictEqual(a.get(0), 1);
		assert.strictEqual(a.get(5), 1);
		assert.strictEqual(a.get(15), 2);
		assert.strictEqual(a.get(20), 2);
		assert.strictEqual(a.get(21), null);
	});

	QUnit.test("Test: deleteRows entirely consumes a range", function (assert) {
		let a = new CRangeAttrArray(99);
		a.setRange(10, 20, 1);
		a.setRange(30, 40, 2);
		a.deleteRows(10, 11); // wipes [10..20]
		assert.strictEqual(a.get(0), null);
		// [30..40] shifts to [19..29].
		assert.strictEqual(a.get(19), 2);
		assert.strictEqual(a.get(29), 2);
		assert.strictEqual(a.get(30), null);
	});

	QUnit.test("Test: setAreaByRow replicates one source row", function (assert) {
		let a = new CRangeAttrArray(99);
		a.setRange(5, 5, 42);
		a.setAreaByRow(5, 10, 4);
		assert.strictEqual(a.get(10), 42);
		assert.strictEqual(a.get(13), 42);
		assert.strictEqual(a.get(14), null);

		// from-row with no value -> destination cleared.
		a.setRange(50, 60, 7);
		a.setAreaByRow(99, 50, 11);
		for (let i = 50; i <= 60; i++) {
			assert.strictEqual(a.get(i), null);
		}
	});

	QUnit.test("Test: copyFrom between two arrays", function (assert) {
		let src = new CRangeAttrArray(99);
		src.setRange(5, 10, 1);
		src.setRange(20, 25, 2);

		let dst = new CRangeAttrArray(99);
		dst.setRange(50, 60, 9); // pre-existing data inside target window
		dst.copyFrom(src, 0, 40, 30); // copy src rows [0..29] into dst rows [40..69]

		// dst rows [45..50] = 1, [60..65] = 2; everything else in [40..69] is null.
		for (let i = 40; i <= 69; i++) {
			let v = dst.get(i);
			if (i >= 45 && i <= 50) assert.strictEqual(v, 1, "i=" + i);
			else if (i >= 60 && i <= 65) assert.strictEqual(v, 2, "i=" + i);
			else assert.strictEqual(v, null, "i=" + i);
		}
		// outside the destination window: untouched, but pre-existing 50..60 was inside
		// the cleared window, so it's gone.
		assert.strictEqual(dst.get(39), null);
		assert.strictEqual(dst.get(70), null);
	});

	QUnit.test("Test: copyFrom from same array (overlap-safe)", function (assert) {
		let a = new CRangeAttrArray(99);
		a.setRange(0, 9, 1);
		a.copyFrom(a, 0, 5, 10);
		// rows [5..14] now equal old rows [0..9] = 1; rows [0..4] preserved (= 1).
		for (let i = 0; i <= 14; i++) {
			assert.strictEqual(a.get(i), 1, "i=" + i);
		}
		assert.strictEqual(a.get(15), null);
	});

	QUnit.test("Test: iter and iterReverse forward/backward", function (assert) {
		let a = new CRangeAttrArray(99);
		a.setRange(5, 10, 1);
		a.setRange(20, 30, 2);
		a.setRange(40, 45, 3);

		let fwd = [];
		a.iter(0, 99, function (lo, hi, v) {
			fwd.push([lo, hi, v]);
		});
		assert.deepEqual(fwd, [[5, 10, 1], [20, 30, 2], [40, 45, 3]]);

		let rev = [];
		a.iterReverse(0, 99, function (lo, hi, v) {
			rev.push([lo, hi, v]);
		});
		assert.deepEqual(rev, [[40, 45, 3], [20, 30, 2], [5, 10, 1]]);

		// clamped iter
		let part = [];
		a.iter(8, 25, function (lo, hi, v) {
			part.push([lo, hi, v]);
		});
		assert.deepEqual(part, [[8, 10, 1], [20, 25, 2]]);

		// callback returning false aborts
		let count = 0;
		a.iter(0, 99, function () {
			count++;
			return false;
		});
		assert.strictEqual(count, 1);
	});

	QUnit.test("Test: mapRange transforms or clears", function (assert) {
		let a = new CRangeAttrArray(99);
		a.setRange(0, 9, 1);
		a.setRange(10, 19, 2);
		a.setRange(20, 29, 3);

		// Double values that are <= 2; clear value 3.
		a.mapRange(0, 99, function (v) {
			if (v === 3) return null;
			if (v <= 2) return v * 10;
			return v;
		});
		assert.strictEqual(a.get(0), 10);
		assert.strictEqual(a.get(9), 10);
		assert.strictEqual(a.get(10), 20);
		assert.strictEqual(a.get(19), 20);
		assert.strictEqual(a.get(20), null);
		assert.strictEqual(a.get(29), null);
	});

	QUnit.test("Test: clone produces independent copy", function (assert) {
		let a = new CRangeAttrArray(99);
		a.setRange(0, 5, 1);
		a.setRange(10, 15, 2);
		let b = a.clone();
		assert.strictEqual(b.get(0), 1);
		assert.strictEqual(b.get(10), 2);

		b.setRange(0, 99, 9);
		assert.strictEqual(a.get(0), 1, "original unaffected");
		assert.strictEqual(a.get(10), 2);
		assert.strictEqual(b.get(50), 9);
	});

	// Property test: random sequence of operations vs naive model.
	QUnit.test("Test: random property check vs naive model", function (assert) {
		let MAX = 199;
		let a = new CRangeAttrArray(MAX);
		let n = new NaiveModel(MAX);

		// Deterministic LCG so failures are reproducible.
		let state = 0xC0FFEE;
		function rnd() {
			state = (state * 1664525 + 1013904223) >>> 0;
			return state;
		}
		function randInt(lo, hi) { // inclusive
			return lo + (rnd() % (hi - lo + 1));
		}

		for (let step = 0; step < 400; step++) {
			let op = rnd() % 5;
			let r1 = randInt(0, MAX);
			let r2 = randInt(r1, MAX);
			if (op === 0) {
				let v = randInt(1, 4);
				a.setRange(r1, r2, v);
				n.setRange(r1, r2, v);
			} else if (op === 1) {
				a.clearRange(r1, r2);
				n.clearRange(r1, r2);
			} else if (op === 2) {
				let count = randInt(1, 5);
				let start = randInt(0, MAX);
				a.insertRows(start, count);
				n.insertRows(start, count);
			} else if (op === 3) {
				let count = randInt(1, 5);
				let start = randInt(0, MAX);
				a.deleteRows(start, count);
				n.deleteRows(start, count);
			} else {
				let v = randInt(1, 4);
				let from = randInt(0, MAX);
				let to = randInt(0, MAX);
				let cnt = randInt(1, 5);
				a.setRange(from, from, v);
				n.setRange(from, from, v);
				a.setAreaByRow(from, to, cnt);
				// emulate setAreaByRow on the naive model
				let src = n.get(from);
				if (src === null) {
					n.clearRange(to, to + cnt - 1);
				} else {
					n.setRange(to, to + cnt - 1, src);
				}
			}
		}
		expectEqualToNaive(assert, a, n, "after random ops");

		// Invariant: ranges sorted, non-overlapping, no adjacent same-value runs.
		let prev = null;
		for (let i = 0; i < a._ranges.length; i++) {
			let r = a._ranges[i];
			assert.ok(r.r1 <= r.r2, "valid range at " + i);
			if (prev) {
				assert.ok(prev.r2 < r.r1, "non-overlapping at " + i);
				if (prev.r2 + 1 === r.r1) {
					assert.notStrictEqual(prev.v, r.v, "adjacent same-value should have been merged at " + i);
				}
			}
			prev = r;
		}
	});

	QUnit.test("Test: exhaustive small cases for setRange", function (assert) {
		// Iterate over all triples of writes within a tiny board and compare to naive.
		const MAX = 5;
		const VALUES = [1, 2];
		let total = 0;
		for (let s1 = 0; s1 <= MAX; s1++) {
			for (let e1 = s1; e1 <= MAX; e1++) {
				for (let v1 = 0; v1 < VALUES.length; v1++) {
					for (let s2 = 0; s2 <= MAX; s2++) {
						for (let e2 = s2; e2 <= MAX; e2++) {
							for (let v2 = 0; v2 < VALUES.length + 1; v2++) {
								let a = new CRangeAttrArray(MAX);
								let n = new NaiveModel(MAX);
								a.setRange(s1, e1, VALUES[v1]);
								n.setRange(s1, e1, VALUES[v1]);
								if (v2 < VALUES.length) {
									a.setRange(s2, e2, VALUES[v2]);
									n.setRange(s2, e2, VALUES[v2]);
								} else {
									a.clearRange(s2, e2);
									n.clearRange(s2, e2);
								}
								for (let i = 0; i <= MAX; i++) {
									if (a.get(i) !== n.get(i)) {
										assert.ok(false, "exhaustive mismatch s1=" + s1 + " e1=" + e1 + " v1=" + v1 +
											" s2=" + s2 + " e2=" + e2 + " v2=" + v2 + " row=" + i);
										return;
									}
								}
								total++;
							}
						}
					}
				}
			}
		}
		assert.ok(total > 0, "exhaustive scenarios checked: " + total);
	});
});
