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

	window["AscCommon"] = window['AscCommon'] || {};

	function GIFDataBase() {

	}
	window["AscCommon"].GIFDataBase = GIFDataBase;
	GIFDataBase.prototype.getWidth = function() {
		return 0;
	};
	GIFDataBase.prototype.getHeight = function() {
		return 0;
	};
	GIFDataBase.prototype.getFramesCount = function() {
		return 0;
	};
	GIFDataBase.prototype.getFrame = function(frameIndex) {
		return null;
	};
	GIFDataBase.prototype.getFrameDelayMs = function(frameIndex) {
		return 0;
	};
	GIFDataBase.prototype.getDurationMs = function() {
		return 0;
	};
	GIFDataBase.prototype.getLoopCount = function() {
		return 0;
	};
	GIFDataBase.prototype.getFrameTimeMs = function(frameIndex) {
		return 0;
	};
	GIFDataBase.prototype.getFrameIndexAtTime = function(timeMs) {
		return 0;
	};

	function GIFPlayer(gifData, viewPortData) {
		this.gifData = gifData;
		this.viewPortData = viewPortData;
		this.lastDrawFrame = -1;

		this.startTime = null;
	}
	GIFPlayer.prototype.isStarted = function() {
		return this.startTime !== null;
	};
	GIFPlayer.prototype.getCurrentFrameIndex = function() {
		if (!this.isStarted()) {
			return -1;
		}
		let currentTime = (new Date()).getTime();
		let elapsedTime = currentTime - this.startTime;
		return this.gifData.getFrameIndexAtTime(elapsedTime);
	}
	GIFPlayer.prototype.onTick = function() {
		if (!this.isStarted()) {
			return;
		}
		if (!this.viewPortData.isVisible()) {
			return;
		}
		let currentTime = (new Date()).getTime();
		let elapsedTime = currentTime - this.startTime;
		let currentFrame = this.gifData.getFrameIndexAtTime(elapsedTime);
		if (this.lastDrawFrame !== currentFrame) {
			this.viewPortData.onUpdateFrame();
		}
	};
	GIFPlayer.prototype.getFrameForDraw = function() {
		this.lastDrawFrame = this.getCurrentFrameIndex();
		return this.gifData.getFrame(this.lastDrawFrame);
	};
	window["AscCommon"].GifPlayer = GIFPlayer;


	function GIFDataGIFuct(gifuctData) {
		AscCommon.GIFDataBase.call(this);
		this.gifuctData = gifuctData;
		this.frames = this.gifuctData.decompressFrames(true);
	}
	GIFDataGIFuct.prototype = Object.create(AscCommon.GIFDataBase.prototype);
	GIFDataGIFuct.prototype.getWidth = function() {
		return this.gifuctData.lsd.width;
	};
	GIFDataGIFuct.prototype.getHeight = function() {
		return this.gifuctData.lsd.height;
	};
	GIFDataGIFuct.prototype.getFramesCount = function() {
		return this.gifuctData.decompressFrames(false).length;
	};
	GIFDataGIFuct.prototype.getFrame = function(frameIndex) {
		if (frameIndex < 0 || frameIndex >= this.frames.length) {
			return null;
		}
		return this.frames[frameIndex];
	};
	GIFDataGIFuct.prototype.getFrameDelayMs = function(frameIndex) {
		let frame = this.getFrame(frameIndex);
		if (!frame || !frame.delay) {
			return 0;
		}
		return frame.delay * 10;
	};
	GIFDataGIFuct.prototype.getDurationMs = function() {
		let totalDuration = 0;
		let framesCount = this.getFramesCount();
		for (let i = 0; i < framesCount; i++) {
			totalDuration += this.getFrameDelayMs(i);
		}
		return totalDuration;
	};
	GIFDataGIFuct.prototype.getLoopCount = function() {
		let appExt = this.gifuctData.appExtensions.find(function(ext) {
			return ext.identifier === "NETSCAPE" && ext.authCode === "2.0";
		});
		if (appExt && appExt.loopCount !== undefined) {
			return appExt.loopCount;
		}
		return 0;
	};
	GIFDataGIFuct.prototype.getFrameTimeMs = function(frameIndex) {
		let timeMs = 0;
		for (let i = 0; i < frameIndex; i++) {
			timeMs += this.getFrameDelayMs(i);
		}
		return timeMs;
	};
	GIFDataGIFuct.prototype.getFrameIndexAtTime = function(timeMs) {
		let durationMs = this.getDurationMs();
		if (durationMs === 0) {
			return 0;
		}
		let loopCount = this.getLoopCount();
		if (loopCount === 0) {
			timeMs = timeMs % durationMs;
		} else {
			let totalLoopDuration = durationMs * loopCount;
			if (timeMs >= totalLoopDuration) {
				return this.getFramesCount() - 1;
			}
			timeMs = timeMs % durationMs;
		}
		let accumulatedTime = 0;
		let framesCount = this.getFramesCount();
		for (let i = 0; i < framesCount; i++) {
			accumulatedTime += this.getFrameDelayMs(i);
			if (timeMs < accumulatedTime) {
				return i;
			}
		}
		return framesCount - 1;
	};
	window["AscCommon"].GIFDataGIFuct = GIFDataGIFuct;


})(window);
