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

(function(undefined) {
	
	function OutlineSlide() {
		this.title = null;
		this.subTitle = null;
		this.content = [];
		this.slide = null;
	}
	OutlineSlide.prototype.setSlide = function (pr) {
		this.slide = pr;
	};
	OutlineSlide.prototype.setTitle = function (pr) {
		this.title = pr;
	};
	OutlineSlide.prototype.setSubTitle = function (pr) {
		this.subTitle = pr;
	};
	OutlineSlide.prototype.addContent = function (pr) {
		this.content.push(pr);
	};

	function OutlineView() {
		this.outlineShape = null;
		this.paragraphMap = [];
		this.slideFirstParagraphs = [];
		this.currentSlideIndex = -1;
	}
	OutlineView.prototype.getPresentation = function () {
		return Asc.editor.private_GetLogicDocument();
	};
	OutlineView.prototype.setOutlineShape = function (shape) {
		this.outlineShape = shape;
	};
	OutlineView.prototype.updateAll = function (width, height, currentSlideIndex) {
		this.currentSlideIndex = (currentSlideIndex !== undefined) ? currentSlideIndex : -1;
		const presentation = this.getPresentation();
		const outlineSlides = [];
		for (let i = 0; i < presentation.Slides.length; i += 1) {
			const slide = presentation.Slides[i];
			const outlineSlide = slide.getOutlineSlide();
			outlineSlides.push(outlineSlide);
		}
		const outlineShape = this.getOutlineShape(outlineSlides, width, height);
		this.setOutlineShape(outlineShape);
	};
	OutlineView.prototype.addContentToOutlineShape = function (outlineShape, slideShape, slideIndex, isTitle) {
		if (!slideShape) {
			return;
		}
		const outlineContent = outlineShape.txBody.content;
		const slideContent = slideShape.txBody.content;
		const paragraphs = slideContent.Content;
		const phIdx = slideShape.getPlaceholderIndex ? slideShape.getPlaceholderIndex() : null;
		const shapeOrder = (phIdx !== null && phIdx !== undefined) ? (parseInt(phIdx, 10) || 0) : 0;
		for (let i = 0; i < paragraphs.length; i += 1) {
			const paragraph = paragraphs[i];
			const copyParagraph = this.getCopyParagraph(outlineContent, paragraph, isTitle);
			outlineContent.AddToContent(outlineContent.Content.length, copyParagraph);
			this.paragraphMap.push({
				slideIndex: slideIndex,
				isTitle: isTitle,
				isFirstOfShape: (i === 0),
				shapeOrder: shapeOrder,
				sourceShape: slideShape,
				sourceParagraphIdx: i
			});
		}
	};
	OutlineView.prototype.fillOutlineShape = function (outlineShape, outlineSlides, width) {
		for (let i = 0; i < outlineSlides.length; i += 1) {
			const slide = outlineSlides[i];
			const startIdx = outlineShape.txBody.content.Content.length;
			this.slideFirstParagraphs.push(startIdx);
			this.addContentToOutlineShape(outlineShape, slide.title, i, true);
			if (slide.subTitle) {
				this.addContentToOutlineShape(outlineShape, slide.subTitle, i, false);
			}
			for (let j = 0; j < slide.content.length; j += 1) {
				this.addContentToOutlineShape(outlineShape, slide.content[j], i, false);
			}
		}
	};
	OutlineView.prototype.applyParagraphProps = function (outlineParagraph, slideParagraph, isTitle) {
		const compiledPr = slideParagraph.getCompiledPr();
		const copyParaPr = new CParaPr();
		if (compiledPr.ParaPr.Bullet) {
			copyParaPr.Bullet = compiledPr.ParaPr.Bullet.createDuplicate();
			copyParaPr.Lvl = compiledPr.ParaPr.Lvl;
			copyParaPr.Ind = compiledPr.ParaPr.Ind.Copy();
			// copyParaPr.Ind.FirstLine *= 0.5;
			// copyParaPr.Ind.Left *= 0.5;
			// copyParaPr.Ind.Right *= 0.5;
		}
		outlineParagraph.SetPr(copyParaPr);
		outlineParagraph.CheckRunContent(function (run) {
			const textPr = new CTextPr();
			textPr.SetFontSize(10);
			if (isTitle) {
				textPr.SetBold(true);
			}
			run.SetPr(textPr);
		});
	};
	OutlineView.prototype.getOutlineShape = function (outlineSlides, width) {
		this.paragraphMap = [];
		this.slideFirstParagraphs = [];
		return AscFormat.ExecuteNoHistory(function () {
			const outlineShape = new AscFormat.CShape();
			outlineShape.setBDeleted(false);
			outlineShape.createTextBody();
			outlineShape.txBody.bodyPr.setInsets(0, 0, 0, 0);
			const outlineContent = outlineShape.txBody.content;
			outlineContent.ClearContent(false);
			this.fillOutlineShape(outlineShape, outlineSlides, width);
			outlineShape.extX = width;
			outlineShape.extY = 2000;
			outlineShape.recalculateContent();
			outlineShape.contentWidth = width;
			return outlineShape;
		}, this, []);
	};
	OutlineView.prototype.getCopyParagraph = function (parent, paragraph, isTitle) {
		return AscFormat.ExecuteNoHistory(function () {
			const copyParagraph = paragraph.Copy(parent, null, null);
			this.applyParagraphProps(copyParagraph, paragraph, isTitle);
			return copyParagraph;
		}, this, []);
	};
	OutlineView.prototype.draw = function (graphics) {
		if (this.outlineShape) {
			this.outlineShape.draw(graphics);
		}
	};

	OutlineView.prototype.getParagraphY = function (paragraphIdx) {
		if (!this.outlineShape || !this.outlineShape.txBody) return null;
		const content = this.outlineShape.txBody.content;
		if (!content || !content.Content) return null;
		const paragraph = content.Content[paragraphIdx];
		if (paragraph && paragraph.Pages && paragraph.Pages[0] !== undefined
				&& paragraph.Lines && paragraph.Lines[0]) {
			return paragraph.Pages[0].Y + paragraph.Lines[0].Top;
		}
		return null;
	};

	OutlineView.prototype.getOutlineParagraphAtPosition = function (yMm) {
		if (!this.outlineShape || !this.paragraphMap.length) return null;
		const mapLen = this.paragraphMap.length;
		for (let i = 0; i < mapLen; i += 1) {
			const paraY = this.getParagraphY(i);
			if (paraY === null) continue;
			let nextY;
			if (i + 1 < mapLen) {
				nextY = this.getParagraphY(i + 1);
			}
			if (nextY === null || nextY === undefined) {
				nextY = Infinity;
			}
			if (yMm >= paraY && yMm < nextY) {
				return {
					paragraphIdx: i,
					slideIndex: this.paragraphMap[i].slideIndex,
					isTitle: this.paragraphMap[i].isTitle,
					sourceShape: this.paragraphMap[i].sourceShape,
					sourceParagraphIdx: this.paragraphMap[i].sourceParagraphIdx
				};
			}
		}
		return null;
	};

	OutlineView.prototype.createDecorShape = function (brushRGB, penRGB, w, h, geometryType, penW, label) {
		return AscFormat.ExecuteNoHistory(function () {
			var shape = new AscFormat.CShape();
			shape.setBDeleted(false);
			shape.extX = w;
			shape.extY = h;

			shape.brush = AscFormat.CreateSolidFillRGBA(brushRGB.R, brushRGB.G, brushRGB.B, 255);

			var pen = new AscFormat.CLn();
			pen.Fill = AscFormat.CreateSolidFillRGBA(penRGB.R, penRGB.G, penRGB.B, 255);
			pen.w = penW;
			shape.pen = pen;

			geometryType = geometryType || "rect";
			shape.calcGeometry = AscFormat.CreateGeometry(geometryType);
			shape.calcGeometry.Recalculate(w, h);

			if (label !== null && label !== undefined) {
				var txBody = shape.createTextBody();
				txBody.bodyPr.setAnchor(AscFormat.VERTICAL_ANCHOR_TYPE_CENTER);
				txBody.bodyPr.setInsets(0, 0, 0, 0);

				var para = txBody.content.Content[0];
				var paraPr = new CParaPr();
				paraPr.Jc = AscCommon.align_Center;
				para.SetPr(paraPr);

				var run = new ParaRun(para, false);
				var textPr = new CTextPr();
				textPr.SetFontSize(10);
				run.Set_Pr(textPr);
				run.AddText(label);
				para.AddToContent(0, run);


				shape.recalculateContent();
			}

			return shape;
		}, this, []);
	};

	OutlineView.prototype.drawDecorShape = function (graphics, shape, x, y) {
		var t = new AscCommon.CMatrix();
		t.tx = x;
		t.ty = y;
		shape.draw(graphics, t, t);
	};

	OutlineView.prototype.getOutlineCursorPos = function () {
		if (!this.outlineShape) return null;
		var content = this.outlineShape.txBody.content;
		const curParagraph = content.GetCurrentParagraph();
		return curParagraph && curParagraph.Get_ParaContentPos(false, false);
	};
	OutlineView.prototype.getRGBFromHex = function (color) {
		const prepareColor = parseInt(color.slice(1), 16);
		return {R: (prepareColor >> 16) & 0xff, G: (prepareColor >> 8) & 0xff, B: prepareColor & 0xff};
	};
	OutlineView.prototype.drawDecorations = function (graphics, currentSlideIndex, scrollYMm, leftMarginMM) {
		if (!this.outlineShape || !this.paragraphMap.length) return;

		const rectX = leftMarginMM * 0.1;
		const rectW = leftMarginMM * 0.3;
		const content = this.outlineShape.txBody.content;
		const backgroundRGB = this.getRGBFromHex(AscCommon.GlobalSkin.BackgroundColorThumbnails);
		for (let i = 0; i < this.slideFirstParagraphs.length; i += 1) {
			const firstIdx = this.slideFirstParagraphs[i];
			const topY = this.getParagraphY(firstIdx);
			if (topY === null) continue;

			const firstPara = content.Content[firstIdx];
			let firstLineH = rectW;
			if (firstPara && firstPara.Lines && firstPara.Lines[0]) {
				const fl = firstPara.Lines[0];
				firstLineH = (fl.Bottom || 0) - (fl.Top || 0);
			}
			const barH = rectW * (4 / 5);

			const penRGB = this.getRGBFromHex(i === currentSlideIndex ? AscCommon.GlobalSkin.ThumbnailsPageOutlineActive : AscCommon.GlobalSkin.ThumbnailsPageOutline);
			const barShape = this.createDecorShape(backgroundRGB, penRGB, rectW, barH, "roundRect", 12700);
			this.drawDecorShape(graphics, barShape, rectX, topY - scrollYMm);
		}

		const slideShapeEntries = {};
		for (let i = 0; i < this.paragraphMap.length; i += 1) {
			const entry = this.paragraphMap[i];
			if (!entry.isFirstOfShape || entry.isTitle) continue;
			if (!slideShapeEntries[entry.slideIndex]) {
				slideShapeEntries[entry.slideIndex] = [];
			}
			slideShapeEntries[entry.slideIndex].push({ paragraphIdx: i, shapeOrder: entry.shapeOrder });
		}
		const badgeNumbers = {};
		for (const slideIdx in slideShapeEntries) {
			const shapes = slideShapeEntries[slideIdx];
			shapes.sort(function (a, b) { return a.shapeOrder - b.shapeOrder; });
			for (let j = 0; j < shapes.length; j += 1) {
				badgeNumbers[shapes[j].paragraphIdx] = j + 1;
			}
		}



		const numberPenRGB = this.getRGBFromHex(AscCommon.GlobalSkin.ThumbnailsPageOutline);
		const numberWShape = rectW * (4 / 5);
		for (let i = 0; i < this.paragraphMap.length; i += 1) {
			const entry = this.paragraphMap[i];
			if (!entry.isFirstOfShape || entry.isTitle) continue;

			const paraY = this.getParagraphY(i);
			if (paraY === null) continue;

			const para = content.Content[i];
			let firstLineH = rectW;
			if (para && para.Lines && para.Lines[0]) {
				const fl = para.Lines[0];
				firstLineH = (fl.Bottom || 0) - (fl.Top || 0);
			}
			const badgeH = Math.min(rectW, Math.max(0.1, firstLineH));

			const badgeShape = this.createDecorShape(backgroundRGB, numberPenRGB, numberWShape, badgeH, "rect", 0, String(badgeNumbers[i]));
			this.drawDecorShape(graphics, badgeShape, rectX, paraY - scrollYMm);
		}
	};
	OutlineView.prototype.getTargetDocContent = function () {
		return this.outlineShape && this.outlineShape.txBody.content;
	};
	OutlineView.prototype.selectionSetStart = function (e, x, y, slideIndex) {
		if (this.outlineShape) {
			this.outlineShape.selectionSetStart(e, x, y, slideIndex);
		}
	};
	OutlineView.prototype.updateOutlineShapeTransform = function (offsetX, offsetY) {
		if (!this.outlineShape) {
			return;
		}
		this.outlineShape.transform.tx = offsetX || 0;
		this.outlineShape.transform.ty = offsetY || 0;
		this.outlineShape.transformText.tx = offsetX || 0;
		this.outlineShape.transformText.ty = offsetY || 0;
		this.outlineShape.invertTransform = AscCommon.global_MatrixTransformer.Invert(this.outlineShape.transform);
		this.outlineShape.invertTransformText = AscCommon.global_MatrixTransformer.Invert(this.outlineShape.transformText);
		this.outlineShape.transformText2 = this.outlineShape.transformText;
		this.outlineShape.invertTransformText2 = this.outlineShape.invertTransformText;
	};


	window["AscCommonSlide"] = window["AscCommonSlide"] || {};
	window["AscCommonSlide"].OutlineSlide = OutlineSlide;
	window["AscCommonSlide"].OutlineView = OutlineView;
})();
