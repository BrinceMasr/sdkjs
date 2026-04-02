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
		this.outlineToSourceMap = {};
		this.sourceToOutlineMap = {};
		this.outlineInfo = {};
	}
	OutlineView.prototype.reset = function () {
		this.outlineToSourceMap = {};
		this.sourceToOutlineMap = {};
		this.outlineInfo = {};
	};
	OutlineView.prototype.getPresentation = function () {
		return Asc.editor.private_GetLogicDocument();
	};
	OutlineView.prototype.setOutlineShape = function (shape) {
		this.outlineShape = shape;
	};
	OutlineView.prototype.updateAll = function (width, height) {
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
	OutlineView.prototype.addOutlineParagraph = function (sourceParagraph, outlineParagraph, pr) {
		const outlineId = outlineParagraph.Get_Id();
		this.outlineToSourceMap[outlineId] = sourceParagraph;
		this.sourceToOutlineMap[sourceParagraph.Get_Id()] = outlineParagraph;
		if (pr) {
			this.outlineInfo[outlineId] = {};
			if (pr.titleShapeIndex !== undefined) {
				this.outlineInfo[outlineId] = {titleShapeIndex: pr.titleShapeIndex};
			}
			if (pr.contentShapeIndex !== undefined) {
				this.outlineInfo[outlineId] = {contentShapeIndex: pr.contentShapeIndex};
			}
		}
	};
	OutlineView.prototype.addContentToOutlineShape = function (outlineShape, slideShape, slideIndex, pr) {
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
			const copyParagraph = this.getCopyParagraph(outlineContent, paragraph, !!pr && pr.titleShapeIndex !== undefined);
			outlineContent.AddToContent(outlineContent.Content.length, copyParagraph);
			this.addOutlineParagraph(paragraph, copyParagraph, i === 0 ? pr : null);
		}
	};
	OutlineView.prototype.fillOutlineShape = function (outlineShape, outlineSlides, width) {
		for (let i = 0; i < outlineSlides.length; i += 1) {
			const slide = outlineSlides[i];

			this.addContentToOutlineShape(outlineShape, slide.title, i, {titleShapeIndex: i});
			let startShapeIndex = 0;
			if (slide.subTitle) {
				startShapeIndex = 1;
				this.addContentToOutlineShape(outlineShape, slide.subTitle, i, {contentShapeIndex: 0});
			}
			for (let j = 0; j < slide.content.length; j += 1) {
				this.addContentToOutlineShape(outlineShape, slide.content[j], i, {contentShapeIndex: startShapeIndex + j});
			}
		}
	};
	OutlineView.prototype.getTextPr = function (isTitle) {
		const textPr = new CTextPr();
		textPr.SetFontSize(10);
		if (isTitle) {
			textPr.SetBold(true);
		}
		return textPr;
	};
	OutlineView.prototype.getParaPr = function (compiledParaPr) {
		const copyParaPr = new CParaPr();
		if (compiledParaPr.ParaPr.Bullet) {
			copyParaPr.Bullet = compiledParaPr.ParaPr.Bullet.createDuplicate();
			copyParaPr.Lvl = compiledParaPr.ParaPr.Lvl;
		}
		copyParaPr.Ind = compiledParaPr.ParaPr.Ind.Copy();
		copyParaPr.Ind.FirstLine *= 0.5;
		copyParaPr.Ind.Left *= 0.5;
		copyParaPr.Ind.Right *= 0.5;
		copyParaPr.Spacing.Before = 1;
		copyParaPr.Spacing.After = 1;
		copyParaPr.Spacing.Line = 1;
		return copyParaPr;
	};
	OutlineView.prototype.applyParagraphProps = function (outlineParagraph, slideParagraph, isTitle) {
		const compiledPr = slideParagraph.getCompiledPr();
		const copyParaPr = this.getParaPr(compiledPr);
		const oThis = this;
		outlineParagraph.SetPr(copyParaPr);
		outlineParagraph.CheckRunContent(function (run) {
			const textPr = oThis.getTextPr(isTitle);
			run.SetPr(textPr);
		});
	};
	OutlineView.prototype.getOutlineShape = function (outlineSlides, width) {
		this.reset();
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

	OutlineView.prototype.getRGBFromHex = function (color) {
		const prepareColor = parseInt(color.slice(1), 16);
		return {R: (prepareColor >> 16) & 0xff, G: (prepareColor >> 8) & 0xff, B: prepareColor & 0xff};
	};
	OutlineView.prototype.getOutlineParagraphs = function () {
		const content = this.getDocContent();
		if (content) {
			return content.Content;
		}
		return [];
	};
	OutlineView.prototype.isHaveParagraphs = function () {
		return this.getOutlineParagraphs();
	};
	OutlineView.prototype.getParagraphY = function (paragraph) {
		if (this.outlineShape) {
			return this.outlineShape.transformText.TransformPointY(0, paragraph.Y);
		}
		return null;
	}
	OutlineView.prototype.drawDecorations = function (graphics, currentSlideIndex) {
		if (!this.outlineShape || !this.isHaveParagraphs()) return;

		const rectX = 2;
		const rectW = 6;
		const backgroundRGB = this.getRGBFromHex(AscCommon.GlobalSkin.BackgroundColorThumbnails);

		const numberPenRGB = this.getRGBFromHex(AscCommon.GlobalSkin.ThumbnailsPageOutline);
		const numberWShape = rectW * (3 / 5);

		g_oTextMeasurer.SetTextPr(this.getTextPr(), null);
		g_oTextMeasurer.SetFontSlot(AscWord.fontslot_ASCII);
		const height = g_oTextMeasurer.GetHeight();

		const outlineParagraphs = this.getOutlineParagraphs();
		for (let i = 0; i < outlineParagraphs.length; i += 1) {
			const paragraph = outlineParagraphs[i];
			const info = this.outlineInfo[paragraph.Get_Id()];
			if (info) {
				const topY = this.getParagraphY(paragraph);
				if (info.titleShapeIndex !== undefined) {
					const penRGB = this.getRGBFromHex(info.titleShapeIndex === currentSlideIndex ? AscCommon.GlobalSkin.ThumbnailsPageOutlineActive : AscCommon.GlobalSkin.ThumbnailsPageOutline);
					const barShape = this.createDecorShape(backgroundRGB, penRGB, rectW, height, "roundRect", 20000);
					this.drawDecorShape(graphics, barShape, rectX, topY);
				} else if (info.contentShapeIndex !== undefined) {
					const paraY = this.getParagraphY(paragraph);
					const badgeShape = this.createDecorShape(backgroundRGB, numberPenRGB, numberWShape, height, "rect", 0, String(info.contentShapeIndex + 1));
					this.drawDecorShape(graphics, badgeShape, rectX + rectW - numberWShape, paraY);
				}
			}
		}
	};
	OutlineView.prototype.getDocContent = function () {
		return this.outlineShape && this.outlineShape.txBody.content;
	};
	OutlineView.prototype.selectionSetStart = function (e, x, y) {
		if (this.outlineShape) {
			this.outlineShape.selectionSetStart(e, x, y, 0);
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
	OutlineView.prototype.getSelectedSlide = function () {
		const docContent = this.getDocContent();
		if (!docContent) {
			return 0;
		}
		const outlineParagraphs = this.getOutlineParagraphs();
		const currentParagraph = docContent.GetCurrentParagraph();
		if(currentParagraph) {
			for (let i = currentParagraph.Index; i >= 0; i -= 1) {
				const paragraph = outlineParagraphs[i];
				const info = this.outlineInfo[paragraph.Get_Id()];
				if (info && info.titleShapeIndex !== undefined) {
					return info.titleShapeIndex;
				}
			}
		}
		return null;
	};


	window["AscCommonSlide"] = window["AscCommonSlide"] || {};
	window["AscCommonSlide"].OutlineSlide = OutlineSlide;
	window["AscCommonSlide"].OutlineView = OutlineView;
})();
