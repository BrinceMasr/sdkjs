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

(function (undefined) {

	function OutlineSlide() {
		this.title = null;
		this.content = [];
		this.slide = null;
	}

	OutlineSlide.prototype.setSlide = function (pr) {
		this.slide = pr;
	};
	OutlineSlide.prototype.setTitle = function (pr) {
		this.title = pr;
	};
	OutlineSlide.prototype.addContent = function (pr) {
		this.content.push(pr);
	};
	OutlineSlide.prototype.forEachShape = function (callback) {
		let res = null;
		if (this.title) {
			res = callback(this.title, 0);
			if (res) {
				return res;
			}
		}
		for (let i = 0; i < this.content.length; i += 1) {
			res = callback(this.content[i], i + 1);
			if (res) {
				return res;
			}
		}
		return res;
	};
	function SavedPosition(startPos, endPos) {
		this.startPos = startPos || null;
		this.endPos = endPos || null;
	}

	function UpdateNewParagraphsManager(outlineView, newParagraphs) {
		this.outlineView = outlineView;
		this.paragraphs = newParagraphs;
		this.slideNumberToOutlineSlide = {};

		this.paragraphIdToSlideIndex = {};
		this.paragraphIdToShape = {};
		this.paragraphIdToShapeIndex = {};
	}
	UpdateNewParagraphsManager.prototype.getOutlineSlide = function (num) {
		if (!this.slideNumberToOutlineSlide[num]) {
			const presentation = this.outlineView.getPresentation();
			this.slideNumberToOutlineSlide[num] = presentation.Slides[num].getOutlineSlide();
		}
		return this.slideNumberToOutlineSlide[num];
	};
	UpdateNewParagraphsManager.prototype.getSlideIndex = function (paragraph) {
		const paragraphId = paragraph.Get_Id();
		if (this.paragraphIdToSlideIndex[paragraphId] === undefined) {
			const shape = paragraph.GetParentShape();
			const slide = shape.parent;
			this.paragraphIdToSlideIndex[paragraphId] = slide.num;
		}
		return this.paragraphIdToSlideIndex[paragraphId];
	};
	UpdateNewParagraphsManager.prototype.getShapeIndex = function (paragraph) {
		const paragraphId = paragraph.Get_Id();
		if (this.paragraphIdToShapeIndex[paragraphId] === undefined) {
			const oThis = this;
			const slideIndex = this.getSlideIndex(paragraph);
			const outlineSlide = this.getOutlineSlide(slideIndex);
			const paragraphShape = this.getParagraphShape(paragraph);
			outlineSlide.forEachShape(function (shape, idx) {
				if (paragraphShape === shape) {
					oThis.paragraphIdToShapeIndex[paragraphId] = idx;
					return true;
				}
			});
		}
		return this.paragraphIdToShapeIndex[paragraphId];
	};
	UpdateNewParagraphsManager.prototype.getContentIndex = function (paragraph) {
		return paragraph.Index;
	};
	UpdateNewParagraphsManager.prototype.getParagraphShape = function (paragraph) {
		const paragraphId = paragraph.Get_Id();
		if (!this.paragraphIdToShape[paragraphId]) {
			this.paragraphIdToShape[paragraphId] = paragraph.GetParentShape();
		}
		return this.paragraphIdToShape[paragraphId];
	};
	UpdateNewParagraphsManager.prototype.compareParagraphs = function (aParagraph, bParagraph) {
		const aSlideIndex = this.getSlideIndex(aParagraph);
		const bSlideIndex = this.getSlideIndex(bParagraph);
		if (aSlideIndex !== bSlideIndex) {
			return aSlideIndex - bSlideIndex;
		}
		const aShapeIndex = this.getShapeIndex(aParagraph);
		const bShapeIndex = this.getShapeIndex(bParagraph);
		if (aShapeIndex !== bShapeIndex) {
			return aShapeIndex - bShapeIndex;
		}
		const aContentIndex = this.getContentIndex(aParagraph);
		const bContentIndex = this.getContentIndex(bParagraph);
		return aContentIndex - bContentIndex;
	};
	UpdateNewParagraphsManager.prototype.update = function () {
		const outlineView = this.outlineView;
		const outlineDocContent = this.outlineView.getDocContent();
		const outlineParagraphs = outlineDocContent.Content;
		for (let i = 0; i < this.paragraphs.length; i++) {
			const sourceParagraph = this.paragraphs[i];
			const insertIndex = this.getInsertIndex(sourceParagraph, outlineParagraphs);
			outlineView.addUpdatedParagraph(sourceParagraph, insertIndex);
		}
	};
	UpdateNewParagraphsManager.prototype.getInsertIndex = function (sourceParagraph, outlineParagraphs) {
		const outlineView = this.outlineView;
		for (let i = outlineParagraphs.length - 1; i >= 0; i -= 1) {
			const outlineParagraph = outlineParagraphs[i];
			const sourceOutlineParagraph = outlineView.outlineToSourceMap[outlineParagraph.Get_Id()];
			if (!sourceOutlineParagraph) {
				const slideInfo = outlineView.outlineInfo[outlineParagraph.Get_Id()];
				const slideIndex = slideInfo.titleShapeIndex;
				const sourceSlideIndex = this.getSlideIndex(sourceParagraph);
				if (slideIndex < sourceSlideIndex) {
					return i;
				}
			} else if (this.compareParagraphs(sourceOutlineParagraph, sourceParagraph) < 0) {
				return i + 1;
			}
		}
		return 0;


			//todo binary search
		// if (outlineParagraphs.length === 0) {
		// 	return 0;
		// }
		// const outlineView = this.outlineView;
		// let leftPointer = 0;
		// let rightPointer = outlineParagraphs.length - 1;
		// while (leftPointer !== rightPointer) {
		// 	const centerPointer = leftPointer + Math.floor((rightPointer - leftPointer) / 2);
		// 	const outlineParagraph = outlineParagraphs[centerPointer];
		// 	const sourceOutlineParagraph = outlineView.outlineToSourceMap[outlineParagraph.Get_Id()];
		// 	const compareResult = this.compareParagraphs(sourceOutlineParagraph, sourceParagraph);
		// 	if (compareResult < 0) {
		// 		leftPointer = centerPointer;
		// 	} else {
		// 		rightPointer = centerPointer;
		// 	}
		// }
		//
		// const centerOutlineParagraph = outlineParagraphs[leftPointer];
		// const sourceOutlineParagraph = outlineView.outlineToSourceMap[centerOutlineParagraph.Get_Id()];
		// const compareResult = this.compareParagraphs(sourceOutlineParagraph, sourceParagraph);
		// if (compareResult < 0) {
		// 	return leftPointer + 1
		// }
		// return leftPointer;
	};

	function OutlineView() {
		this.outlineShape = null;
		this.reset();
	}

	OutlineView.prototype.reset = function () {
		this.outlineToSourceMap = {};
		this.sourceToOutlineMap = {};
		this.outlineInfo = {};
		this.resetMapToCheckParagraphs();
		this.resetPosition();
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
		this.createOutlineShape(outlineSlides, width, height);
	};
	OutlineView.prototype.addOutlineParagraph = function (sourceParagraph, outlineParagraph, pr) {
		const outlineId = outlineParagraph.Get_Id();
		this.outlineToSourceMap[outlineId] = sourceParagraph;
		if (sourceParagraph) {
			this.sourceToOutlineMap[sourceParagraph.Get_Id()] = outlineParagraph;
		}
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
	OutlineView.prototype.removeOutlineParagraph = function (outlineParagraph) {
		const outlineId = outlineParagraph.Get_Id();
		const sourceParagraph = this.outlineToSourceMap[outlineId];
		delete this.outlineToSourceMap[outlineId];
		delete this.outlineInfo[outlineId];
		if (sourceParagraph) {
			delete this.sourceToOutlineMap[sourceParagraph.Get_Id()];
		}
	};
	OutlineView.prototype.addCopyParagraph = function (paragraph, pos, isFirstParagraph, pr) {
		AscFormat.ExecuteNoHistory(function () {
			const outlineContent = this.getDocContent();
			const copyParagraph = this.getCopyParagraph(outlineContent, paragraph, !!pr && pr.titleShapeIndex !== undefined);
			outlineContent.AddToContent(pos, copyParagraph);
			this.addOutlineParagraph(paragraph, copyParagraph, isFirstParagraph ? pr : null);
		}, this, []);
	}
	OutlineView.prototype.removeParagraph = function (outlineParagraph, pos) {
		AscFormat.ExecuteNoHistory(function () {
			const outlineContent = this.getDocContent();
			outlineContent.Remove_FromContent(pos, 1);
			this.removeOutlineParagraph(outlineParagraph);
		}, this, []);
	}
	OutlineView.prototype.addContentToOutlineShape = function (outlineShape, slideShape, pr) {
		if (!slideShape) {
			return;
		}
		const outlineContent = outlineShape.txBody.content;
		const slideContent = slideShape.txBody.content;
		const paragraphs = slideContent.Content;
		for (let i = 0; i < paragraphs.length; i += 1) {
			const paragraph = paragraphs[i];
			this.addCopyParagraph(paragraph, outlineContent.Content.length, i === 0, pr);
		}
	};
	OutlineView.prototype.addMockTitleToOutlineShape = function (outlineShape, pr) {
		AscFormat.ExecuteNoHistory(function () {
			const outlineContent = outlineShape.txBody.content;
			const paragraph = new AscWord.Paragraph(outlineContent, true);
			this.applyParagraphProps(paragraph, null, true);
			outlineContent.AddToContent(outlineContent.Content.length, paragraph);
			this.addOutlineParagraph(null, paragraph, pr);
		}, this, []);

	};
	OutlineView.prototype.fillOutlineShape = function (outlineShape, outlineSlides, width) {
		for (let i = 0; i < outlineSlides.length; i += 1) {
			const slide = outlineSlides[i];

			const titlePr = {titleShapeIndex: i};
			if (slide.title !== null) {
				this.addContentToOutlineShape(outlineShape, slide.title, titlePr);
			} else {
				this.addMockTitleToOutlineShape(outlineShape, titlePr);
			}

			let shapeCount = 0;
			for (let j = 0; j < slide.content.length; j += 1) {
				const contentShape = slide.content[j];
				if (!contentShape.txBody.content.IsEmpty()) {
					this.addContentToOutlineShape(outlineShape, slide.content[j], slide.content.length > 1 ? {contentShapeIndex: shapeCount} : null);
					shapeCount += 1;
				}
			}
		}
	};
	OutlineView.prototype.getTextPr = function (isTitle, isMathRun) {
		const textPr = new CTextPr();
		textPr.SetFontSize(10);
		if (isTitle) {
			textPr.SetBold(true);
		}
		if (isMathRun) {
			textPr.SetFontFamily("Cambria Math");
			textPr.RFonts.SetAll("Cambria Math");
		} else {
			textPr.SetFontFamily("Arial");
			textPr.RFonts.SetAll("Arial");
		}
		return textPr;
	};
	OutlineView.prototype.getParaPr = function (compiledParaPr, isTitle) {
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
		const compiledPr = slideParagraph ? slideParagraph.getCompiledPr() : {ParaPr: g_oDocumentDefaultParaPr};
		const copyParaPr = this.getParaPr(compiledPr, isTitle);
		const oThis = this;
		outlineParagraph.SetPr(copyParaPr);
		outlineParagraph.CheckRunContent(function (run) {
			const textPr = oThis.getTextPr(isTitle, run.IsMathRun());
			run.SetPr(textPr);
		});
		const parTextPr = this.getTextPr();
		outlineParagraph.TextPr.Value = parTextPr;
		outlineParagraph.TextPr.CalcValue = parTextPr;
	};
	OutlineView.prototype.createOutlineShape = function (outlineSlides, width) {
		this.reset();
		return AscFormat.ExecuteNoHistory(function () {
			const outlineShape = new AscFormat.CShape();
			this.setOutlineShape(outlineShape);
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
			const copyParagraph = paragraph.Copy2(parent, null, null);
			this.applyParagraphProps(copyParagraph, paragraph, isTitle);
			return copyParagraph;
		}, this, []);
	};
	OutlineView.prototype.draw = function (graphics) {
		if (this.outlineShape) {
			this.update();
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
				var textPr = this.getTextPr();
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
		return this.getTransformY(paragraph.Y);
	};
	OutlineView.prototype.getTransformY = function (y) {
		if (this.outlineShape) {
			return this.outlineShape.transformText.TransformPointY(0, y);
		}
		return null;
	};
	OutlineView.prototype.getTransformX = function (x) {
		if (this.outlineShape) {
			return this.outlineShape.transformText.TransformPointX(x, 0);
		}
		return null;
	};
	OutlineView.prototype.getInvertTransformY = function (y) {
		if (this.outlineShape) {
			return this.outlineShape.invertTransformText.TransformPointY(0, y);
		}
		return null;
	};
	OutlineView.prototype.getInvertTransformX = function (x) {
		if (this.outlineShape) {
			return this.outlineShape.invertTransformText.TransformPointX(x, 0);
		}
		return null;
	};
	OutlineView.prototype.drawDecorations = function (graphics, currentSlideIndex, focusSlideIndex) {
		if (!this.outlineShape || !this.isHaveParagraphs()) return;

		const rectX = 2;
		const rectW = 6;
		const backgroundRGB = this.getRGBFromHex(AscCommon.GlobalSkin.BackgroundColorThumbnails);

		const normalSlideRGB = this.getRGBFromHex(AscCommon.GlobalSkin.ThumbnailsPageOutline);
		const activeSlideRGB = this.getRGBFromHex(AscCommon.GlobalSkin.ThumbnailsPageOutlineActive);
		const hoverSlideRGB = this.getRGBFromHex(AscCommon.GlobalSkin.ThumbnailsPageOutlineHover);
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
					let penRGB;
					if (info.titleShapeIndex === currentSlideIndex) {
						penRGB = activeSlideRGB;
					} else if (info.titleShapeIndex === focusSlideIndex) {
						penRGB = hoverSlideRGB;
					} else {
						penRGB = normalSlideRGB;
					}
					const barShape = this.createDecorShape(backgroundRGB, penRGB, rectW, height, "roundRect", 20000);
					this.drawDecorShape(graphics, barShape, rectX, topY);
				} else if (info.contentShapeIndex !== undefined) {
					const badgeShape = this.createDecorShape(backgroundRGB, normalSlideRGB, numberWShape, height, "rect", 0, String(info.contentShapeIndex + 1));
					this.drawDecorShape(graphics, badgeShape, rectX + rectW - numberWShape, topY);
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
		if (currentParagraph) {
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
	OutlineView.prototype.getTransformText = function () {
		if (this.outlineShape) {
			return this.outlineShape.transformText;
		}
		return new AscCommon.CMatrix();
	};
	OutlineView.prototype.getOutlineHeight = function () {
		if (this.outlineShape) {
			return this.outlineShape.transformText.TransformPointY(0, this.outlineShape.contentHeight);
		}
		return 0;
	}
	OutlineView.prototype.rebuildPos = function (pos, sourceContent, sourceParagraph) {
		const newPos = pos.slice();
		newPos[0] = {Class: sourceContent, Position: sourceParagraph.GetIndex()};
		newPos[1] = {Class: sourceParagraph, Position: pos[1].Position};
		return newPos;
	}
	OutlineView.prototype.getSelectionUseContentsInfo = function (startPos, endPos) {
		const contents = [];
		let direction = AscWord.Direction.FORWARD;
		let startPosIndex = startPos[0].Position;
		let endPosIndex = endPos[0].Position;
		if (startPosIndex > endPosIndex) {
			const temp = startPosIndex;
			startPosIndex = endPosIndex;
			endPosIndex = temp;
			direction = AscWord.Direction.BACKWARD;
		}
		const content = this.getDocContent();
		for (let i = startPosIndex; i <= endPosIndex; i += 1) {
			const contentInfo = contents[contents.length - 1];
			const outlineParagraph = content.Content[i];
			const sourceParagraph = this.outlineToSourceMap[outlineParagraph.Get_Id()];
			if (sourceParagraph) {
				const sourceContent = sourceParagraph.GetParent();
				if (contentInfo && contentInfo.content === sourceContent) {
					if (direction === AscWord.Direction.FORWARD) {
						contentInfo.endPos = endPos;
						contentInfo.endParagraph = sourceParagraph;
					} else {
						contentInfo.startPos = startPos;
						contentInfo.startParagraph = sourceParagraph;
					}
				} else {
					contents.push({
						startPos: startPos,
						startParagraph: sourceParagraph,
						endPos: endPos,
						endParagraph: sourceParagraph,
						content: sourceContent,
						direction: direction
					});
				}
			}
		}
		return contents;
	}
	OutlineView.prototype.getContentPos = function (callback) {
		let res;
		const content = this.getDocContent();
		if (content.IsSelectionUse()) {
			const startPos = content.GetContentPosition(true, true);
			const endPos = content.GetContentPosition(true, false);
			res = callback(true, startPos, endPos);
		} else {
			const contentPos = content.GetContentPosition(false, false);
			res = callback(false, contentPos);
		}
		return res;
	};
	OutlineView.prototype.forEachSelectedContent = function (callback) {
		this.update();
		const content = this.getDocContent();
		if (!content) {
			return;
		}

		if (content.IsSelectionUse()) {
			const startPos = content.GetContentPosition(true, true);
			const endPos = content.GetContentPosition(true, false);
			const contents = this.getSelectionUseContentsInfo(startPos, endPos);
			if (contents.length === 1) {
				const contentInfo = contents[0];
				const content = contentInfo.content;
				const startPos = this.rebuildPos(contentInfo.startPos, content, contentInfo.startParagraph);
				const endPos = this.rebuildPos(contentInfo.endPos, content, contentInfo.endParagraph);
				content.SetContentSelection(startPos, endPos, 0, 0, 0, 0);
				const res = callback(content, 0, 1);
				content.RemoveSelection();
				return res;
			} else if (contents.length > 1) {
				const startContentInfo = contents[0];
				const startContent = startContentInfo.content;
				startContent.MoveCursorToEndPos(false, true);
				const targetStartPos = startContent.GetContentPosition(true, true);
				if (startContentInfo.direction === AscWord.Direction.FORWARD) {
					const startPos = this.rebuildPos(startContentInfo.startPos, startContent, startContentInfo.startParagraph);
					startContent.SetContentSelection(startPos, targetStartPos, 0, 0, 0);
				} else {
					const startPos = this.rebuildPos(startContentInfo.endPos, startContent, startContentInfo.endParagraph);
					startContent.SetContentSelection(targetStartPos, startPos, 0, 0, 0);
				}
				const startRes = callback(startContent, 0, contents.length);
				startContent.RemoveSelection();
				if (startRes) {
					return true;
				}
				for (let i = 1; i < contents.length - 1; i += 1) {
					const contentInfo = contents[i];
					const content = contentInfo.content;
					content.SelectAll(contentInfo.direction);
					const res = callback(content, i, contents.length);
					content.RemoveSelection();
					if (res) {
						return true;
					}
				}
				const endContentInfo = contents[contents.length - 1];
				const endContent = endContentInfo.content;
				endContent.MoveCursorToStartPos();
				const endTargetPos = endContent.GetContentPosition(true, true);
				if (endContentInfo.direction === AscWord.Direction.FORWARD) {
					const endPos = this.rebuildPos(endContentInfo.endPos, endContent, endContentInfo.endParagraph);
					endContent.SetContentSelection(endTargetPos, endPos, 0, 0, 0);
				} else {
					const endPos = this.rebuildPos(endContentInfo.startPos, endContent, endContentInfo.startParagraph);
					endContent.SetContentSelection(endPos, endTargetPos, 0, 0, 0);
				}
				const endRes = callback(endContent, contents.length - 1, contents.length);
				endContent.RemoveSelection();
				if (endRes) {
					return true;
				}
			}
		} else {
			const contentPos = content.GetContentPosition(false, false);
			const paragraph = content.Content[content.CurPos.ContentPos];
			const sourceParagraph = this.outlineToSourceMap[paragraph.Get_Id()];
			if (sourceParagraph) {
				const sourceContent = sourceParagraph.GetParent();
				const startPos = this.rebuildPos(contentPos, sourceContent, sourceParagraph);
				sourceContent.SetContentPosition(startPos, 0, 0);
				const res = callback(sourceContent, 0, false);
				sourceContent.RemoveSelection();
				if (res) {
					return true;
				}
			} else {

			}
		}
		return false;
	}
	OutlineView.prototype.getSelectedParagraphs = function () {
		const docContent = this.getDocContent();
		const selectedParagraphs = [];
		if (docContent) {
			docContent.GetCurrentParagraph(false, selectedParagraphs);
		}
		return selectedParagraphs;
	};
	OutlineView.prototype.checkRemoveEdgeContents = function (firstContent, lastContent) {
		if (firstContent && lastContent) {
			const firstShape = firstContent.Is_DrawingShape(true);
			const lastShape = lastContent.Is_DrawingShape(true);
			const firstSlide = firstShape.parent;
			const lastSlide = lastShape.parent;

			const firstContentParagraph = firstContent.Content[firstContent.Content.length - 1];
			firstContentParagraph.Concat(lastContent.Content[0]);
			lastContent.Remove_FromContent(0, 1);
			let isSaveLastShape = true;
			if (firstShape.isOutlineTitlePlaceholder()) {
				if (lastContent.Content.length === 1) {
					isSaveLastShape = false;
				}
			} else {
				for (let i = 0; i < lastContent.Content.length; i += 1) {
					firstContent.Add_ToContent(firstContent.Content.length, lastContent.Content[i].Copy(firstContent));
				}
				isSaveLastShape = false;
			}


			if (firstSlide !== lastSlide) {
				const outlineSlide = lastSlide.getOutlineSlide();
				let isLastShapeChecked = false;
				outlineSlide.forEachShape(function (shape) {
					if (isLastShapeChecked) {
						const copyShape = shape.copy();
						copyShape.setBDeleted(false);
						copyShape.setParent(firstSlide);
						copyShape.addToDrawingObjects(firstSlide.cSld.spTree.length);
						shape.deleteDrawingBase();
					} else {
						isLastShapeChecked = shape === lastShape;
					}
				});
			}
			if (!isSaveLastShape) {
				lastShape.deleteDrawingBase();
			}
		}
	};
	OutlineView.prototype.remove = function (bOnAddText) {
		const checkSlidesForRemove = {};
		let firstContent = null;
		let lastContent = null;
		const oThis = this;
		this.forEachSelectedContent(function (content, idx, count) {
			if (content.IsSelectedAll()) {
				const shape = content.Is_DrawingShape(true);
				const slide = shape.parent;
				const slideNum = slide.num;
				checkSlidesForRemove[slideNum] = slide;
				shape.deleteDrawingBase();
				content.RemoveSelection();
			} else {
				content.Remove(1, true, false, bOnAddText);
				if (idx === 0) {
					firstContent = content;
				} else if (idx === count - 1) {
					lastContent = content;
				}
			}
			// oThis.savePositionAfterEdit(content, idx, count);
		});
		this.getDocContent().RemoveSelection()
		this.checkRemoveEdgeContents(firstContent, lastContent);
		const slideNumbers = Object.keys(checkSlidesForRemove).map(function (num) {
			return parseInt(num, 10);
		}).sort(function (a, b) {
			return b - a;
		});
		const presentation = this.getPresentation();
		for (let i = 0; i < slideNumbers.length; i++) {
			const slideNum = slideNumbers[i];
			const slide = checkSlidesForRemove[slideNum];
			if (!slide.isHaveOutlineShapes()) {
				presentation.removeSlideByObject(slide, true, slideNum);
			}
		}
		const minSlideIndex = slideNumbers[slideNumbers.length - 1] || 0;
		presentation.DrawingDocument.m_oWordControl.GoToPage(Math.min(presentation.GetSlidesCount() - 1, minSlideIndex), undefined, undefined, true);
	};
	OutlineView.prototype.savePositionAfterEdit = function (content, idx, count) {
		if (count === 1) {
			if (content.IsSelectionUse()) {
				const startPos = content.GetContentPosition(true, true);
				const endPos = content.GetContentPosition(true, false);
				this.setSavedPosition(new SavedPosition(startPos, endPos));
			} else {
				this.setSavedPosition(new SavedPosition(content.GetContentPosition(false, false)));
			}
		} else {
			if (idx === 0) {
					const savedPosition = this.getSavedPosition();
					if (content.IsSelectionUse()) {
						if (content.Selection.StartPos > content.Selection.EndPos) {
							savedPosition.startPos = content.GetContentPosition(true, false);
						} else {
							savedPosition.startPos = content.GetContentPosition(true, true);
						}
					} else if (!savedPosition.startPos) {
						savedPosition.startPos = content.GetContentPosition(false, false);
					}

			} else if (idx === count - 1) {
					const savedPosition = this.getSavedPosition();
					if (content.IsSelectionUse()) {
						if (content.Selection.StartPos > content.Selection.EndPos) {
							savedPosition.endPos = content.GetContentPosition(true, true);
						} else {
							savedPosition.endPos = content.GetContentPosition(true, false);
						}
					} else if (!savedPosition.startPos) {
						savedPosition.startPos = content.GetContentPosition(false, false);
					}

			}
		}
	};
	OutlineView.prototype.rebuildSavedPositionPos = function (pos) {
		const docContent = this.getDocContent();
		if (docContent) {
			const newPos = pos.slice();
			const sourceParagraph = newPos[1].Class;
			const outlineParagraph = this.sourceToOutlineMap[sourceParagraph.GetId()];
			if (outlineParagraph) {
				newPos[0] = {Class: docContent, Position: outlineParagraph.Index};
				newPos[1] = {Class: outlineParagraph, Position: pos[1].Position};
				return newPos;
			}
		}
		return null;
	};
	OutlineView.prototype.applySavedPositionToOutline = function () {
		const docContent = this.getDocContent();
		if (docContent) {
			const savedPostion = this.getSavedPosition();
			if (savedPostion.startPos) {
				if (savedPostion.endPos) {
					const startPos = this.rebuildSavedPositionPos(savedPostion.startPos);
					const endPos = this.rebuildSavedPositionPos(savedPostion.endPos);
					docContent.SetContentSelection(startPos, endPos, 0, 0, 0);
				} else {
					docContent.SetContentPosition(this.rebuildSavedPositionPos(savedPostion.startPos), 0, 0);
				}
				this.updateSelectionState();
			}
		}
	}
	OutlineView.prototype.setSavedPosition = function (position) {
		this.savedPosition = position;
	};
	OutlineView.prototype.getSavedPosition = function () {
		return this.savedPosition;
	};
	OutlineView.prototype.resetPosition = function () {
		this.setSavedPosition(new SavedPosition());
	};
	OutlineView.prototype.resetMapToCheckParagraphs = function () {
		this.mapToCheckParagraphs = {};
	};
	OutlineView.prototype.paragraphAdd = function (paraItem) {
		const oThis = this;
		if (paraItem.Type === para_TextPr) {
			this.forEachSelectedContent(function (content, idx, contentCount) {
				content.AddToParagraph(paraItem);
				oThis.savePositionAfterEdit(content, idx, contentCount);
			});
		} else {
			this.forEachSelectedContent(function (content, idx, contentCount) {
				if (contentCount > 1) {

				} else {
					content.AddToParagraph(paraItem);
				}
				oThis.savePositionAfterEdit(content, idx, contentCount);
			});
		}
	}

	OutlineView.prototype.getParagraphParaPr = function () {
		let paraPr;
		this.forEachSelectedContent(function (content) {
			if (content) {
				const contentParaPr = content.GetCalculatedParaPr();
				if (paraPr) {
					paraPr.Compare(contentParaPr);
				} else {
					paraPr = contentParaPr;
				}
			}
		});
		return paraPr;
	};

	OutlineView.prototype.getParagraphTextPr = function () {
		let textPr;
		this.forEachSelectedContent(function (content) {
			if (content) {
				const contentTextPr = content.GetCalculatedTextPr();
				if (textPr) {
					textPr.Compare(contentTextPr);
				} else {
					textPr = contentTextPr;
				}
			}
		});
		return textPr;
	};
	OutlineView.prototype.unlinkSourceParagraph = function (sourceParagraph) {
		const outlineParagraph = this.sourceToOutlineMap[sourceParagraph.Get_Id()];
		const index = outlineParagraph.Index;
		this.removeParagraph(outlineParagraph, index);
	};
	OutlineView.prototype.getPropertiesFromSourceShape = function (sourceParagraph) {
		let pr = null;
		if (sourceParagraph.Index === 0) {
			const parentShape = sourceParagraph.GetParentShape();
			pr = {};
			if (parentShape.getPlaceholderType() === AscFormat.phType_ctrTitle || parentShape.getPlaceholderType() === AscFormat.phType_title) {
				pr.titleShapeIndex = parentShape.parent.num;
			} else {
				pr.contentShapeIndex = this.getShapeContentIndex(parentShape.parent, parentShape);
			}
		}
		return pr;
	};
	OutlineView.prototype.getShapeContentIndex = function (slide, shape) {
		let countOfLowerShapes = 0;
		const mainIndex = parseInt(shape.getPlaceholderIndex(), 10) || 0;
		for (let i = 0; i < slide.cSld.spTree.length; i += 1) {
			const shape = slide.cSld.spTree[i];
			if (shape.isOutlineContentPlaceholder()) {
				const shapeIndex = parseInt(shape.getPlaceholderIndex(), 10) || 0;
				if (mainIndex > shapeIndex) {
					countOfLowerShapes += 1;
				}
			}
		}
		return countOfLowerShapes;
	}
	OutlineView.prototype.updateFromSourceParagraph = function (sourceParagraph) {
		const outlineParagraph = this.sourceToOutlineMap[sourceParagraph.Get_Id()];
		const index = outlineParagraph.Index;
		this.removeParagraph(outlineParagraph, index);
		this.addUpdatedParagraph(sourceParagraph, index);
	};
	OutlineView.prototype.addUpdatedParagraph = function (sourceParagraph, index) {
		const pr = this.getPropertiesFromSourceShape(sourceParagraph);
		this.addCopyParagraph(sourceParagraph, index, sourceParagraph.Index === 0, pr);
	};
	OutlineView.prototype.update = function () {
		const existingParagraphs = [];
		const newParagraphs = [];
		for (let id in this.mapToCheckParagraphs) {
			const paragraph = this.mapToCheckParagraphs[id];
			const shape = paragraph.GetParentShape();
			if (shape && shape.isOutlinePlaceholder()) {
				if (shape.parent && shape.parent.getObjectType() === AscDFH.historyitem_type_Slide) {
					if (this.sourceToOutlineMap[paragraph.Get_Id()]) {
						existingParagraphs.push(paragraph);
					} else if (paragraph.IsUseInDocument()) {
						newParagraphs.push(paragraph);
					}
				}
			}
		}


		const isNeedRecalculate = !!existingParagraphs.length || !!newParagraphs.length;
		if (isNeedRecalculate) {
			this.updateExistingParagraphs(existingParagraphs);
			this.updateNewParagraphs(newParagraphs);
			this.outlineShape && this.outlineShape.recalculateContent();
			this.applySavedPositionToOutline();
			this.resetMapToCheckParagraphs();
			this.resetPosition();
		}
	};
	OutlineView.prototype.updateExistingParagraphs = function (existingParagraphs) {
		const oThis = this;
		existingParagraphs.sort(function (aParagraph, bParagraph) {
			const aOutlineParagraph = oThis.sourceToOutlineMap[aParagraph.Get_Id()];
			const bOutlineParagraph = oThis.sourceToOutlineMap[bParagraph.Get_Id()];
			return bOutlineParagraph.Index - aOutlineParagraph.Index;
		});
		for (let i = 0; i < existingParagraphs.length; i += 1) {
			const paragraph = existingParagraphs[i];
			if (paragraph.IsUseInDocument()) {
				this.updateFromSourceParagraph(paragraph);
			} else {
				this.unlinkSourceParagraph(paragraph);
			}
		}
	}

	OutlineView.prototype.updateNewParagraphs = function (newParagraphs) {
		const updateManager = new UpdateNewParagraphsManager(this, newParagraphs);
		updateManager.update();
	}
	OutlineView.prototype.checkSourceParagraph = function (paragraph) {
		if (AscCommon.History.IsOn()) {
			this.mapToCheckParagraphs[paragraph.GetId()] = paragraph;
		}
	};
	OutlineView.prototype.checkSourceSlide = function (slide) {
		for (let i = 0; i < slide.cSld.spTree.length; i += 1) {
			const shape = slide.cSld.spTree[i];
			this.checkSourceShape(shape);
		}
	};
	OutlineView.prototype.checkSourceShape = function (shape) {
		if (shape.isOutlinePlaceholder()) {
			const content = shape.getDocContent();
			for (let i = 0; i < content.Content.length; i += 1) {
				const paragraph = content.Content[i];
				this.checkSourceParagraph(paragraph);
			}
		}
	};
	OutlineView.prototype.getApi = function () {
		return Asc.editor;
	};
	OutlineView.prototype.selectAll = function () {
		const content = this.getDocContent();
		if (content) {
			content.SelectAll();
			const drawingDocument = this.outlineShape.getDrawingDocument();
			this.outlineShape.updateSelectionState(drawingDocument);
		}
	};
	OutlineView.prototype.executeShortcut = function (shortcutAction) {
		let res = {keyResult: keydownresult_PreventAll};
		switch (shortcutAction) {
			case Asc.c_oAscPresentationShortcutType.EditSelectAll: {
				this.selectAll();
				break;
			}
			default: {
				res = null;
			}
		}
		return res;
	}
	OutlineView.prototype.updateInterfaceState = function () {

	};
	OutlineView.prototype.updateSelectionState = function () {
		if (this.outlineShape) {
			const drawingDocument = this.outlineShape.getDrawingDocument();
			this.outlineShape.updateSelectionState(drawingDocument);
		}
	};
	OutlineView.prototype.moveCursorToStartPos = function (AddToSelect) {
		const content = this.getDocContent();
		content.MoveCursorToStartPos(AddToSelect);
		this.updateSelectionState();
		this.updateInterfaceState();
	};

	OutlineView.prototype.moveCursorToEndPos = function (AddToSelect) {
		const content = this.getDocContent();
		content.MoveCursorToEndPos(AddToSelect);
		this.updateSelectionState();
		this.updateInterfaceState();
	};

	OutlineView.prototype.moveCursorLeft = function (AddToSelect, Word) {
		const content = this.getDocContent();

		if (this.isRtl())
			content.MoveCursorRight(AddToSelect, Word);
		else
			content.MoveCursorLeft(AddToSelect, Word);
		this.updateSelectionState();
		this.updateInterfaceState();
	};

	OutlineView.prototype.moveCursorRight = function (AddToSelect, Word) {
		const content = this.getDocContent();
		if (this.isRtl())
			content.MoveCursorLeft(AddToSelect, Word);
		else
			content.MoveCursorRight(AddToSelect, Word);
		this.updateSelectionState();
		this.updateInterfaceState();
	};

	OutlineView.prototype.moveCursorUp = function (AddToSelect) {
		const content = this.getDocContent();
		content.MoveCursorUp(AddToSelect);
		this.updateSelectionState();
		this.updateInterfaceState();
	};

	OutlineView.prototype.moveCursorDown = function (AddToSelect) {
		const content = this.getDocContent();
		content.MoveCursorDown(AddToSelect);
		this.updateSelectionState();
		this.updateInterfaceState();
	};

	OutlineView.prototype.moveCursorToEndOfLine = function (AddToSelect) {
		const content = this.getDocContent();
		content.MoveCursorToEndOfLine(AddToSelect);
		this.updateSelectionState();
		this.updateInterfaceState();
	};

	OutlineView.prototype.moveCursorToStartOfLine = function (AddToSelect) {
		const content = this.getDocContent();
		content.MoveCursorToStartOfLine(AddToSelect);
		this.updateSelectionState();
		this.updateInterfaceState();
	};
	OutlineView.prototype.isRtl = function () {
		let oContent = this.getDocContent();
		if (oContent) {
			let curPara = oContent.GetCurrentParagraph();
			return !!(curPara && curPara.isRtlDirection());
		}
		return false;
	}
	OutlineView.prototype.onKeyDown = function (e) {
		const api = this.getApi();
		const shortcutAction = api.getShortcut(e);
		const shortcutRes = this.executeShortcut(shortcutAction);
		if (shortcutRes) {
			return shortcutRes.keyResult;
		}
		return this.executeHotKey(e);
	};
	OutlineView.prototype.executeHotKey = function (e) {
		const isMacOs = AscCommon.AscBrowser.isMacOs;
		const presentation = this.getPresentation();
		switch (e.KeyCode) {
			case Asc.c_oAscKeyCodes.ArrowLeft: {
				if (isMacOs && e.CtrlKey) {
					this.moveCursorToStartOfLine(e.ShiftKey);
				} else {
					const bIsWord = isMacOs ? e.AltKey : e.CtrlKey;
					this.moveCursorLeft(e.ShiftKey, bIsWord);
				}
				break;
			}
			case Asc.c_oAscKeyCodes.ArrowRight: {
				if (isMacOs && e.CtrlKey) {
					this.moveCursorToEndOfLine(e.ShiftKey);
				} else {
					const bIsWord = isMacOs ? e.AltKey : e.CtrlKey;
					this.moveCursorRight(e.ShiftKey, bIsWord);
				}
				break;
			}
			case Asc.c_oAscKeyCodes.ArrowUp: {
				this.moveCursorUp(e.ShiftKey);
				break;
			}
			case Asc.c_oAscKeyCodes.ArrowDown: {
				this.moveCursorDown(e.ShiftKey);
				break;
			}
			case Asc.c_oAscKeyCodes.Home: {
				if (e.CtrlKey) {
					this.moveCursorToStartPos(e.ShiftKey);
				} else {
					this.moveCursorToStartOfLine(e.ShiftKey);
				}
				break;
			}
			case Asc.c_oAscKeyCodes.End: {
				if (e.CtrlKey) {
					this.moveCursorToEndPos(e.ShiftKey);
				} else {
					this.moveCursorToEndOfLine(e.ShiftKey);
				}
				break;
			}
			case Asc.c_oAscKeyCodes.Backspace: {
				presentation.StartAction();
				this.remove();
				presentation.FinalizeAction();
				this.updateInterfaceState()
				this.updateSelectionState();
				break;
			}
		}
	};
	OutlineView.prototype.getDrawingDocument = function () {
		const presentation = this.getPresentation();
		return presentation.GetDrawingDocument();
	};
	OutlineView.prototype.isFocusOnOutline = function () {
		const presentation = this.getPresentation();
		return presentation.IsFocusOnOutline();
	};
	OutlineView.prototype.getSelectedSlidesRange = function () {
		const selectedArray = this.getSelectedSlideArray();
		if (selectedArray.length) {
			return {Min: selectedArray[0], Max: selectedArray[selectedArray.length - 1]};
		}
		return {Min: -1, Max: -1};
	};
	OutlineView.prototype.getSlideIndex = function (outlineParagraph) {
		const outlineId = outlineParagraph.Get_Id();
		const sourceParagraph = this.outlineToSourceMap[outlineId];
		if (sourceParagraph) {
			const parentShape = sourceParagraph.GetParentShape();
			if (parentShape && parentShape.parent) {
				return parentShape.parent.num;
			}
		} else if (this.outlineInfo[outlineId] && this.outlineInfo[outlineId].titleShapeIndex !== undefined) {
			return this.outlineInfo[outlineId].titleShapeIndex;
		}
		return -1;
	};
	OutlineView.prototype.getSelectedSlideArrayFromSelection = function (startPos, endPos) {
		const startParagraph = startPos[1].Class;
		const endParagraph = endPos[1].Class;
		let startSlideIndex = this.getSlideIndex(startParagraph);
		let endSlideIndex = this.getSlideIndex(endParagraph);
		if (startSlideIndex !== -1 && endSlideIndex !== -1) {
			if (startSlideIndex > endSlideIndex) {
				const temp = startSlideIndex;
				startSlideIndex = endSlideIndex;
				endSlideIndex = temp;
			}
			const selectedSlides = [];
			for (let i = startSlideIndex; i <= endSlideIndex; i += 1) {
				selectedSlides.push(i);
			}
			return selectedSlides;
		} else if (startSlideIndex !== -1) {
			return [startSlideIndex];
		} else if (endSlideIndex !== -1) {
			return [endSlideIndex];
		}
		return [];
	};
	OutlineView.prototype.getSelectedSlideArrayFromCurPos = function (curPos) {
		const startParagraph = curPos[1].Class;
		const startSlideIndex = this.getSlideIndex(startParagraph);
		if (startSlideIndex !== -1) {
			return  [startSlideIndex];
		}
		return [];
	};

	OutlineView.prototype.getSelectedSlideArray = function () {
		if (this.isFocusOnOutline()) {
			const oThis = this;
			return this.getContentPos(function (isSelectionUse, startPos, endPos) {
				if (isSelectionUse) {
					return oThis.getSelectedSlideArrayFromSelection(startPos, endPos);
				} else {
					return oThis.getSelectedSlideArrayFromCurPos(startPos);
				}
			});
		} else {
			const drawingDocument = this.getDrawingDocument();
			if (drawingDocument.SlideCurrent >= 0) {
				return [drawingDocument.SlideCurrent];
			}
		}
		return [];
	};
	OutlineView.prototype.isSelectedPage = function (pageNum) {
		const selectedPages = this.getSelectedSlideArray();
		return selectedPages.indexOf(pageNum) !== -1;
	};
	OutlineView.prototype.getTitleSlideParagraph = function (pageNum) {
		const docContent = this.getDocContent();
		if (docContent) {
			for (let i = 0; i < docContent.Content.length; i += 1) {
				const paragraph = docContent.Content[i];
				const info = this.outlineInfo[paragraph.Get_Id()];
				if (info && info.titleShapeIndex === pageNum) {
					return paragraph;
				}
			}
		}
	};
	OutlineView.prototype.selectPage = function (pageNum) {
		const docContent = this.getDocContent();
		if (docContent) {
			const paragraph = this.getTitleSlideParagraph(pageNum);
			if (paragraph) {
				paragraph.MoveCursorToStartPos();
				paragraph.SetThisElementCurrent();
			}

			this.updateSelectionState();
		}
	};

	OutlineView.prototype.getNearestPage = function (x, y) {
		const docContent = this.getDocContent();
		if (docContent) {
			const tx = this.getInvertTransformX(x);
			const ty = this.getInvertTransformY(y);
			const nearestPos = docContent.Get_NearestPos(0, tx, ty);
			return this.getSlideIndex(nearestPos.Paragraph);
		}
		return -1;
	};
	OutlineView.prototype.hitInTextRect = function (x, y) {
		if (this.outlineShape) {
			return AscFormat.HitToRect(x, y, this.outlineShape.invertTransformText, 0, 0, this.outlineShape.contentWidth, 20000);
		}
		return false;
	};



	window["AscCommonSlide"] = window["AscCommonSlide"] || {};
	window["AscCommonSlide"].OutlineSlide = OutlineSlide;
	window["AscCommonSlide"].OutlineView = OutlineView;
})();
