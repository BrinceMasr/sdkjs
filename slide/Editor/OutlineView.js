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
			const shapeIndex = this.getShapeIndex(sourceParagraph);
			let isReplaceMockParagraph = false;
			if (shapeIndex === 0 && sourceParagraph.Index === 0) {
				isReplaceMockParagraph = outlineView.addUpdatedParagraphWithCheckMockParagraph(sourceParagraph, insertIndex);
			}
			if (!isReplaceMockParagraph) {
				outlineView.addUpdatedParagraph(sourceParagraph, insertIndex);
			}
		}
	};
	UpdateNewParagraphsManager.prototype.getInsertIndex = function (sourceParagraph, outlineParagraphs) {
		const outlineView = this.outlineView;
		for (let i = outlineParagraphs.length - 1; i >= 0; i -= 1) {
			const outlineParagraph = outlineParagraphs[i];
			const sourceOutlineParagraph = outlineView.outlineToSourceMap[outlineParagraph.Get_Id()];
			if (!sourceOutlineParagraph) {
				const slideInfo = outlineView.outlineInfo[outlineParagraph.Get_Id()];
				const slideIndex = slideInfo.slide.num;
				const sourceSlideIndex = this.getSlideIndex(sourceParagraph);
				if (slideIndex <= sourceSlideIndex) {
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

	function ForEachSelectManager(outlineView) {
		this.outlineView = outlineView;
	}
	ForEachSelectManager.prototype.getSelectionUseContentsInfo = function () {
		const content = this.getDocContent();
		const startPos = content.GetContentPosition(true, true);
		const endPos = content.GetContentPosition(true, false);

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
		for (let i = startPosIndex; i <= endPosIndex; i += 1) {
			const contentInfo = contents[contents.length - 1];
			const outlineParagraph = content.Content[i];
			const sourceParagraph = this.getSourceParagraph(outlineParagraph);
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
			} else {
				contents.push({
					content: null,
					outlineParagraph: outlineParagraph
				});
			}
		}
		return contents;
	};
	ForEachSelectManager.prototype.getSourceParagraph = function (outlineParagraph) {
		return this.outlineView.outlineToSourceMap[outlineParagraph.Get_Id()]
	};
	ForEachSelectManager.prototype.getDocContent = function () {
		return this.outlineView.getDocContent();
	};
	ForEachSelectManager.prototype.processNullableContent = function (contentInfo, callback) {
		return callback(contentInfo);
	};
	ForEachSelectManager.prototype.processSingleSelectedContent = function (contentInfo, callback) {
		if (contentInfo.content) {
			return this.processSingleSelectedNotNullableContent(contentInfo, callback);
		}
		return this.processNullableContent(contentInfo, callback);

	};
	ForEachSelectManager.prototype.processSingleSelectedNotNullableContent = function (contentInfo, callback) {
		if (contentInfo.content !== null) {
			const content = contentInfo.content;
			const startPos = this.rebuildPos(contentInfo.startPos, content, contentInfo.startParagraph);
			const endPos = this.rebuildPos(contentInfo.endPos, content, contentInfo.endParagraph);
			content.SetContentPosition(startPos, 0, 0);
			content.SetContentSelection(startPos, endPos, 0, 0, 0, 0);
			const res = callback({content: content, index: 0, count: 1});
			content.RemoveSelection();
			return res;
		}
	};
	ForEachSelectManager.prototype.processStartContent = function (contentInfo, count, callback) {
		const content = contentInfo.content;
		if (!content) {
			return this.processNullableContent(contentInfo, callback);
		}

		content.MoveCursorToEndPos(false, true);
		const targetPos = content.GetContentPosition(true, true);
		if (contentInfo.direction === AscWord.Direction.FORWARD) {
			const startPos = this.rebuildPos(contentInfo.startPos, content, contentInfo.startParagraph);
			content.SetContentSelection(startPos, targetPos, 0, 0, 0);
		} else {
			const startPos = this.rebuildPos(contentInfo.endPos, content, contentInfo.endParagraph);
			content.SetContentSelection(targetPos, startPos, 0, 0, 0);
		}
		const res = callback({content: content, index: 0, count: count});
		content.RemoveSelection();
		return res;
	};
	ForEachSelectManager.prototype.processCenterContents = function (contents, callback) {
		for (let i = 1; i < contents.length - 1; i += 1) {
			const contentInfo = contents[i];
			const content = contentInfo.content;
			if (!content) {
				const nullableRes = this.processNullableContent(contentInfo, callback);
				if (nullableRes) {
					return nullableRes;
				}
			} else {
				content.SelectAll(contentInfo.direction);
				const res = callback({content: content, index: i, count: contents.length});
				content.RemoveSelection();
				if (res) {
					return true;
				}
			}
		}
	};
	ForEachSelectManager.prototype.processEndContent = function (contentInfo, count, callback) {
		const content = contentInfo.content;
		if (!content) {
			return this.processNullableContent(contentInfo, callback);
		}
		content.MoveCursorToStartPos();
		const targetPos = content.GetContentPosition(true, true);
		if (contentInfo.direction === AscWord.Direction.FORWARD) {
			const endPos = this.rebuildPos(contentInfo.endPos, content, contentInfo.endParagraph);
			content.SetContentSelection(targetPos, endPos, 0, 0, 0);
		} else {
			const endPos = this.rebuildPos(contentInfo.startPos, content, contentInfo.startParagraph);
			content.SetContentSelection(endPos, targetPos, 0, 0, 0);
		}
		const res = callback({content: content, index: count - 1, count: count});
		content.RemoveSelection();
		return res;
	};
	ForEachSelectManager.prototype.processMultipleSelectedContent = function (contents, callback) {
		const contentLength = contents.length;
		let res;
		res = this.processStartContent(contents[0], contentLength, callback);
		if (res) {
			return res;
		}
		res = this.processCenterContents(contents, callback);
		if (res) {
			return res;
		}
		return this.processEndContent(contents[contents.length - 1], contentLength, callback);
	};
	ForEachSelectManager.prototype.processSelectedContent = function (callback) {
		const contents = this.getSelectionUseContentsInfo();
		if (contents.length === 1) {
			const contentInfo = contents[0];
			return this.processSingleSelectedContent(contentInfo, callback);
		}
		if (contents.length > 1) {
			return this.processMultipleSelectedContent(contents, callback);
		}
	};
	ForEachSelectManager.prototype.processCursorContent = function (callback) {
		const content = this.getDocContent();
		const contentPos = content.GetContentPosition(false, false);
		const paragraph = content.Content[content.CurPos.ContentPos];
		const sourceParagraph = this.getSourceParagraph(paragraph);
		if (sourceParagraph) {
			return this.processRealCursorContent(sourceParagraph, contentPos, callback);
		}
		return this.processNullableContent({outlineParagraph: paragraph}, callback);
	};
	ForEachSelectManager.prototype.processRealCursorContent = function (sourceParagraph, contentPos, callback) {
		const sourceContent = sourceParagraph.GetParent();
		const startPos = this.rebuildPos(contentPos, sourceContent, sourceParagraph);
		sourceContent.SetContentPosition(startPos, 0, 0);
		const res = callback({content: sourceContent, index: 0, count: 1});
		sourceContent.RemoveSelection();
		return res;
	};
	ForEachSelectManager.prototype.rebuildPos = function (pos, sourceContent, sourceParagraph) {
		const newPos = pos.slice();
		newPos[0] = {Class: sourceContent, Position: sourceParagraph.GetIndex()};
		newPos[1] = {Class: sourceParagraph, Position: pos[1].Position};
		return newPos;
	};
	ForEachSelectManager.prototype.forEachSelectedContent = function (callback) {
		const content = this.getDocContent();
		if (!content) {
			return;
		}
		if (content.IsSelectionUse()) {
			return this.processSelectedContent(callback);
		}
		return this.processCursorContent(callback);
	};

	function RemoveOutlineParagraphsManager(outlineView) {
		this.outlineView = outlineView;
		this.checkSlidesForRemove = {};
	}
	RemoveOutlineParagraphsManager.prototype.getDocContent = function () {
		return this.outlineView.getDocContent();
	};
	RemoveOutlineParagraphsManager.prototype.getSourceParagraph = function (outlineParagraph) {
		return this.outlineView.outlineToSourceMap[outlineParagraph.Get_Id()];
	};
	RemoveOutlineParagraphsManager.prototype.getContentPosParagraph = function (docContent) {
		return docContent.Content[docContent.CurPos.ContentPos];
	};
	RemoveOutlineParagraphsManager.prototype.processNotSelectedContent = function () {
		if (this.isNeedProcessCursorAtBegin()) {
			this.processCursorAtBegin();
		} else {
			this.processCommonPosContent();
		}
	};
	RemoveOutlineParagraphsManager.prototype.processCommonPosContent = function () {
		const oThis = this;
		const outlineView = this.outlineView;
		outlineView.forEachSelectedContent(function (selectProps) {
			const content = selectProps.content;
			if (content) {
				content.Remove(-1, true, false, false);
				oThis.saveContentPosSourceParagraph(content);
			}
		});
	};
	RemoveOutlineParagraphsManager.prototype.processCursorAtBegin = function () {
		const docContent = this.getDocContent();
		const contentPos = docContent.GetContentPosition(false, true);
		const outlineParagraph = contentPos[1].Class;
		if (outlineParagraph.Index === 0) {
			return;
		}
		const sourceParagraph = this.getSourceParagraph(outlineParagraph);
		if (sourceParagraph) {
			this.processRealCursorAtBegin(outlineParagraph);
		} else {
			this.processNullableContent(outlineParagraph, true);
		}
	};
	RemoveOutlineParagraphsManager.prototype.processRealCursorAtBegin = function (outlineParagraph) {
		const docContent = this.getDocContent();
		const sourceParagraph = this.getSourceParagraph(outlineParagraph);
		const currentParagraphContent = sourceParagraph.GetParent();
		const previousOutlineParagraph = docContent.Content[outlineParagraph.Index - 1];
		const previousSourceParagraph = this.getSourceParagraph(previousOutlineParagraph);
		if (previousSourceParagraph) {
			const previousParagraphContent = previousSourceParagraph.GetParent();
			this.checkRemoveEdgeContents(previousParagraphContent, currentParagraphContent);
		} else {
			this.processNullableContent(previousOutlineParagraph, true);
		}
	};
	RemoveOutlineParagraphsManager.prototype.isNeedProcessCursorAtBegin = function () {
		const docContent = this.getDocContent();
		const paragraph = this.getContentPosParagraph(docContent);
		return paragraph.IsCursorAtBegin();
	};
	RemoveOutlineParagraphsManager.prototype.processSourceSelectedContent = function (content, idx, count, isOnAddText, edgeContents) {
		if (content.IsSelectedAll() || content.IsEmpty()) {
			const shape = content.Is_DrawingShape(true);
			const slide = shape.parent;
			const slideNum = slide.num;
			this.checkSlidesForRemove[slideNum] = slide;
			shape.deleteDrawingBase();
			content.RemoveSelection();
		} else {
			content.Remove(1, true, false, count === 1 && isOnAddText);
			if (idx === 0) {
				edgeContents.firstContent = content;
			} else if (idx === count - 1) {
				edgeContents.lastContent = content;
			}
		}
	};
	RemoveOutlineParagraphsManager.prototype.processSelectedContent = function (isOnAddText) {
		const outlineView = this.outlineView;
		const edgeContents = {firstContent: null, lastContent: null};
		const oThis = this;
		let isSingleContent = false;
		outlineView.forEachSelectedContent(function (selectProps) {
			if (selectProps.outlineParagraph) {
				const outlineParagraph = selectProps.outlineParagraph;
				oThis.processNullableContent(outlineParagraph);
			} else {
				const content = selectProps.content;
				const idx = selectProps.index;
				const count = selectProps.count;
				oThis.processSourceSelectedContent(content, idx, count, isOnAddText, edgeContents);
				if (count === 1) {
					isSingleContent = true;
					oThis.saveContentPosSourceParagraph(content);
				}
			}
		});

		if (isSingleContent) {
			return;
		}

		if (edgeContents.firstContent && edgeContents.lastContent) {
			this.checkRemoveEdgeContents(edgeContents.firstContent, edgeContents.lastContent);
		} else if (edgeContents.firstContent) {
			const content = edgeContents.firstContent.Content;
			this.saveEndPosSourceParagraph(content[content.length - 1]);
		} else if (edgeContents.lastContent) {
			const content = edgeContents.lastContent.Content;
			this.saveStartPosSourceParagraph(content[0]);
		}
	};
	RemoveOutlineParagraphsManager.prototype.remove = function (isOnAddText) {
		const docContent = this.getDocContent();
		if (isOnAddText) {
			if (!docContent.IsSelectionUse()) {
				return;
			}
		}
		const outlineView = this.outlineView;
		if (docContent.IsSelectionUse()) {
			this.processSelectedContent(isOnAddText);
		} else {
			this.processNotSelectedContent();
		}
		outlineView.removeSlides(this.checkSlidesForRemove);
	};
	RemoveOutlineParagraphsManager.prototype.saveEndPosParagraph = function (outlineParagraph) {
		const sourceParagraph = this.getSourceParagraph(outlineParagraph);
		if (sourceParagraph) {
			this.saveEndPosSourceParagraph(sourceParagraph);
		}
	};
	RemoveOutlineParagraphsManager.prototype.saveContentPosSourceParagraph = function (sourceContent) {
		const content = sourceContent;
		const contentPos = content.GetContentPosition(false, true);
		this.outlineView.setSavedPosition(new SavedPosition(contentPos));
	};
	RemoveOutlineParagraphsManager.prototype.saveEndPosSourceParagraph = function (sourceParagraph) {
		const contentPos = [{Class: sourceParagraph.GetParent(), Position: sourceParagraph.Index}];
		sourceParagraph.GetEndContentPosition(contentPos);
		this.outlineView.setSavedPosition(new SavedPosition(contentPos));
	};
	RemoveOutlineParagraphsManager.prototype.saveStartPosSourceParagraph = function (sourceParagraph) {
		const contentPos = [{Class: sourceParagraph.GetParent(), Position: sourceParagraph.Index}];
		sourceParagraph.GetStartContentPosition(contentPos);
		this.outlineView.setSavedPosition(new SavedPosition(contentPos));
	};
	RemoveOutlineParagraphsManager.prototype.processNullableContent = function (outlineParagraph, isEdgeParagraph) {
		const outlineView = this.outlineView;
		const info = outlineView.outlineInfo[outlineParagraph.Get_Id()];
		const slide = info && info.slide;
		if (!slide) {
			return;
		}
		this.checkSlidesForRemove[slide.num] = slide;
		outlineView.checkParagraph(outlineParagraph);

		if (isEdgeParagraph) {
			const docContent = this.getDocContent();
			const previousOutlineParagraph = docContent.Content[outlineParagraph.Index - 1];
			this.saveEndPosParagraph(previousOutlineParagraph);
		}
	};
	RemoveOutlineParagraphsManager.prototype.checkRemoveEdgeContents = function (firstContent, lastContent) {
		const firstShape = firstContent.Is_DrawingShape(true);
		const lastShape = lastContent.Is_DrawingShape(true);
		const firstSlide = firstShape.parent;
		const lastSlide = lastShape.parent;

		const firstContentParagraph = firstContent.Content[firstContent.Content.length - 1];
		this.saveEndPosSourceParagraph(firstContentParagraph);
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
					if (shape.isOutlinePlaceholderWithContent()) {
						const copyShape = shape.copy();
						copyShape.setBDeleted(false);
						copyShape.setParent(firstSlide);
						copyShape.addToDrawingObjects(firstSlide.cSld.spTree.length);
						shape.deleteDrawingBase();
					}
				} else {
					isLastShapeChecked = shape === lastShape;
				}
			});
		}
		if (!isSaveLastShape) {
			lastShape.deleteDrawingBase();
		}

	};

	function ParagraphAddManager(outlineView) {
		this.outlineView = outlineView;
	}
	ParagraphAddManager.prototype.forEachSelectedContent = function (callback) {
		return this.outlineView.forEachSelectedContent(callback);
	};
	ParagraphAddManager.prototype.savePositionAfterEdit = function (content, idx, contentCount) {
		return this.outlineView.savePositionAfterEdit(content, idx, contentCount);
	};
	ParagraphAddManager.prototype.processMockParagraphAdd = function (outlineParagraph, paraItem) {
		const info = this.outlineView.outlineInfo[outlineParagraph.Get_Id()];
		const slide = info && info.slide;
		if (slide) {
			const createdSp = slide.createTitle();
			createdSp.clearContent();
			const docContent = createdSp.getDocContent();
			docContent.Update_ContentIndexing();
			const sourceParagraph = docContent.Content[0];
			this.outlineView.removeOutlineParagraph(outlineParagraph);
			this.outlineView.addOutlineParagraph(sourceParagraph, outlineParagraph, info);
			const newOutlineParagraph = this.outlineView.updateFromSourceParagraph(sourceParagraph);
			newOutlineParagraph.Index = outlineParagraph.Index;
			newOutlineParagraph.MoveCursorToStartPos();
			const oThis = this;
			docContent.MoveCursorToStartPos();
			if (docContent) {
				oThis.processSourceContentAdd(docContent, 0, 1, paraItem);
			}
		}
	};
	ParagraphAddManager.prototype.processSourceContentAdd = function (content, idx, contentCount, paraItem) {
		content.AddToParagraph(paraItem);
		this.savePositionAfterEdit(content, idx, contentCount);
	};
	ParagraphAddManager.prototype.processParagraphAdd = function (paraItem) {
		const oThis = this;
		this.forEachSelectedContent(function (selectProps) {
			if (selectProps.outlineParagraph) {
				oThis.processMockParagraphAdd(selectProps.outlineParagraph, paraItem);
			} else {
				const content = selectProps.content;
				const idx = selectProps.index;
				const contentCount = selectProps.count;
				oThis.processSourceContentAdd(content, idx, contentCount, paraItem);
			}
		});
	};
	ParagraphAddManager.prototype.processRemoveContentTypes = function () {
		this.outlineView.remove(true);
	};
	ParagraphAddManager.prototype.paragraphAdd = function (paraItem) {
		switch (paraItem.Type) {
			case para_Math:
			case para_NewLine:
			case para_Text:
			case para_Space:
			case para_Tab:
			case para_PageNum:
			case para_Field:
			case para_FootnoteReference:
			case para_FootnoteRef:
			case para_Separator:
			case para_ContinuationSeparator:
			case para_InstrText:
			case para_EndnoteReference:
			case para_EndnoteRef: {
				this.processRemoveContentTypes();
			}
			default: {
				break;
			}
		}
		this.processParagraphAdd(paraItem);
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
			if (pr.slide !== undefined) {
				this.outlineInfo[outlineId] = {slide: pr.slide};
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
		return AscFormat.ExecuteNoHistory(function () {
			const outlineContent = this.getDocContent();
			const copyParagraph = this.getCopyParagraph(outlineContent, paragraph, !!(pr && pr.slide));
			outlineContent.AddToContent(pos, copyParagraph);
			this.addOutlineParagraph(paragraph, copyParagraph, isFirstParagraph ? pr : null);
			return copyParagraph;
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
	OutlineView.prototype.addMockTitleToOutlineShape = function (outlineShape, insertIndex, pr) {
		AscFormat.ExecuteNoHistory(function () {
			const outlineContent = outlineShape.txBody.content;
			const paragraph = this.getMockTitleParagraph(outlineContent);
			outlineContent.AddToContent(insertIndex, paragraph);
			this.addOutlineParagraph(null, paragraph, pr);
		}, this, []);

	};
	OutlineView.prototype.addMockTitle = function (insertIndex, pr) {
		this.addMockTitleToOutlineShape(this.outlineShape, insertIndex, pr);
	};
	OutlineView.prototype.getMockTitleParagraph = function (content) {
		return AscFormat.ExecuteNoHistory(function () {
			const paragraph = new AscWord.Paragraph(content, true);
			this.applyParagraphProps(paragraph, null, true);
			return paragraph;
		}, this, []);

	};
	OutlineView.prototype.fillOutlineShape = function (outlineShape, outlineSlides, width) {
		const paragraphs = outlineShape.txBody.content.Content;
		for (let i = 0; i < outlineSlides.length; i += 1) {
			const slide = outlineSlides[i];

			const titlePr = {slide: slide.slide};
			if (slide.title !== null) {
				this.addContentToOutlineShape(outlineShape, slide.title, titlePr);
			} else {
				this.addMockTitleToOutlineShape(outlineShape, paragraphs.length, titlePr);
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
			copyParaPr.Bullet.bulletColor = new AscFormat.CBulletColor();
			copyParaPr.Bullet.bulletColor.type = AscFormat.BULLET_TYPE_COLOR_CLR;
			copyParaPr.Bullet.bulletColor.UniColor = AscFormat.CreateUniColorFromRGB(0, 0, 0);
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
				if (info.slide !== undefined) {
					let penRGB;
					if (info.slide.num === currentSlideIndex) {
						penRGB = activeSlideRGB;
					} else if (info.slide.num === focusSlideIndex) {
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
				if (info && info.slide !== undefined) {
					return info.slide.num;
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
		const manager = new ForEachSelectManager(this);
		return manager.forEachSelectedContent(callback);
	}

	OutlineView.prototype.remove = function (bOnAddText) {
		const manager = new RemoveOutlineParagraphsManager(this);
		manager.remove(bOnAddText);
	};
	OutlineView.prototype.removeSlides = function (slideForRemoveMap) {
		const slides = [];
		const slideNumbers = Object.keys(slideForRemoveMap);
		for (let i = 0 ; i < slideNumbers.length; i += 1) {
			const num = slideNumbers[i];
			const slide = slideForRemoveMap[num];
			if (!slide.isHaveOutlineShapes()) {
				slides.push(slideForRemoveMap[num]);
			}
		}
		if (!slides.length) {
			return;
		}
		const presentation = this.getPresentation();
		let minSlideIndex = presentation.GetSlidesCount() - 1;
		for (let i = 0; i < slides.length; i++) {
			const index = presentation.GetSlideIndex(slides[i]);
			if (minSlideIndex > index) {
				minSlideIndex = index;
			}
		}

		const unpreserveInfo = presentation.getUnpreserveLayoutsAndMasters(slides);
		const checkArray = [].concat(unpreserveInfo.layouts, unpreserveInfo.masters, slides);
		for (let i = 0; i < checkArray.length; i += 1) {
			const slide = checkArray[i];
			if (!slide.isHaveOutlineShapes()) {
				presentation.removeSlideByObject(slide);
			}
		}
		presentation.updateSlideIndexes();
		presentation.DrawingDocument.UpdateThumbnailsAttack();
		presentation.DrawingDocument.m_oWordControl.GoToPage(Math.min(presentation.GetSlidesCount() - 1, minSlideIndex), undefined, undefined, true);
		presentation.Api.sync_HideComment();
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
				docContent.RemoveSelection();
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
		const manager = new ParagraphAddManager(this);
		manager.paragraphAdd(paraItem);
	}

	OutlineView.prototype.getParagraphParaPr = function () {
		let paraPr;
		this.forEachSelectedContent(function (selectProps) {
			if (selectProps.content) {
				const content = selectProps.content;
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
		this.forEachSelectedContent(function (selectProps) {
			if (selectProps.content) {
				const content = selectProps.content;
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
	OutlineView.prototype.unlinkOutlineParagraph = function (outlineParagraph) {
		const index = outlineParagraph.Index;
		this.removeParagraph(outlineParagraph, index);
	};
	OutlineView.prototype.getPropertiesFromSourceShape = function (sourceParagraph) {
		let pr = null;
		if (sourceParagraph.Index === 0) {
			const parentShape = sourceParagraph.GetParentShape();
			pr = {};
			if (parentShape.getPlaceholderType() === AscFormat.phType_ctrTitle || parentShape.getPlaceholderType() === AscFormat.phType_title) {
				pr.slide = parentShape.parent;
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
		return this.addUpdatedParagraph(sourceParagraph, index);
	};
	OutlineView.prototype.addUpdatedParagraph = function (sourceParagraph, index, pr) {
		pr = pr || this.getPropertiesFromSourceShape(sourceParagraph);
		return this.addCopyParagraph(sourceParagraph, index, sourceParagraph.Index === 0, pr);
	};

	OutlineView.prototype.addUpdatedParagraphWithCheckMockParagraph = function (sourceParagraph, index) {
		const docContent = this.getDocContent();
		const outlineParagraph = docContent.Content[index];
		if (outlineParagraph) {
			const pr = this.getPropertiesFromSourceShape(sourceParagraph);
			const info = this.outlineInfo[outlineParagraph.Get_Id()];
			if (info && pr && pr.slide === info.slide) {
				this.removeParagraph(outlineParagraph, index);
				this.addUpdatedParagraph(sourceParagraph, index, pr);
				return true;
			}
		}
		return false;
	};

	function UpdateData() {
		this.existingParagraphs = [];
		this.newParagraphs = [];
	}
	UpdateData.prototype.isNeedRecalculate = function () {
		return !!this.existingParagraphs.length || !!this.newParagraphs.length;
	};
	OutlineView.prototype.getUpdateData = function () {
		const updateData = new UpdateData();
		for (let id in this.mapToCheckParagraphs) {
			const paragraph = this.mapToCheckParagraphs[id];
			const shape = paragraph.GetParentShape();
			if (shape === this.outlineShape) {
				const sourceParagraph = this.outlineToSourceMap[paragraph.Get_Id()];
				if (!sourceParagraph) {
					updateData.existingParagraphs.push(paragraph)
				}
			} else if (shape && shape.isOutlinePlaceholder()) {
				if (shape.parent && shape.parent.getObjectType() === AscDFH.historyitem_type_Slide) {
					const outlineParagraph = this.sourceToOutlineMap[paragraph.Get_Id()];
					if (outlineParagraph) {
						updateData.existingParagraphs.push(outlineParagraph);
					} else if (paragraph.IsUseInDocument()) {
						updateData.newParagraphs.push(paragraph);
					}
				}
			}
		}
		this.resetMapToCheckParagraphs();
		return updateData;
	};
	OutlineView.prototype.update = function () {
		const updateData = this.getUpdateData();
		if (updateData.isNeedRecalculate()) {
			this.updateExistingParagraphs(updateData.existingParagraphs);
			this.updateNewParagraphs(updateData.newParagraphs);
			this.outlineShape && this.outlineShape.recalculateContent();
			this.applySavedPositionToOutline();
			this.resetPosition();
		}
	};
	function UseInDocumentManager() {
		this.useInDocumentShapes = {};
		this.useInDocumentParagraphs = {};
		this.useInDocumentSlides = null;
	}
	UseInDocumentManager.prototype.getPresentation = function () {
		return Asc.editor.WordControl.m_oLogicDocument;
	};
	UseInDocumentManager.prototype.isForceUseInDocumentParagraph = function (paragraph) {
		if (paragraph.Index === -1) {
			return false;
		}
		const shape = paragraph.GetParentShape();
		const isUseInDocumentShape = this.isUseInDocumentShape(shape);
		if (!isUseInDocumentShape) {
			return false;
		}
		const content = shape.getDocContent();
		if (content.Content[paragraph.Index] !== paragraph) {
			return false;
		}
		return true;
	};
	UseInDocumentManager.prototype.isUseInDocumentParagraph = function (paragraph) {
		if (!paragraph) {
			return false;
		}
		if (this.useInDocumentParagraphs[paragraph.GetId()] === undefined) {
			this.useInDocumentParagraphs[paragraph.GetId()] = this.isForceUseInDocumentParagraph(paragraph);
		}
		return this.useInDocumentParagraphs[paragraph.GetId()]
	};
	UseInDocumentManager.prototype.isUseInDocumentShape = function (shape) {
		if (!shape) {
			return false;
		}
		if (this.useInDocumentShapes[shape.GetId()] === undefined) {
			const slide = shape.parent;
			const isUseInDocumentSlide = this.isUseInDocumentSlide(slide);
			if (isUseInDocumentSlide) {
				for (let i = 0; i < slide.cSld.spTree.length; i += 1) {
					const sp = slide.cSld.spTree[i];
					this.useInDocumentShapes[sp.GetId()] = true;
				}
			}
		}
		return this.useInDocumentShapes[shape.GetId()];
	};
	UseInDocumentManager.prototype.isUseInDocumentSlide = function (slide) {
		if (!slide) {
			return false;
		}
		if (this.useInDocumentSlides === null) {
			this.useInDocumentSlides = {};
			const presentation = this.getPresentation();
			for (let i = 0; i < presentation.Slides.length; i += 1) {
				const slide = presentation.Slides[i];
				this.useInDocumentSlides[slide.GetId()] = true;
			}
		}
		return this.useInDocumentSlides[slide.GetId()];
	};
	function UpdateExistingParagraphManager(outlineView, useInDocumentManager) {
		this.outlineView = outlineView;
		this.useInDocumentManager = useInDocumentManager;
	}
	UpdateExistingParagraphManager.prototype.isNeedAddMockParagraph = function (sourceParagraph) {
		if (!sourceParagraph || sourceParagraph.Index !== 0) {
			return false;
		}
		const shape = sourceParagraph.GetParentShape();
		if (!shape.isOutlineTitlePlaceholder() || this.useInDocumentManager.isUseInDocumentShape(shape)) {
			return false;
		}

		const slide = shape.parent;
		if (!this.useInDocumentManager.isUseInDocumentSlide(slide)) {
			return false;
		}

		return true;
	};
	UpdateExistingParagraphManager.prototype.update = function (existingOutlineParagraphs) {
		existingOutlineParagraphs.sort(function (aParagraph, bParagraph) {
			return bParagraph.Index - aParagraph.Index;
		});
		const outlineView = this.outlineView;
		for (let i = 0; i < existingOutlineParagraphs.length; i += 1) {
			const outlineParagraph = existingOutlineParagraphs[i];
			const sourceParagraph = outlineView.outlineToSourceMap[outlineParagraph.Get_Id()];
			if (this.useInDocumentManager.isUseInDocumentParagraph(sourceParagraph)) {
				outlineView.updateFromSourceParagraph(sourceParagraph);
			} else {
				if (this.isNeedAddMockParagraph(sourceParagraph)) {
					const info = outlineView.outlineInfo[outlineParagraph.GetId()];
					outlineView.unlinkOutlineParagraph(outlineParagraph);
					outlineView.addMockTitle(outlineParagraph.Index, info)
				} else {
					outlineView.unlinkOutlineParagraph(outlineParagraph);
				}
			}
		}
	}
	OutlineView.prototype.updateExistingParagraphs = function (existingOutlineParagraphs) {
		const manager = new UpdateExistingParagraphManager(this, new UseInDocumentManager());
		manager.update(existingOutlineParagraphs);
	}

	OutlineView.prototype.updateNewParagraphs = function (newParagraphs) {
		const updateManager = new UpdateNewParagraphsManager(this, newParagraphs);
		updateManager.update();
	}
	OutlineView.prototype.checkParagraph = function (paragraph) {
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
				this.checkParagraph(paragraph);
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
		} else if (this.outlineInfo[outlineId] && this.outlineInfo[outlineId].slide !== undefined) {
			return this.outlineInfo[outlineId].slide.num;
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
				if (info && info.slide && info.slide.num === pageNum) {
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
