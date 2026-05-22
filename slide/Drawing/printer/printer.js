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

(function () {

	function executeWithContentLimits(callback, content, x, y, xLimit, yLimit) {
		const oldX = content.X;
		const oldY = content.Y;
		const oldXLimit = content.XLimit;
		const oldYLimit = content.YLimit;
		content.X = x;
		content.Y = y;
		content.XLimit = xLimit;
		content.YLimit = yLimit;
		callback();
		content.X = oldX;
		content.Y = oldY;
		content.XLimit = oldXLimit;
		content.YLimit = oldYLimit;
	}

	function NotePage(note, pageIndex) {
		this.note = note;
		this.pageIndex = pageIndex;
	}

	function PrintManager(presentation, printOptions) {
		this.printOptions = printOptions;
		this.presentation = presentation;
		this.specificPrinter = null;
		this.init();
	}
	PrintManager.prototype.init = function () {
		switch (this.getPrintType()) {
			case Asc.c_oAscSlidesOnPagePrintType.FullPageSlides: {
				this.specificPrinter = new SlidePrinter(this.presentation, this.printOptions);
				break;
			}
			case Asc.c_oAscSlidesOnPagePrintType.Handouts: {
				this.specificPrinter = new HandoutPrinter(this.presentation, this.printOptions);
				break;
			}
			case Asc.c_oAscSlidesOnPagePrintType.Outline: {
				this.specificPrinter = new OutlinePrinter(this.presentation, this.printOptions);
				break;
			}
			case Asc.c_oAscSlidesOnPagePrintType.SlideWithNotes: {
				this.specificPrinter = new NotesPrinter(this.presentation, this.printOptions);
				break;
			}
		}
	};
	PrintManager.prototype.getPrintType = function () {
		return this.printOptions.slidesOnPagePrintOptions.printType;
	};
	PrintManager.prototype.drawPage = function (graphics, index) {
		this.specificPrinter.drawPage(graphics, index);
	};
	PrintManager.prototype.getPagesCount = function () {
		return this.specificPrinter.getPagesCount();
	};
	PrintManager.prototype.getPageSizes = function () {
		return this.specificPrinter.getPageSizes();
	};
	const SPECIFICPRINTER_FITPAGE_MARGIN = 1;
	function SpecificPrinter(presentation, printOptions) {
		this.presentation = presentation;
		this.printOptions = printOptions;
		this.init();
	}
	SpecificPrinter.prototype.init = function () {
		this.resetCache();
	};
	SpecificPrinter.prototype.resetCache = function () {
		this.cache = Object.assign({
			printIndexes: null
		}, this.getSpecificCache());
	};
	SpecificPrinter.prototype.getSpecificCache = function () {
		return {};
	};

	SpecificPrinter.prototype.getPresentation = function () {
		return this.presentation;
	}
	SpecificPrinter.prototype.getForcePrintIndexes = function () {
		const range = this.printOptions.rangeOptions;
		const presentation = this.getPresentation();
		const isDrawHiddenSlides = this.printOptions.slidePrintOptions.asc_getIsPrintHiddenSlides();
		const slides = presentation.Slides;
		let result;
		switch (range.rangeType) {
			case Asc.c_oAscPresentationRangeType.AllSlides: {
				result = [];
				for (let i = 0; i < slides.length; i += 1) {
					result.push(i);
				}
				break;
			}
			case Asc.c_oAscPresentationRangeType.SelectedSlides: {
				if (presentation.IsSlidePageMode() || presentation.IsOutlineMode()) {
					const thumbnails = presentation.DrawingDocument.m_oWordControl.Thumbnails;
					result = thumbnails.GetSelectedArray();
				} else {
					result = [];
					for (let i = 0; i < slides.length; i += 1) {
						result.push(i);
					}
				}
				break;
			}
			case Asc.c_oAscPresentationRangeType.ActiveSlide: {
				if (presentation.CurPage !== -1) {
					result = [presentation.CurPage];
				} else {
					return [];
				}
				break;
			}
			case Asc.c_oAscPresentationRangeType.CustomRange: {
				const customRange = range.customRange;
				const slides = presentation.Slides;
				const ranges = customRange.trim().split(",");
				result = [];
				for (let i = 0; i < ranges.length; i += 1) {
					const range = ranges[i].trim().split("-");
					if (range.length === 2) {
						const start = parseInt(range[0], 10);
						const end = parseInt(range[1], 10);
						if (Number.isNaN(start) || Number.isNaN(end)) {
							return [];
						}
						if (start < 0 || start >= slides.length || end < 0 || end >= slides.length) {
							return [];
						}
						if (start <= end) {
							for (let slideIndex = start; slideIndex <= end; slideIndex += 1) {
								result.push(slideIndex);
							}
						} else {
							for (let slideIndex = end; end >= start; slideIndex -= 1) {
								result.push(slideIndex);
							}
						}
					} else if (range.length === 1) {
						const slideIndex = parseInt(range[0], 10);
						if (Number.isNaN(slideIndex) || slideIndex < 0 || slideIndex >= slides.length) {
							return [];
						}
						result.push(slideIndex);
					} else {
						return [];
					}
				}
				break;
			}
		}
		if (isDrawHiddenSlides) {
			return result;
		}
		return result.filter(function (slideIndex) {
			const slide = slides[slideIndex];
			return slide.isVisible();
		});
	};
	SpecificPrinter.prototype.getPrintIndexes = function () {
		if (this.cache.printIndexes === null) {
			this.cache.printIndexes = this.getForcePrintIndexes();
		}
		return this.cache.printIndexes;
	};
	SpecificPrinter.prototype.getPagesCount = function () {
		return 0;
	};
	SpecificPrinter.prototype.getContentSizes = function () {
		return null;
	};
	SpecificPrinter.prototype.getPageSizes = function () {
		const pageOptions = this.printOptions.pageOptions;
		return { width: pageOptions.width, height: pageOptions.height };
	};
	SpecificPrinter.prototype.isScaleToFitPaper = function () {
		return this.printOptions.slidesOnPageOptions.asc_getIsScaleToFitPaper();
	};
	SpecificPrinter.prototype.drawPage = function (graphics, index) {
		const pageSizes = this.getPageSizes();
		const contentSizes = this.getContentSizes();
		const margin = this.isScaleToFitPaper() ? SPECIFICPRINTER_FITPAGE_MARGIN : 0;
		const scale = this.isScaleToFitPaper()
			? Math.min((pageSizes.width - 2 * margin) / contentSizes.width, (pageSizes.height - 2 * margin) / contentSizes.height)
			: 1;
		const m = new AscCommon.CMatrix();
		m.Scale(scale, scale);
		m.Translate((pageSizes.width - contentSizes.width * scale) / 2, (pageSizes.height - contentSizes.height * scale) / 2);
		const baseTransform = graphics.GetBaseTransform();
		const composedTransform = baseTransform ? baseTransform.CreateDublicate() : new AscCommon.CMatrix();
		composedTransform.Multiply(m, AscCommon.MATRIX_ORDER_PREPEND);

		graphics.SaveGrState();
		graphics.SetIntegerGrid(false);
		graphics.AddClipRect(0, 0, pageSizes.width, pageSizes.height);
		graphics.SetBaseTransform(composedTransform);
		graphics.reset();
		graphics.SetIntegerGrid(false);
		this.draw(graphics, index);
		graphics.SetBaseTransform(baseTransform);
		graphics.reset();
		graphics.RestoreGrState();

		if (this.isDrawFrame()) {
			graphics.SaveGrState();
			graphics.SetBaseTransform(composedTransform);
			graphics.reset();
			graphics.SetIntegerGrid(false);
			this.drawFrame(graphics);
			graphics.SetBaseTransform(baseTransform);
			graphics.reset();
			graphics.RestoreGrState();
		}
	};
	SpecificPrinter.prototype.draw = function (graphics, index) {

	};
	SpecificPrinter.prototype.isDrawFrame = function () {
		return this.printOptions.slidesOnPageOptions.asc_getIsFrameSlides();
	};
	SpecificPrinter.prototype.drawFrame = function (graphics) {
		const contentSizes = this.getContentSizes();
		graphics.p_color(0, 0, 0, 255);
		this.strokeRect(graphics, 0, 0, contentSizes.width, contentSizes.height);
	};
	SpecificPrinter.prototype.strokeRect = function (graphics, x, y, width, height) {
		if (graphics.AddSmartRect) {
			graphics.AddSmartRect(x, y, width, height, 0);
		} else {
			graphics.p_width(0);
			graphics._s();
			graphics._m(x, y);
			graphics._l(x + width, y);
			graphics._l(x + width, y + height);
			graphics._l(x, y + height);
			graphics._z();
			graphics.ds();
		}
	};

	const NOTESPRINTER_HORIZONTAL_FIELD = 20;
	const NOTESPRINTER_VERTICAL_FIELD = 15;
	function NotesPrinter(presentation, printOptions) {
		SpecificPrinter.call(this, presentation, printOptions);
	}
	AscFormat.InitClassWithoutType(NotesPrinter, SpecificPrinter);
	NotesPrinter.prototype.getSpecificCache = function () {
		return {
			pages: null
		};
	};
	NotesPrinter.prototype.drawPlaceholders = function (note, graphics) {
		note.cSld.forEachSp(function (shape) {
			const placeholderType = shape.getPlaceholderType();
			switch (placeholderType) {
				case AscFormat.phType_dt:
				case AscFormat.phType_sldNum:
				case AscFormat.phType_ftr:
				case AscFormat.phType_hdr: {
					shape.draw(graphics);
				}
			}
		});
	};
	NotesPrinter.prototype.draw = function (graphics, index) {
		const pages = this.getNotesPages();
		const page = pages[index];
		const note = page.note;
		const contentSizes = this.getContentSizes();
		graphics.SaveGrState();
		graphics.AddClipRect(0, 0, contentSizes.width, contentSizes.height);
		if (page.pageIndex === 0) {
			note.draw(graphics);
		} else {
			graphics.SaveGrState();
			this.drawPlaceholders(note, graphics);
			const transform = new AscCommon.CMatrix();
			transform.tx = NOTESPRINTER_HORIZONTAL_FIELD;
			transform.ty = NOTESPRINTER_VERTICAL_FIELD;
			graphics.transform3(transform);
			const notesShape = note.getBodyShape();
			const docContent = notesShape.getDocContent();
			docContent.Set_StartPage(0);
			docContent.Draw(page.pageIndex, graphics);
			graphics.RestoreGrState();
		}
		graphics.RestoreGrState();
	};
	NotesPrinter.prototype.getNotesPages = function () {
		if (this.cache.pages === null) {
			const pages = [];
			const printIndexes = this.getPrintIndexes();
			const presentation = this.getPresentation();
			const notesPageHeight = presentation.GetNotesHeightMM() - NOTESPRINTER_VERTICAL_FIELD * 2;
			const notesPageWidth = presentation.GetNotesWidthMM() - NOTESPRINTER_HORIZONTAL_FIELD * 2;
			const slides = presentation.Slides;
			for (let i = 0; i < printIndexes.length; i += 1) {
				const index = printIndexes[i];
				const slide = slides[index];
				const note = slide.notes;
				const shape = note.getBodyShape();
				note.recalculate();
				if (shape) {
					note.cSld.forEachSp(function (shape) {
						shape.setRecalculateInfo();
						shape.recalculate()
					});
					shape.setRecalculateInfo();
					shape.recalculate();
					var oDocContent = shape.getDocContent();
					if(oDocContent) {
						oDocContent.CalculateAllFields();
						oDocContent.Reset(0, 0, shape.extX, 20000);
						var CurPage = 0;
						var RecalcResult = recalcresult2_NextPage;
						while (recalcresult2_End !== RecalcResult) {
							const callback = function () {
								pages.push(new NotePage(note, CurPage));
								RecalcResult = oDocContent.Recalculate_Page(CurPage++, true);
							};
							if (CurPage === 0) {
								executeWithContentLimits(callback, oDocContent, 0, 0, shape.extX, shape.extY);
							} else {
								executeWithContentLimits(callback, oDocContent, 0, 0, notesPageWidth, notesPageHeight);
							}
						}
						shape.contentWidth = shape.extX;
						shape.contentHeight = 20000;
					}
				}
			}
			this.cache.pages = pages;
		}
		return this.cache.pages;
	};
	NotesPrinter.prototype.getPagesCount = function () {
		return this.getNotesPages().length;
	};
	NotesPrinter.prototype.getContentSizes = function () {
		const presentation = this.getPresentation();
		return { width: presentation.GetNotesWidthMM(), height: presentation.GetNotesHeightMM() };
	};

	const OUTLINEPRINTER_DECORATIONS_OFFSET_LEFT = 15;
	function OutlinePrinter(presentation, printOptions) {
		SpecificPrinter.call(this, presentation, printOptions);
	}
	AscFormat.InitClassWithoutType(OutlinePrinter, SpecificPrinter);

	OutlinePrinter.prototype.getSpecificCache = function () {
		return {
			pagesCount: null,
			outlineView: null,
			decorationsByPage: null
		};
	};
	OutlinePrinter.prototype.getPagesCount = function () {
		if (this.cache.pagesCount === null) {
			const outlineShape = this.getOutlineShape();
			const docContent = outlineShape.getDocContent();
			this.cache.pagesCount = docContent.Pages.length;
		}
		return this.cache.pagesCount;
	};
	OutlinePrinter.prototype.getPageOptions = function () {
		const pageOptions = this.printOptions.pageOptions;
		return {width: pageOptions.width - NOTESPRINTER_HORIZONTAL_FIELD * 2 - OUTLINEPRINTER_DECORATIONS_OFFSET_LEFT, height: pageOptions.height - NOTESPRINTER_VERTICAL_FIELD * 2};
	};
	OutlinePrinter.prototype.getOutlineView = function () {
		if (this.cache.outlineView === null) {
			const pageOptions = this.getPageOptions();
			const indexes = this.getPrintIndexes();
			const presentation = this.getPresentation();
			const slides = presentation.Slides;
			const outlineSlides = [];
			for (let i = 0; i < indexes.length; i++) {
				const slide = slides[indexes[i]];
				outlineSlides.push(slide.getOutlineSlide());
			}
			const outlineView = new AscCommonSlide.OutlineView(presentation.Api);
			outlineView.createOutlineShape(outlineSlides, pageOptions.width, pageOptions.height);
			this.cache.outlineView = outlineView;
			this.recalculateOutlinePages();
		}
		return this.cache.outlineView;
	};
	OutlinePrinter.prototype.getOutlineShape = function () {
		const outlineView = this.getOutlineView();
		return outlineView.outlineShape;
	};
	OutlinePrinter.prototype.recalculateOutlinePages = function () {
		const outlineShape = this.getOutlineShape();
		const content = outlineShape.getDocContent();
		let recalResult = recalcresult2_NextPage;
		let curPage = 0;
		while (recalcresult2_End !== recalResult)
			executeWithContentLimits(function () {
				recalResult = content.Recalculate_Page(curPage++, true);
			}, content, 0, 0, outlineShape.extX, outlineShape.extY);
	};
	OutlinePrinter.prototype.getParagraphPage = function (paragraph) {
		const startPage = paragraph.PageNum;
		for (let i = 0; i < paragraph.Pages.length; i += 1) {
			const page = paragraph.Pages[i];
			if (page.EndLine >= 0) {
				return startPage + i;
			}
		}
		return startPage;
	};
	OutlinePrinter.prototype.getDecorationsByPage = function () {
		if (this.cache.decorationsByPage === null) {
			this.cache.decorationsByPage = {
				titleCache: {},
				contentCache: {}
			};
			const titleCache = this.cache.decorationsByPage.titleCache;
			const contentCache = this.cache.decorationsByPage.contentCache;
			const outlineView = this.getOutlineView();
			const paragraphMap = outlineView.outlineInfo.getOutlineParagraphToInfoMap();
			for (let outlineId in paragraphMap) {
				const info = paragraphMap[outlineId];
				const titleOutlineParagraph = info.outlineParagraph;
				const titlePageNum = this.getParagraphPage(titleOutlineParagraph);
				if (!titleCache[titlePageNum]) {
					titleCache[titlePageNum] = [];
				}
				titleCache[titlePageNum].push(info);
				const contentShapeInfoMap = info.getContentShapeInfoMap(outlineView);
				for (let contentShapeInfoId in contentShapeInfoMap) {
					const contentShapeInfo = contentShapeInfoMap[contentShapeInfoId];
					const contentOutlineParagraph = contentShapeInfo.getOutlineParagraph(outlineView);
					const contentPageNum = this.getParagraphPage(contentOutlineParagraph);
					if (!contentCache[contentPageNum]) {
						contentCache[contentPageNum] = [];
					}
					contentCache[contentPageNum].push(contentShapeInfo);
				}
			}
		}
		return this.cache.decorationsByPage;
	};
	OutlinePrinter.prototype.drawDecorationsByPage = function (graphics, index) {
		const t = new AscCommon.CMatrix();
		t.tx = NOTESPRINTER_HORIZONTAL_FIELD;
		t.ty = NOTESPRINTER_VERTICAL_FIELD;
		const decorationsInfo = this.getDecorationsByPage();
		const outlineView = this.getOutlineView();

		const decoratorDrawer = new AscCommonSlide.DecoratorDrawer(outlineView, t);
		const titlePageInfos = decorationsInfo.titleCache[index];
		const contentPageInfos = decorationsInfo.contentCache[index];
		if (titlePageInfos) {
			for (let i = 0; i < titlePageInfos.length; i += 1) {
				const info = titlePageInfos[i];
				decoratorDrawer.drawTitleDecorations(graphics, info);
			}
		}
		if (contentPageInfos) {
			for (let i = 0; i < contentPageInfos.length; i += 1) {
				const info = contentPageInfos[i];
				decoratorDrawer.drawContentDecorations(graphics, info);
			}
		}
	};
	OutlinePrinter.prototype.getContentSizes = function () {
		const pageOptions = this.printOptions.pageOptions;
		return { width: pageOptions.width, height: pageOptions.height };
	};
	OutlinePrinter.prototype.draw = function (graphics, index) {
		const pageSize = this.getPageSizes();
		graphics.SaveGrState();
		graphics.AddClipRect(0, 0, pageSize.width, pageSize.height);
		this.drawDecorationsByPage(graphics, index);
		const t = new AscCommon.CMatrix();
		t.tx = NOTESPRINTER_HORIZONTAL_FIELD + OUTLINEPRINTER_DECORATIONS_OFFSET_LEFT;
		t.ty = NOTESPRINTER_VERTICAL_FIELD;
		graphics.transform3(t);
		const shape = this.getOutlineShape();
		const content = shape.getDocContent();
		content.Set_StartPage(0);
		content.Draw(index, graphics);
		graphics.RestoreGrState();
	};
	OutlinePrinter.prototype.isDrawFrame = function () {
		return false;
	};

	function HandoutPrinter(presentation, printOptions) {
		SpecificPrinter.call(this, presentation, printOptions);
	}
	AscFormat.InitClassWithoutType(HandoutPrinter, SpecificPrinter);
	HandoutPrinter.prototype.getSpecificCache = function () {
		return {
			pagesCount: null
		};
	};
	HandoutPrinter.prototype.getSlidesOnPageCount = function () {
		return this.printOptions.slidesOnPagePrintOptions.slidesCount;
	};
	HandoutPrinter.prototype.getPagesCount = function () {
		if (this.cache.pagesCount === null) {
			const printIndexes = this.getPrintIndexes();
			this.cache.pagesCount = Math.ceil(printIndexes.length / this.getHandoutSlidesCount());
		}
		return this.cache.pagesCount;
	};
	HandoutPrinter.prototype.getPageSlides = function (pageIndex) {
		const presentation = this.getPresentation();
		const slides = presentation.Slides;
		const slidesOnPageCount = this.getSlidesOnPageCount();
		const printIndexes = this.getPrintIndexes();
		const printSlides = [];
		const startSlideIndex = pageIndex * slidesOnPageCount;
		const endSlideIndex = (pageIndex + 1) * slidesOnPageCount;
		for (let i = startSlideIndex; i < endSlideIndex; i += 1) {
			const index = printIndexes[i];
			const slide = slides[index];
			if (!slide) {
				break;
			}
			printSlides.push(slide);
		}
		return printSlides;
	};
	HandoutPrinter.prototype.getHandoutSettings = function () {
		const printOptions = this.printOptions;
		const slidesOnPagePrintOptions = printOptions.slidesOnPagePrintOptions;
		return slidesOnPagePrintOptions.getHandoutSettings();
	};
	HandoutPrinter.prototype.draw = function (graphics, index) {
		const pageSlides = this.getPageSlides(index);
		const presentation = this.getPresentation();
		const handoutMaster = presentation.handoutMasters[0];
		const contentSizes = this.getContentSizes();
		graphics.SaveGrState();
		graphics.AddClipRect(0, 0, contentSizes.width, contentSizes.height);
		const oldHandoutSettings = handoutMaster.handoutSettings;
		handoutMaster.handoutSettings = this.getHandoutSettings();
		handoutMaster.draw(graphics, null, pageSlides, this.isDrawSlideFrame(), this.isDrawSlideNumber());
		handoutMaster.handoutSettings = oldHandoutSettings;
		graphics.RestoreGrState();
	};
	HandoutPrinter.prototype.getContentSizes = function () {
		const presentation = this.getPresentation();
		return { width: presentation.GetNotesWidthMM(), height: presentation.GetNotesHeightMM() };
	};
	HandoutPrinter.prototype.isDrawFrame = function () {
		return false;
	};
	HandoutPrinter.prototype.isDrawSlideFrame = function () {
		return this.printOptions.slidesOnPageOptions.asc_getIsFrameSlides();
	};
	HandoutPrinter.prototype.isDrawSlideNumber = function () {
		return this.printOptions.slidesOnPageOptions.asc_getIsPrintSlideNumbersOnHandouts();
	};

	function SlidePrinter(presentation, printOptions) {
		SpecificPrinter.call(this, presentation, printOptions);
	}
	AscFormat.InitClassWithoutType(SlidePrinter, SpecificPrinter);
	SlidePrinter.prototype.draw = function (graphics, index) {
		const presentation = this.getPresentation();
		const printIndexes = this.getPrintIndexes();
		const slideIndex = printIndexes[index];
		const slide = presentation.Slides[slideIndex];
		const contentSizes = this.getContentSizes();
		graphics.SaveGrState();
		graphics.AddClipRect(0, 0, contentSizes.width, contentSizes.height);
		slide.draw(graphics);
		graphics.RestoreGrState();
	};
	SlidePrinter.prototype.getPagesCount = function () {
		return this.getPrintIndexes().length;
	};
	SlidePrinter.prototype.getContentSizes = function () {
		const presentation = this.getPresentation();
		return { width: presentation.GetWidthMM(), height: presentation.GetHeightMM() };
	};
	SlidePrinter.prototype.getPageSizes = function () {
		const pageOptions = this.printOptions.pageOptions;
		const contentSizes = this.getContentSizes();
		const isSlideLandscape = contentSizes.width > contentSizes.height;
		const isPageLandscape = pageOptions.width > pageOptions.height;
		if (isSlideLandscape !== isPageLandscape) {
			return { width: pageOptions.height, height: pageOptions.width };
		}
		return { width: pageOptions.width, height: pageOptions.height };
	};

	window["AscCommonSlide"] = window["AscCommonSlide"] || {};
	window["AscCommonSlide"].PrintManager = PrintManager;
})();
