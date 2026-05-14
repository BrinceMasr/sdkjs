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
		this.specificPrinter.getPagesCount();
	};

	PrintManager.prototype.printFullPage = function (index) {
		printPreview.drawOnPaper(index, {width:sizes.width, height:sizes.height, offset: 0}, {
			width:  sizes.width,
			height: sizes.height
		}, renderer, printPreview.drawFullPageSlide.bind(printPreview));
	};

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
				if (presentation.IsSlidePageMode()) {
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
		return result.filter(function (slide) {
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
	SpecificPrinter.prototype.drawPage = function (graphics, index) {

	};
	function NotesPrinter(presentation, printOptions) {
		SpecificPrinter.call(this, presentation, printOptions);
	}
	AscFormat.InitClassWithoutType(NotesPrinter, SpecificPrinter);
	NotesPrinter.prototype.getSpecificCache = function () {
		return {
			pages: null
		};
	};
	NotesPrinter.prototype.drawPage = function (graphics, index) {
		const pages = this.getNotesPages();
		const page = pages[index];
		if (page.pageIndex === 0) {

		} else {
			const notesShape = page.note.getBodyShape();
			const docContent = notesShape.getDocContent();
			docContent.Set_StartPage(0);
			docContent.Draw(page.pageIndex, graphics);
		}
	};
	NotesPrinter.prototype.getNotesPages = function () {
		if (this.cache.pages === null) {
			const pages = [];
			const printIndexes = this.getPrintIndexes();
			const presentation = this.getPresentation();
			const notesHeight = presentation.GetNotesHeightMM();
			const slides = presentation.Slides;
			for (let i = 0; i < printIndexes.length; i += 1) {
				const index = printIndexes[i];
				const slide = slides[index];
				const note = slide.notes;
				const shape = note.getBodyShape();
				if (shape) {
					var oDocContent = shape.getDocContent();
					if(oDocContent) {
						const rect = shape.getTextRect();
						const oldGetColumnContentFrame = shape.GetColumnContentFrame;
						shape.GetColumnContentFrame = function (page, column, sectPr) {
							if (page === 0) {
								return {X: 0, Y: 0, XLimit: rect.r - rect.l, YLimit: rect.b - rect.t};
							}
							return {X: 0, Y: 0, XLimit: rect.r - rect.l, YLimit: notesHeight};
						}
						oDocContent.CalculateAllFields();
						var Width = shape.spPr.xfrm.extX;
						oDocContent.Reset(0, 0, Width, 20000);
						var CurPage = 0;
						var RecalcResult = recalcresult2_NextPage;
						while (recalcresult2_End !== RecalcResult) {
							pages.push(new NotePage(note, CurPage));
							RecalcResult = oDocContent.Recalculate_Page(CurPage++, true);
						}
						shape.contentWidth = Width;
						shape.contentHeight = 20000;
						shape.GetColumnContentFrame = oldGetColumnContentFrame;
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

	function OutlinePrinter(presentation, printOptions) {
		SpecificPrinter.call(this, presentation, printOptions);
	}
	AscFormat.InitClassWithoutType(OutlinePrinter, SpecificPrinter);

	OutlinePrinter.prototype.getSpecificCache = function () {
		return {
			pagesCount: null,
			outlineShape: null
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
		const printerOptions = this.printOptions;
		return printerOptions.pageOptions;
	};
	OutlinePrinter.prototype.getOutlineShape = function () {
		if (this.cache.outlineShape === null) {
			const pageOptions = this.getPageOptions();
			const indexes = this.getPrintIndexes();
			const presentation = this.getPresentation();
			const slides = presentation.Slides;
			const outlineSlides = [];
			for (let i = 0; i < indexes.length; i++) {
				const slide = slides[indexes[i]];
				outlineSlides.push(slide.getOutlineSlide());
			}
			const outlineView = new AscCommonSlide.OutlineView();
			this.cache.outlineShape = outlineView.createOutlineShape(outlineSlides, pageOptions.width, pageOptions.height);
			this.recalculateOutlinePages();
		}
		return this.cache.outlineShape;
	};
	OutlinePrinter.prototype.recalculateOutlinePages = function () {
		const outlineShape = this.getOutlineShape();
		outlineShape.recalculateDocContent();
	}
	OutlinePrinter.prototype.drawPage = function (graphics, index) {
		const shape = this.getOutlineShape();
		const content = shape.getDocContent();
		content.Set_StartPage(0);
		content.Draw(index, graphics);
	};

	function HandoutPrinter(presentation, printOptions) {
		SpecificPrinter.call(this, presentation, printOptions);
	}
	AscFormat.InitClassWithoutType(HandoutPrinter, SpecificPrinter);
	HandoutPrinter.prototype.getSpecificCache = function () {
		return {
			pagesCount: null,
			outlineShape: null
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
	HandoutPrinter.prototype.drawPage = function (graphics, index) {
		const pageSlides = this.getPageSlides(index);
	};

	function SlidePrinter(presentation, printOptions) {
		SpecificPrinter.call(this, presentation, printOptions);
	}
	AscFormat.InitClassWithoutType(SlidePrinter, SpecificPrinter);
	SlidePrinter.prototype.drawPage = function (graphics, index) {
		const presentation = this.getPresentation();
		const printIndexes = this.getPrintIndexes();
		const slideIndex = printIndexes[index];
		const slides = presentation.Slides;
		const slide = slides[slideIndex];

	}
	SlidePrinter.prototype.getPagesCount = function () {
		return this.getPrintIndexes().length;
	};
})();
