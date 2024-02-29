import { WordDocument } from './word-document';
import {
	DomType,
	WmlImage,
	IDomNumbering,
	OpenXmlElement,
	WmlBreak,
	WmlDrawing,
	WmlHyperlink,
	WmlNoteReference,
	WmlSymbol,
	WmlTable,
	WmlTableCell,
	WmlTableColumn,
	WmlTableRow,
	WmlText,
	WrapType,
} from './document/dom';
import { CommonProperties } from './document/common';
import { Options } from './docx-preview';
import { DocumentElement } from './document/document';
import { WmlParagraph } from './document/paragraph';
import { asArray, escapeClassName, isString, keyBy, mergeDeep, uuid } from './utils';
import { computePixelToPoint, updateTabStop } from './javascript';
import { FontTablePart } from './font-table/font-table';
import { FooterHeaderReference, SectionProperties, SectionType } from './document/section';
import { Page, PageProps } from './document/page';
import { RunProperties, WmlRun } from './document/run';
import { WmlBookmarkStart } from './document/bookmarks';
import { IDomStyle } from './document/style';
import { WmlBaseNote, WmlFootnote } from './notes/elements';
import { ThemePart } from './theme/theme-part';
import { BaseHeaderFooterPart } from './header-footer/parts';
import { Part } from './common/part';
import { VmlElement } from './vml/vml';
import { WmlCommentRangeStart, WmlCommentReference } from './comments/elements';
import Konva from 'konva';
import type { Stage } from 'konva/lib/Stage';
import type { Layer } from 'konva/lib/Layer';
import type { Group } from 'konva/lib/Group';

const ns = {
	html: 'http://www.w3.org/1999/xhtml',
	svg: 'http://www.w3.org/2000/svg',
	mathML: 'http://www.w3.org/1998/Math/MathML',
};

interface CellPos {
	col: number;
	row: number;
}

type CellVerticalMergeType = Record<number, HTMLTableCellElement>;

interface Node_DOM extends Node, Comment, CharacterData {
	dataset: DOMStringMap;
}

enum Overflow {
	TRUE = 'true',
	FALSE = 'false',
	UNKNOWN = 'undetected',
}

// HTML渲染器
export class HtmlRendererSync {
	className = 'docx';
	rootSelector: string;
	document: WordDocument;
	options: Options;
	styleMap: Record<string, IDomStyle> = {};
	currentPart: Part = null;
	wrapper: HTMLElement;

	// 当前操作的Page
	currentPage: Page;

	tableVerticalMerges: CellVerticalMergeType[] = [];
	currentVerticalMerge: CellVerticalMergeType = null;
	tableCellPositions: CellPos[] = [];
	currentCellPosition: CellPos = null;

	footnoteMap: Record<string, WmlFootnote> = {};
	endnoteMap: Record<string, WmlFootnote> = {};
	currentFootnoteIds: string[];
	currentEndnoteIds: string[] = [];
	usedHederFooterParts: any[] = [];

	defaultTabSize: string;
	// 当前制表位
	currentTabs: any[] = [];
	tabsTimeout: any = 0;

	// Konva框架--stage元素
	konva_stage: Stage;
	// Konva框架--layer元素
	konva_layer: Layer;

	/**
	 * Object对象 => HTML标签
	 *
	 * @param document word文档Object对象
	 * @param bodyContainer HTML生成容器
	 * @param styleContainer CSS样式生成容器
	 * @param options 渲染配置选项
	 */

	async render(document: WordDocument, bodyContainer: HTMLElement, styleContainer: HTMLElement = null, options: Options) {
		this.document = document;
		this.options = options;
		// class类前缀
		this.className = options.className;
		// 根元素
		this.rootSelector = options.inWrapper ? `.${this.className}-wrapper` : ':root';
		// 文档CSS样式
		this.styleMap = null;
		// 主体容器
		this.wrapper = bodyContainer;
		// styleContainer== null，styleContainer = bodyContainer
		styleContainer = styleContainer || bodyContainer;

		// CSS样式生成容器，清空所有CSS样式
		removeAllElements(styleContainer);
		// HTML生成容器，清空所有HTML元素
		removeAllElements(bodyContainer);

		// 添加注释
		appendComment(styleContainer, 'docxjs library predefined styles');
		// 添加默认CSS样式
		styleContainer.appendChild(this.renderDefaultStyle());

		// 主题CSS样式
		if (document.themePart) {
			appendComment(styleContainer, 'docxjs document theme values');
			this.renderTheme(document.themePart, styleContainer);
		}
		// 文档默认CSS样式，包含表格、列表、段落、字体，样式存在继承顺序
		if (document.stylesPart != null) {
			this.styleMap = this.processStyles(document.stylesPart.styles);

			appendComment(styleContainer, 'docxjs document styles');
			styleContainer.appendChild(this.renderStyles(document.stylesPart.styles));
		}
		// 多级列表样式
		if (document.numberingPart) {
			this.processNumberings(document.numberingPart.domNumberings);

			appendComment(styleContainer, "docxjs document numbering styles");
			styleContainer.appendChild(this.renderNumbering(document.numberingPart.domNumberings, styleContainer));
			//styleContainer.appendChild(this.renderNumbering2(document.numberingPart, styleContainer));
		}
		// 字体列表CSS样式
		if (!options.ignoreFonts && document.fontTablePart) {
			this.renderFontTable(document.fontTablePart, styleContainer);
		}
		// 生成脚注部分的Map
		if (document.footnotesPart) {
			this.footnoteMap = keyBy(document.footnotesPart.notes, x => x.id);
		}
		// 生成尾注部分的Map
		if (document.endnotesPart) {
			this.endnoteMap = keyBy(document.endnotesPart.notes, x => x.id);
		}
		// 文档设置
		if (document.settingsPart) {
			this.defaultTabSize = document.settingsPart.settings?.defaultTabStop;
		}
		// 根据option生成wrapper
		if (this.options.inWrapper) {
			this.wrapper = this.renderWrapper();
			bodyContainer.appendChild(this.wrapper);
		} else {
			this.wrapper = bodyContainer;
		}
		// 生成Canvas画布元素--Konva框架
		this.renderKonva();
		// 主文档--内容
		await this.renderPages(document.documentPart.body);

		// 刷新制表符
		this.refreshTabStops();
	}

	// 渲染默认样式
	renderDefaultStyle() {
		const c = this.className;
		const styleText = `
			.${c}-wrapper { background: gray; padding: 30px; padding-bottom: 0px; display: flex; flex-flow: column; align-items: center; line-height:normal; font-weight:normal; } 
			.${c}-wrapper>section.${c} { background: white; box-shadow: 0 0 10px rgba(0, 0, 0, 0.5); margin-bottom: 30px; }
			.${c} { color: black; hyphens: auto; text-underline-position: from-font; }
			section.${c} { box-sizing: border-box; display: flex; flex-flow: column nowrap; position: relative; overflow: hidden; }
            section.${c}>header { position: absolute; top: 0; z-index: 1; display: flex; align-items: flex-end; }
			section.${c}>article { z-index: 1; }
			section.${c}>footer { position: absolute; bottom: 0; z-index: 1; }
			.${c} table { border-collapse: collapse; }
			.${c} table td, .${c} table th { vertical-align: top; }
			.${c} p { margin: 0pt; min-height: 1em; }
			.${c} span { white-space: pre-wrap; overflow-wrap: break-word; }
			.${c} a { color: inherit; text-decoration: inherit; }
			.${c} img, ${c} svg { vertical-align: baseline; }
			.${c} .clearfix::after { content: ""; display: block; line-height: 0; clear: both; }
		`;

		return createStyleElement(styleText);
	}

	// 文档CSS主题样式
	renderTheme(themePart: ThemePart, styleContainer: HTMLElement) {
		const variables = {};
		const fontScheme = themePart.theme?.fontScheme;

		if (fontScheme) {
			if (fontScheme.majorFont) {
				variables['--docx-majorHAnsi-font'] = fontScheme.majorFont.latinTypeface;
			}

			if (fontScheme.minorFont) {
				variables['--docx-minorHAnsi-font'] = fontScheme.minorFont.latinTypeface;
			}
		}

		const colorScheme = themePart.theme?.colorScheme;

		if (colorScheme) {
			for (const [k, v] of Object.entries(colorScheme.colors)) {
				variables[`--docx-${k}-color`] = `#${v}`;
			}
		}

		const cssText = this.styleToString(`.${this.className}`, variables);
		styleContainer.appendChild(createStyleElement(cssText));
	}

	// 计算className，小写，默认前缀："docx_"
	processStyleName(className: string): string {
		return className ? `${this.className}_${escapeClassName(className)}` : this.className;
	}

	// 处理样式继承
	processStyles(styles: IDomStyle[]) {
		//
		let stylesMap = keyBy(styles.filter(x => x.id != null), x => x.id);
		// 遍历base_on关系,合并样式
		for (const style of styles.filter(x => x.basedOn)) {
			const baseStyle = stylesMap[style.basedOn];

			if (baseStyle) {
				// 深度合并
				style.paragraphProps = mergeDeep(style.paragraphProps, baseStyle.paragraphProps);
				style.runProps = mergeDeep(style.runProps, baseStyle.runProps);

				for (let baseValues of baseStyle.styles) {
					let styleValues = style.styles.find(x => x.target == baseValues.target);

					if (styleValues) {
						this.copyStyleProperties(baseValues.values, styleValues.values);
					} else {
						style.styles.push({ ...baseValues, values: { ...baseValues.values } });
					}
				}
			} else if (this.options.debug) {
				console.warn(`Can't find base style ${style.basedOn}`);
			}
		}

		for (const style of styles) {
			style.cssName = this.processStyleName(style.id);
		}

		return stylesMap;
	}

	renderStyles(styles: IDomStyle[]): HTMLElement {
		let styleText = "";
		let stylesMap = this.styleMap;
		let defaultStyles = keyBy(styles.filter(s => s.isDefault), s => s.target);

		for (const style of styles) {
			let subStyles = style.styles;

			if (style.linked) {
				const linkedStyle = style.linked && stylesMap[style.linked];

				if (linkedStyle) {
					subStyles = subStyles.concat(linkedStyle.styles);
				} else if (this.options.debug) {
					console.warn(`Can't find linked style ${style.linked}`);
				}
			}

			for (const subStyle of subStyles) {
				//TODO temporary disable modificators until test it well
				let selector = `${style.target ?? ''}.${style.cssName}`; //${subStyle.mod ?? ''}

				if (style.target != subStyle.target) {
					selector += ` ${subStyle.target}`;
				}

				if (defaultStyles[style.target] == style) {
					selector = `.${this.className} ${style.target}, ` + selector;
				}

				styleText += this.styleToString(selector, subStyle.values);
			}
		}

		return createStyleElement(styleText);
	}

	processNumberings(numberings: IDomNumbering[]) {
		for (const num of numberings.filter(n => n.pStyleName)) {
			const style = this.findStyle(num.pStyleName);

			if (style?.paragraphProps?.numbering) {
				style.paragraphProps.numbering.level = num.level;
			}
		}
	}

	renderNumbering(numberings: IDomNumbering[], styleContainer: HTMLElement) {
		let styleText = '';
		const resetCounters = [];

		for (const num of numberings) {
			const selector = `p.${this.numberingClass(num.id, num.level)}`;
			let listStyleType = 'none';

			if (num.bullet) {
				const valiable = `--${this.className}-${num.bullet.src}`.toLowerCase();

				styleText += this.styleToString(`${selector}:before`, {
					"content": "' '",
					"display": "inline-block",
					"background": `var(${valiable})`
				}, num.bullet.style);

				this.document.loadNumberingImage(num.bullet.src).then(data => {
					const text = `${this.rootSelector} { ${valiable}: url(${data}) }`;
					styleContainer.appendChild(createStyleElement(text));
				});
			} else if (num.levelText) {
				const counter = this.numberingCounter(num.id, num.level);
				const counterReset = counter + ' ' + (num.start - 1);
				if (num.level > 0) {
					styleText += this.styleToString(`p.${this.numberingClass(num.id, num.level - 1)}`, {
						"counter-reset": counterReset
					});
				}
				// reset all level counters with start value
				resetCounters.push(counterReset);

				styleText += this.styleToString(`${selector}:before`, {
					"content": this.levelTextToContent(num.levelText, num.suff, num.id, this.numFormatToCssValue(num.format)),
					"counter-increment": counter,
					...num.rStyle,
				});
			} else {
				listStyleType = this.numFormatToCssValue(num.format);
			}

			styleText += this.styleToString(selector, {
				display: 'list-item',
				'list-style-position': 'inside',
				'list-style-type': listStyleType,
				...num.pStyle,
			});
		}

		if (resetCounters.length > 0) {
			styleText += this.styleToString(this.rootSelector, {
				'counter-reset': resetCounters.join(' '),
			});
		}

		return createStyleElement(styleText);
	}

	numberingClass(id: string, lvl: number) {
		return `${this.className}-num-${id}-${lvl}`;
	}

	styleToString(selectors: string, values: Record<string, string>, cssText: string = null) {
		let result = `${selectors} {\r\n`;

		for (const key in values) {
			if (key.startsWith('$')) {
				continue;
			}

			result += `  ${key}: ${values[key]};\r\n`;
		}

		if (cssText) {
			result += cssText;
		}

		return result + '}\r\n';
	}

	numberingCounter(id: string, lvl: number) {
		return `${this.className}-num-${id}-${lvl}`;
	}

	levelTextToContent(text: string, suff: string, id: string, numformat: string) {
		const suffMap = {
			tab: '\\9',
			space: '\\a0',
		};

		const result = text.replace(/%\d*/g, s => {
			const lvl = parseInt(s.substring(1), 10) - 1;
			return `"counter(${this.numberingCounter(id, lvl)}, ${numformat})"`;
		});

		return `"${result}${suffMap[suff] ?? ''}"`;
	}

	numFormatToCssValue(format: string) {
		const mapping = {
			none: 'none',
			bullet: 'disc',
			decimal: 'decimal',
			lowerLetter: 'lower-alpha',
			upperLetter: 'upper-alpha',
			lowerRoman: 'lower-roman',
			upperRoman: 'upper-roman',
			decimalZero: 'decimal-leading-zero', // 01,02,03,...
			// ordinal: "", // 1st, 2nd, 3rd,...
			// ordinalText: "", //First, Second, Third, ...
			// cardinalText: "", //One,Two Three,...
			// numberInDash: "", //-1-,-2-,-3-, ...
			// hex: "upper-hexadecimal",
			aiueo: 'katakana',
			aiueoFullWidth: 'katakana',
			chineseCounting: 'simp-chinese-informal',
			chineseCountingThousand: 'simp-chinese-informal',
			chineseLegalSimplified: 'simp-chinese-formal', // 中文大写
			chosung: 'hangul-consonant',
			ideographDigital: 'cjk-ideographic',
			ideographTraditional: 'cjk-heavenly-stem', // 十天干
			ideographLegalTraditional: 'trad-chinese-formal',
			ideographZodiac: 'cjk-earthly-branch', // 十二地支
			iroha: 'katakana-iroha',
			irohaFullWidth: 'katakana-iroha',
			japaneseCounting: 'japanese-informal',
			japaneseDigitalTenThousand: 'cjk-decimal',
			japaneseLegal: 'japanese-formal',
			thaiNumbers: 'thai',
			koreanCounting: 'korean-hangul-formal',
			koreanDigital: 'korean-hangul-formal',
			koreanDigital2: 'korean-hanja-informal',
			hebrew1: 'hebrew',
			hebrew2: 'hebrew',
			hindiNumbers: 'devanagari',
			ganada: 'hangul',
			taiwaneseCounting: 'cjk-ideographic',
			taiwaneseCountingThousand: 'cjk-ideographic',
			taiwaneseDigital: 'cjk-decimal',
		};

		return mapping[format] ?? format;
	}

	// renderNumbering2(numberingPart: NumberingPartProperties, container: HTMLElement): HTMLElement {
	//     let css = "";
	//     let numberingMap = keyBy(numberingPart.abstractNumberings, x => x.id);
	//     let bulletMap = keyBy(numberingPart.bulletPictures, x => x.id);
	//     let topCounters = [];
	//
	//     for(let num of numberingPart.numberings) {
	//         let absNum = numberingMap[num.abstractId];
	//
	//         for(let lvl of absNum.levels) {
	//             let className = this.numberingClass(num.id, lvl.level);
	//             let listStyleType = "none";
	//
	//             if(lvl.text && lvl.format == 'decimal') {
	//                 let counter = this.numberingCounter(num.id, lvl.level);
	//
	//                 if (lvl.level > 0) {
	//                     css += this.styleToString(`p.${this.numberingClass(num.id, lvl.level - 1)}`, {
	//                         "counter-reset": counter
	//                     });
	//                 } else {
	//                     topCounters.push(counter);
	//                 }
	//
	//                 css += this.styleToString(`p.${className}:before`, {
	//                     "content": this.levelTextToContent(lvl.text, num.id),
	//                     "counter-increment": counter
	//                 });
	//             } else if(lvl.bulletPictureId) {
	//                 let pict = bulletMap[lvl.bulletPictureId];
	//                 let variable = `--${this.className}-${pict.referenceId}`.toLowerCase();
	//
	//                 css += this.styleToString(`p.${className}:before`, {
	//                     "content": "' '",
	//                     "display": "inline-block",
	//                     "background": `var(${variable})`
	//                 }, pict.style);
	//
	//                 this.document.loadNumberingImage(pict.referenceId).then(data => {
	//                     var text = `.${this.className}-wrapper { ${variable}: url(${data}) }`;
	//                     container.appendChild(createStyleElement(text));
	//                 });
	//             } else {
	//                 listStyleType = this.numFormatToCssValue(lvl.format);
	//             }
	//
	//             css += this.styleToString(`p.${className}`, {
	//                 "display": "list-item",
	//                 "list-style-position": "inside",
	//                 "list-style-type": listStyleType,
	//                 //TODO
	//                 //...num.style
	//             });
	//         }
	//     }
	//
	//     if (topCounters.length > 0) {
	//         css += this.styleToString(`.${this.className}-wrapper`, {
	//             "counter-reset": topCounters.join(" ")
	//         });
	//     }
	//
	//     return createStyleElement(css);
	// }

	// 字体列表CSS样式
	renderFontTable(fontsPart: FontTablePart, styleContainer: HTMLElement) {
		for (const f of fontsPart.fonts) {
			for (const ref of f.embedFontRefs) {
				this.document.loadFont(ref.id, ref.key).then(fontData => {
					const cssValues = {
						'font-family': f.name,
						src: `url(${fontData})`,
					};

					if (ref.type == 'bold' || ref.type == 'boldItalic') {
						cssValues['font-weight'] = 'bold';
					}

					if (ref.type == 'italic' || ref.type == 'boldItalic') {
						cssValues['font-style'] = 'italic';
					}

					appendComment(styleContainer, `docxjs ${f.name} font`);
					const cssText = this.styleToString('@font-face', cssValues);
					styleContainer.appendChild(createStyleElement(cssText));
					this.refreshTabStops();
				});
			}
		}
	}

	// 生成父级容器
	renderWrapper() {
		return createElement('div', { className: `${this.className}-wrapper` });
	}

	// 复制CSS样式
	copyStyleProperties(input: Record<string, string>, output: Record<string, string>, attrs: string[] = null): Record<string, string> {
		if (!input) {
			return output;
		}
		if (output == null) {
			output = {};
		}
		if (attrs == null) {
			attrs = Object.getOwnPropertyNames(input);
		}

		for (const key of attrs) {
			if (input.hasOwnProperty(key) && !output.hasOwnProperty(key))
				output[key] = input[key];
		}

		return output;
	}

	// 递归明确元素parent父级关系
	processElement(element: OpenXmlElement) {
		if (element.children) {
			for (const e of element.children) {
				e.parent = element;
				// 判断类型
				if (e.type == DomType.Table) {
					// 渲染表格
					this.processTable(e);
					this.processElement(e);
				} else {
					// 递归渲染
					this.processElement(e);
				}
			}
		}
	}

	// 处理表格style样式
	processTable(table: WmlTable) {
		for (const r of table.children) {
			for (const c of r.children) {
				c.cssStyle = this.copyStyleProperties(table.cellStyle, c.cssStyle, [
					'border-left',
					'border-right',
					'border-top',
					'border-bottom',
					'padding-left',
					'padding-right',
					'padding-top',
					'padding-bottom',
				]);
			}
		}
	}

	/*
	 * section与page概念区别
	 * 章节(section)是根据内容的逻辑结构和组织来划分的，不同章节设置独立的格式。
	 * 页面是文档实际呈现的物理单位，而章节则是逻辑上的分割点。
	 */

	// 根据分页符拆分页面
	splitPage(elements: OpenXmlElement[]): Page[] {
		// 当前操作page，elements数组包含子元素
		let current_page: Page = new Page({} as PageProps);
		// 切分出的所有pages
		const pages: Page[] = [current_page];

		for (const elem of elements) {
			// 标记顶层元素的层级level
			elem.level = 1;
			// 添加elem进入当前操作page
			current_page.elements.push(elem);
			/* 段落基本结构：paragraph => run => text... */
			if (elem.type == DomType.Paragraph) {
				const p = elem as WmlParagraph;
				// 节属性，代表分节符
				const sectProps: SectionProperties = p.sectionProps;
				// 节属性生成唯一uuid，每一个节中page均是同一个uuid，代表属于同一个节
				if (sectProps) {
					sectProps.sectionId = uuid();
				}
				// 查找内置默认段落样式
				const default_paragraph_style = this.findStyle(p.styleName);

				// 检测段落内置样式是否存在段前分页符
				if (default_paragraph_style?.paragraphProps?.pageBreakBefore) {
					// 标记当前page已拆分
					current_page.isSplit = true;
					// 保存当前page的sectionProps
					current_page.sectProps = sectProps;
					// 重置新的page
					current_page = new Page({} as PageProps);
					// 添加新page
					pages.push(current_page);
				}

				// 段落部分Break索引
				let pBreakIndex = -1;
				// Run部分Break索引
				let rBreakIndex = -1;

				// 查询段落中Break索引
				if (p.children) {
					// 计算段落Break索引
					pBreakIndex = p.children.findIndex(r => {
						// 计算Run Break索引
						rBreakIndex = r.children?.findIndex((t: OpenXmlElement) => {
							// 如果不是分页符、换行符、分栏符
							if (t.type != DomType.Break) {
								return false;
							}
							// 默认忽略lastRenderedPageBreak，
							if ((t as WmlBreak).break == 'lastRenderedPageBreak') {
								// 判断前一个p段落，
								// 如果含有分页符、分节符，那它们一定位于上一个page，数组为空；
								// 如果前一个段落是普通段落，数组长度大于0，则代表文字过多超过一页，需要自动分页
								return (current_page.elements.length > 2 || !this.options.ignoreLastRenderedPageBreak);
							}
							// 分页符
							if ((t as WmlBreak).break === 'page') {
								return true;
							}
						});
						rBreakIndex = rBreakIndex ?? -1;
						return rBreakIndex != -1;
					});
				}
				// 段落Break索引
				if (pBreakIndex != -1) {
					// 一般情况下，标记当前page：已拆分
					current_page.isSplit = true;
					// 检测分页符之前的所有元素是否存在表格
					const exist_table: boolean = current_page.elements.some(
						elem => elem.type === DomType.Table
					);
					// 存在表格
					if (exist_table) {
						// 表格可能需要计算之后拆分，标记当前page：未拆分
						current_page.isSplit = false;
					}
					// 检测分页符之前的所有元素是否存在目录
					let exist_TOC: boolean = current_page.elements.some((paragraph) => {
						return paragraph.children.some((elem) => {
							if (elem.type === DomType.Hyperlink) {
								return (elem as WmlHyperlink)?.href?.includes('Toc')
							}
							return false;
						});
					});
					// 	存在目录
					if (exist_TOC) {
						// 目录可能需要计算之后拆分，标记当前page：未拆分
						current_page.isSplit = false;
					}
				}
				/*
				 *
				 * 分页有两种情况：
				 * 1、段落中存在节属性sectProps，且类型不是continuous/nextColumn
				 * 2、段落存在Break索引
				 *
				 */
				if (pBreakIndex != -1 || (sectProps && sectProps.type != SectionType.Continuous && sectProps.type != SectionType.NextColumn)) {
					// 保存当前page的pageProps
					current_page.sectProps = sectProps;
					// 重置新的page
					current_page = new Page({} as PageProps);
					// 添加新page
					pages.push(current_page);
				}
				// 根据段落Break索引，拆分Run部分
				if (pBreakIndex != -1) {
					// 即将拆分的Run部分
					const breakRun = p.children[pBreakIndex];
					// 是否需要拆分Run
					const is_split = rBreakIndex < breakRun.children.length - 1;

					if (pBreakIndex < p.children.length - 1 || is_split) {
						// 原始的Run数组
						const origin_runs = p.children;
						// 切出Break索引后面的Run，创建新段落
						const new_paragraph: WmlParagraph = {
							...p,
							children: origin_runs.slice(pBreakIndex),
						};
						// 保存Break索引前面的Run
						p.children = origin_runs.slice(0, pBreakIndex);
						// 添加新段落
						current_page.elements.push(new_paragraph);

						if (is_split) {
							// Run下面原始的元素
							const origin_elements = breakRun.children;
							// 切出Run Break索引前面的元素，创建新Run
							const newRun = {
								...breakRun,
								children: origin_elements.slice(0, rBreakIndex),
							};
							// 将新Run放入上一个page的段落
							p.children.push(newRun);
							// 切出Run Break索引后面的元素
							breakRun.children = origin_elements.slice(rBreakIndex);
						}
					}
				}
			}

			// elem元素是表格，需要渲染过程中拆分page
			if (elem.type === DomType.Table) {
				// 标记当前page：未拆分
				current_page.isSplit = false;
			}
		}
		// 一个节可能分好几个页，但是节属性sectionProps存在当前节中最后一段对应的 paragraph 元素的子元素。即：[null,null,null,setPr];
		let currentSectProps = null;
		// 倒序给每一页填充sectionProps，方便后期页面渲染
		for (let i = pages.length - 1; i >= 0; i--) {
			if (pages[i].sectProps == null) {
				pages[i].sectProps = currentSectProps;
			} else {
				currentSectProps = pages[i].sectProps;
			}
		}
		return pages;
	}

	// 生成所有的页面Page
	async renderPages(document: DocumentElement) {
		// 生成页面parent父级关系
		this.processElement(document);
		// 根据options.breakPages，选择是否分页
		let pages: Page[];
		if (this.options.breakPages) {
			// 根据分页符，初步拆分页面
			pages = this.splitPage(document.children);
		} else {
			// 不分页则，只有一个page
			pages = [new Page({ sectProps: document.props, elements: document.children, } as PageProps)];
		}
		// 初步分页结果,缓存至body中
		document.pages = pages;
		// 前一个节属性，判断分节符的第一个page
		let prevProps = null;
		// 深拷贝初步分页结果，后续拆分操作将不断扩充数组，导致下面循环异常
		let origin_pages = structuredClone(pages);
		// 遍历生成每一个page
		for (let i = 0; i < origin_pages.length; i++) {
			this.currentFootnoteIds = [];
			const page: Page = origin_pages[i];
			const { sectProps } = page;
			// sectionProps属性不存在，则使用文档级别props;
			page.sectProps = sectProps ?? document.props;
			// 是否本小节的第一个page
			page.isFirstPage = prevProps != page.sectProps;
			// TODO 是否最后一个page,此时分页未完成，计算并不准确，影响到尾注的渲染
			page.isLastPage = i === origin_pages.length - 1;
			// 溢出检测默认不开启
			page.checkingOverflow = false;
			// 将上述数据存储在currentPage中
			this.currentPage = page;
			// 存储前一个节属性
			prevProps = page.sectProps;
			// 渲染单个page
			await this.renderPage();
		}

	}

	// 生成单个page,如果发现超出一页，递归拆分出下一个page
	async renderPage() {
		// 解构当前操作的page中的属性
		const { pageId, sectProps, elements, isFirstPage, isLastPage } = this.currentPage;
		// 根据sectProps，创建page
		const pageElement = this.createPage(this.className, sectProps);

		// 给page添加背景样式
		this.renderStyleValues(
			this.document.documentPart.body.cssStyle,
			pageElement
		);
		// 已拆分的Pages数组
		let pages = this.document.documentPart.body.pages;
		// 计算当前Page的索引
		let pageIndex = pages.findIndex((page) => page.pageId === pageId);
		// 渲染page页眉
		if (this.options.renderHeaders) {
			await this.renderHeaderFooterRef(
				sectProps.headerRefs,
				sectProps,
				pageIndex,
				isFirstPage,
				pageElement
			);
		}
		// 渲染page脚注
		if (this.options.renderFootnotes) {
			await this.renderNotes(
				this.currentFootnoteIds,
				this.footnoteMap,
				pageElement
			);
		}
		// 渲染page尾注，判断最后一页
		if (this.options.renderEndnotes && isLastPage) {
			await this.renderNotes(
				this.currentEndnoteIds,
				this.endnoteMap,
				pageElement
			);
		}
		// 渲染page页脚
		if (this.options.renderFooters) {
			await this.renderHeaderFooterRef(
				sectProps.footerRefs,
				sectProps,
				pageIndex,
				isFirstPage,
				pageElement
			);
		}

		// TODO 分栏情况下，有可能一个page一种分栏，在分节符（continuous）情况下，一个page拥有多种分栏；

		// page内容区---Article元素
		const contentElement = this.createPageContent(sectProps);
		// 根据options.breakPages，设置article的高度
		if (this.options.breakPages) {
			// 切分页面，高度固定
			contentElement.style.height = sectProps.contentSize.height;
		} else {
			// 不分页则，拥有最小高度
			contentElement.style.minHeight = sectProps.contentSize.height;
		}
		// 缓存当前操作的Article元素
		this.currentPage.contentElement = contentElement;
		// 将Article插入page
		pageElement.appendChild(contentElement);
		// 标识--开启溢出计算
		this.currentPage.checkingOverflow = true;
		// 生成article内容
		let is_overflow = await this.renderElements(elements, contentElement);
		// 元素没有溢出Page
		if (is_overflow === Overflow.FALSE) {
			// 修改当前Page的状态
			this.currentPage.isSplit = true;
			// 替换当前page
			pages[pageIndex] = this.currentPage;
		}
		// 标识--结束溢出计算
		this.currentPage.checkingOverflow = false;
	}

	// 创建Page
	createPage(className: string, props: SectionProperties) {
		const oPage = createElement('section', { className });

		if (props) {
			// 生成uuid标识，相同的uuid即属于同一个节
			oPage.dataset.sectionId = props.sectionId;
			// 页边距
			if (props.pageMargins) {
				oPage.style.paddingLeft = props.pageMargins.left;
				oPage.style.paddingRight = props.pageMargins.right;
				oPage.style.paddingTop = props.pageMargins.top;
				oPage.style.paddingBottom = props.pageMargins.bottom;
			}
			// 页面尺寸
			if (props.pageSize) {
				if (!this.options.ignoreWidth) {
					oPage.style.width = props.pageSize.width;
				}
				if (!this.options.ignoreHeight) {
					oPage.style.minHeight = props.pageSize.height;
				}
			}
		}
		// 插入生成的page
		this.wrapper.appendChild(oPage);

		return oPage;
	}

	// TODO 一个页面可能存在多个章节section，每个section拥有不同的分栏
	// 多列分栏布局
	createPageContent(props: SectionProperties): HTMLElement {
		// 指代页面page，HTML5缺少page，以article代替
		const oArticle = createElement('article');
		const { count, space, separator } = props?.columns;
		// 设置多列样式
		if (count > 1) {
			oArticle.style.columnCount = `${count}`;
			oArticle.style.columnGap = space;
		}
		if (separator) {
			oArticle.style.columnRule = '1px solid black';
		}
		return oArticle;
	}

	// TODO 分页不准确，页脚页码混乱
	// 渲染页眉/页脚的Ref
	async renderHeaderFooterRef(refs: FooterHeaderReference[], props: SectionProperties, pageIndex: number, isFirstPage: boolean, parent: HTMLElement) {
		if (!refs) return;
		// 根据首页、奇数、偶数类型，查找ref指向
		let ref: FooterHeaderReference;
		if (props.titlePage && isFirstPage) {
			// 第一页
			ref = refs.find(x => x.type == "first");
		} else if (pageIndex % 2 == 1) {
			// 注意，pageIndex从0开始，却代表第一页，此处判断条件确实对应偶数页
			ref = refs.find(x => x.type == "even");
		} else {
			// 奇数页
			ref = refs.find(x => x.type == "default");
		}
		// 查找ref对应的part部分
		let part = this.document.findPartByRelId(ref?.id, this.document.documentPart) as BaseHeaderFooterPart;

		if (part) {
			this.currentPart = part;
			if (!this.usedHederFooterParts.includes(part.path)) {
				this.processElement(part.rootElement);
				this.usedHederFooterParts.push(part.path);
			}
			// 根据页眉页脚，设置CSS
			switch (part.rootElement.type) {
				case DomType.Header:
					part.rootElement.cssStyle = {
						left: props.pageMargins?.left,
						width: props.contentSize?.width,
						height: props.pageMargins?.top,
					};
					break;
				case DomType.Footer:
					part.rootElement.cssStyle = {
						left: props.pageMargins?.left,
						width: props.contentSize?.width,
						height: props.pageMargins?.bottom,
					};
					break;
				default:
					console.warn('set header/footer style error', part.rootElement.type);
					break;
			}

			await this.renderElements([part.rootElement], parent);
			this.currentPart = null;
		}
	}

	// 渲染脚注/尾注
	async renderNotes(noteIds: string[], notesMap: Record<string, WmlBaseNote>, parent: HTMLElement) {
		const notes = noteIds.map(id => notesMap[id]).filter(x => x);

		if (notes.length > 0) {
			const oList = createElement('ol', null);
			await this.renderElements(notes, oList);
			parent.appendChild(oList);
		}
	}

	// 根据XML对象渲染出多元素
	async renderElements(elems: OpenXmlElement[], parent: HTMLElement | Element) {
		let is_overflow: Overflow = Overflow.FALSE;
		for (let i = 0; i < elems.length; i++) {
			// 顶层元素
			if (elems[i].level === 1) {
				// currentPage缓存当前操作顶层元素的索引值,将会不断覆盖，直到此顶层元素溢出，elementIndex即是溢出元素的elements索引。
				this.currentPage.breakIndex = i;
			}
			// 根据XML对象渲染单个元素
			const element = (await this.renderElement(elems[i], parent)) as Node_DOM;
			// 如果溢出文档，跳出循环
			if (element?.dataset?.overflow === Overflow.TRUE) {
				is_overflow = Overflow.TRUE;
				break;
			}
		}
		return is_overflow;
	}

	// 根据XML对象渲染单个元素
	async renderElement(elem: OpenXmlElement, parent?: HTMLElement | Element): Promise<Node_DOM> {
		let oNode;

		switch (elem.type) {
			case DomType.Paragraph:
				oNode = await this.renderParagraph(elem as WmlParagraph, parent);
				break;

			case DomType.Run:
				oNode = await this.renderRun(elem as WmlRun, parent);
				break;

			case DomType.Text:
				oNode = await this.renderText(elem as WmlText, parent);
				break;

			case DomType.Table:
				oNode = await this.renderTable(elem as WmlTable, parent);
				break;

			case DomType.Row:
				oNode = await this.renderTableRow(elem, parent);
				break;

			case DomType.Cell:
				oNode = await this.renderTableCell(elem);
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.Hyperlink:
				oNode = await this.renderHyperlink(elem, parent);
				break;

			case DomType.Drawing:
				oNode = await this.renderDrawing(elem as WmlDrawing, parent);
				break;

			case DomType.Image:
				oNode = await this.renderImage(elem as WmlImage, parent);
				break;

			case DomType.BookmarkStart:
				oNode = this.renderBookmarkStart(elem as WmlBookmarkStart);
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.BookmarkEnd:
				//ignore bookmark end
				oNode = null;
				break;

			case DomType.Tab:
				oNode = await this.renderTab(elem, parent);
				break;

			case DomType.Symbol:
				oNode = await this.renderSymbol(elem as WmlSymbol, parent);
				break;

			case DomType.Break:
				oNode = await this.renderBreak(elem as WmlBreak, parent);
				break;

			case DomType.Inserted:
				oNode = await this.renderInserted(elem);
				if (parent) {
					await this.appendChildren(parent, oNode);
				}
				break;

			case DomType.Deleted:
				oNode = await this.renderDeleted(elem);
				if (parent) {
					await this.appendChildren(parent, oNode);
				}
				break;

			case DomType.DeletedText:
				oNode = await this.renderDeletedText(elem as WmlText, parent);
				break;

			case DomType.NoBreakHyphen:
				oNode = createElement('wbr');
				if (parent) {
					await this.appendChildren(parent, oNode);
				}
				break;

			case DomType.CommentRangeStart:
				oNode = this.renderCommentRangeStart(elem);
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.CommentRangeEnd:
				oNode = this.renderCommentRangeEnd(elem);
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.CommentReference:
				oNode = this.renderCommentReference(elem);
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.Footer:
				oNode = await this.renderHeaderFooter(elem, 'footer');
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.Header:
				oNode = await this.renderHeaderFooter(elem, 'header');
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.Footnote:
			case DomType.Endnote:
				oNode = await this.renderContainer(elem, 'li');
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.FootnoteReference:
				oNode = this.renderFootnoteReference(elem as WmlNoteReference);
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.EndnoteReference:
				oNode = this.renderEndnoteReference(elem as WmlNoteReference);
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.VmlElement:
				oNode = await this.renderVmlElement(elem as VmlElement, parent);
				break;

			case DomType.VmlPicture:
				oNode = await this.renderVmlPicture(elem);
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlMath:
				oNode = await this.renderContainerNS(elem, ns.mathML, 'math', {
					xmlns: ns.mathML,
				});
				// 作为子元素插入,针对此元素进行溢出检测
				if (parent) {
					oNode.dataset.overflow = await this.appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlMathParagraph:
				oNode = await this.renderContainer(elem, 'span');
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlFraction:
				oNode = await this.renderContainerNS(elem, ns.mathML, 'mfrac');
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlBase:
				oNode = await this.renderContainerNS(elem, ns.mathML, elem.parent.type == DomType.MmlMatrixRow ? "mtd" : "mrow");
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlNumerator:
			case DomType.MmlDenominator:
			case DomType.MmlFunction:
			case DomType.MmlLimit:
			case DomType.MmlBox:
				oNode = await this.renderContainerNS(elem, ns.mathML, 'mrow');
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlGroupChar:
				oNode = await this.renderMmlGroupChar(elem);
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlLimitLower:
				oNode = await this.renderContainerNS(elem, ns.mathML, 'munder');
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlMatrix:
				oNode = await this.renderContainerNS(elem, ns.mathML, 'mtable');
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlMatrixRow:
				oNode = await this.renderContainerNS(elem, ns.mathML, 'mtr');
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlRadical:
				oNode = await this.renderMmlRadical(elem);
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlSuperscript:
				oNode = await this.renderContainerNS(elem, ns.mathML, 'msup');
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlSubscript:
				oNode = await this.renderContainerNS(elem, ns.mathML, 'msub');
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlDegree:
			case DomType.MmlSuperArgument:
			case DomType.MmlSubArgument:
				oNode = await this.renderContainerNS(elem, ns.mathML, 'mn');
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlFunctionName:
				oNode = await this.renderContainerNS(elem, ns.mathML, 'ms');
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlDelimiter:
				oNode = await this.renderMmlDelimiter(elem);
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlRun:
				oNode = await this.renderMmlRun(elem);
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlNary:
				oNode = await this.renderMmlNary(elem);
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlPreSubSuper:
				oNode = await this.renderMmlPreSubSuper(elem);
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlBar:
				oNode = await this.renderMmlBar(elem);
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlEquationArray:
				oNode = await this.renderMllList(elem);
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;
		}
		// 标记其XML标签名
		if (oNode && oNode?.nodeType === 1) {
			oNode.dataset.tag = elem.type;
		}

		return oNode;
	}

	// 判断是否存在分页元素
	isPageBreakElement(elem: OpenXmlElement): boolean {
		// 分页符、换行符、分栏符
		if (elem.type != DomType.Break) {
			return false;
		}
		// 默认以lastRenderedPageBreak作为分页依据
		if ((elem as WmlBreak).break == 'lastRenderedPageBreak') {
			return !this.options.ignoreLastRenderedPageBreak;
		}
		// 分页符
		if ((elem as WmlBreak).break === 'page') {
			return true;
		}
	}

	// 根据XML对象渲染子元素，并插入父级元素
	async renderChildren(elem: OpenXmlElement, parent: Element) {
		return await this.renderElements(elem.children, parent);
	}

	// 插入子元素，针对后代元素进行溢出检测
	async appendChildren(parent: HTMLElement | Element, children: ChildrenType, xml_element?: OpenXmlElement): Promise<Overflow> {
		// 插入元素
		appendChildren(parent, children);
		// 是否溢出标识
		let is_overflow = false;
		let {
			pageId,
			sectProps,
			isSplit,
			contentElement,
			breakIndex,
			checkingOverflow,
			elements: origin_elements
		} = this.currentPage;
		// 当前page已拆分，忽略溢出检测
		if (isSplit) {
			return Overflow.UNKNOWN;
		}
		// 当前page未拆分，是否需要溢出检测
		if (checkingOverflow) {
			// 溢出检测
			is_overflow = checkOverflow(contentElement);
			// 溢出
			if (is_overflow) {
				// TODO 充分利用parent属性
				// 溢出元素 == row
				if (xml_element?.type === DomType.Row) {
					const table: OpenXmlElement = origin_elements[breakIndex];
					// 溢出元素所在tr的索引;
					const row_index = table.children.findIndex(
						elem => elem === xml_element
					);
					// 查找表格中的table header，可能有多行
					const table_headers = table.children.filter((row: WmlTableRow) => row.isHeader);
					// 删除table前面已经渲染的row，保留后续未渲染元素
					table.children.splice(0, row_index);
					// 填充table header
					table.children = [...table_headers, ...table.children];
				}
				// TODO 由于初期针对段落的分页已经完成，此处拆分针对目录部分。按照顶层元素拆分，分页并不精确。
				// 根据顶层元素的索引breakIndex，保留后续元素，以备下一个Page渲染，保留原始数组前面已经渲染的元素
				let elements = origin_elements.splice(breakIndex);
				// 删除DOM中导致溢出的元素
				removeElements(children, parent);
				// 已拆分的Pages数组
				let pages = this.document.documentPart.body.pages;
				// 计算当前Page的索引
				let pageIndex = pages.findIndex((page) => page.pageId === pageId);
				// 修改当前Page的状态
				this.currentPage.isSplit = true;
				this.currentPage.checkingOverflow = false;
				// 生成新的page，新Page的sectionProps沿用前一页的sectionProps
				const page: Page = new Page({ sectProps, elements } as PageProps);
				// 替换当前page
				pages[pageIndex] = this.currentPage;
				// 缓存新拆分出去的page
				pages.splice(pageIndex + 1, 0, page);
				// 覆盖current_page的属性
				this.currentPage = page;
				// 重启新一个page的渲染
				await this.renderPage();
			}
		}

		return is_overflow ? Overflow.TRUE : Overflow.FALSE;
	}

	async renderContainer(elem: OpenXmlElement, tagName: keyof HTMLElementTagNameMap, props?: Record<string, any>) {
		const parent = createElement(tagName, props);
		await this.renderChildren(elem, parent);
		return parent;
	}

	async renderContainerNS(elem: OpenXmlElement, ns: string, tagName: string, props?: Record<string, any>) {
		const parent = createElementNS(ns, tagName as any, props);
		await this.renderChildren(elem, parent);
		return parent;
	}

	async renderParagraph(elem: WmlParagraph, parent: HTMLElement | Element) {
		// 创建段落元素
		const oParagraph = createElement('p');
		// 生成段落的uuid标识，
		oParagraph.dataset.uuid = uuid();
		// 渲染class
		this.renderClass(elem, oParagraph);
		// 渲染style
		this.renderStyleValues(elem.cssStyle, oParagraph);
		// 渲染常规--字体、颜色
		this.renderCommonProperties(oParagraph.style, elem);
		// 查找段落内置样式class
		const style = this.findStyle(elem.styleName);
		elem.tabs ??= style?.paragraphProps?.tabs; //TODO
		// 列表序号
		const numbering = elem.numbering ?? style?.paragraphProps?.numbering;

		if (numbering) {
			oParagraph.classList.add(
				this.numberingClass(numbering.id, numbering.level)
			);
		}

		// TODO 子代元素（Run）=> 孙代元素（Drawing）,可能有n个drawML对象。目前仅考虑一个DrawML的情况，多个DrawML对象定位存在bug
		// 是否需要清除浮动
		const is_clear = elem.children.some(run => {
			// 是否存在上下型环绕
			const is_exist_drawML = run?.children?.some(
				child => child.type === DomType.Drawing && child.props.wrapType === WrapType.TopAndBottom
			);
			// 是否存在br元素拥有clear属性
			const is_clear_break = run?.children?.some(
				child => child.type === DomType.Break && child?.props?.clear
			);
			return is_exist_drawML || is_clear_break;
		});
		// 仅在上下型环绕清除浮动
		if (is_clear) {
			oParagraph.classList.add('clearfix');
		}
		// 后代元素定位参照物
		oParagraph.style.position = 'relative';

		// 如果拥有父级
		if (parent) {
			// 作为子元素插入,针对此元素进行溢出检测
			const is_overflow: Overflow = await this.appendChildren(parent, oParagraph);
			if (is_overflow === Overflow.TRUE) {
				oParagraph.dataset.overflow = Overflow.TRUE;
				return oParagraph;
			}
		}

		// 针对后代子元素进行溢出检测
		oParagraph.dataset.overflow = await this.renderChildren(elem, oParagraph);

		return oParagraph;
	}

	// TODO 需要处理溢出
	async renderRun(elem: WmlRun, parent?: HTMLElement | Element) {
		if (elem.fieldRun) {
			return null;
		}

		const oSpan = createElement('span');

		if (elem.id) {
			oSpan.id = elem.id;
		}

		this.renderClass(elem, oSpan);
		this.renderStyleValues(elem.cssStyle, oSpan);

		// 作为子元素插入，先执行溢出检测，方便对后代元素进行溢出检测
		if (parent) {
			const is_overflow: Overflow = await this.appendChildren(parent, oSpan);
			if (is_overflow === Overflow.TRUE) {
				oSpan.dataset.overflow = Overflow.TRUE;
				return oSpan;
			}
		}

		if (elem.verticalAlign) {
			const wrapper = createElement(elem.verticalAlign as any);
			oSpan.dataset.overflow = await this.renderChildren(elem, wrapper);
			oSpan.dataset.overflow = await this.appendChildren(oSpan, wrapper);
		} else {
			oSpan.dataset.overflow = await this.renderChildren(elem, oSpan);
		}

		return oSpan;
	}

	// TODO 需要处理溢出
	async renderText(elem: WmlText, parent?: HTMLElement | Element) {
		const oText = document.createTextNode(elem.text);
		// 作为子元素插入，忽略溢出检测
		if (parent) {
			appendChildren(parent, oText);
		}
		return oText;
	}

	async renderTable(elem: WmlTable, parent?: HTMLElement | Element) {
		const oTable = createElement('table');
		// 生成表格的uuid标识，
		oTable.dataset.uuid = uuid();

		this.tableCellPositions.push(this.currentCellPosition);
		this.tableVerticalMerges.push(this.currentVerticalMerge);
		this.currentVerticalMerge = {};
		this.currentCellPosition = { col: 0, row: 0 };

		this.renderClass(elem, oTable);
		this.renderStyleValues(elem.cssStyle, oTable);
		// 如果拥有父级
		if (parent) {
			// 作为子元素插入,针对此元素进行溢出检测
			const is_overflow: Overflow = await this.appendChildren(parent, oTable);
			if (is_overflow === Overflow.TRUE) {
				oTable.dataset.overflow = Overflow.TRUE;
				return oTable;
			}
		}
		// 渲染表格column列
		if (elem.columns) {
			await this.renderTableColumns(elem.columns, oTable);
		}
		// 针对后代子元素进行溢出检测
		oTable.dataset.overflow = await this.renderChildren(elem, oTable);

		this.currentVerticalMerge = this.tableVerticalMerges.pop();
		this.currentCellPosition = this.tableCellPositions.pop();

		return oTable;
	}

	// 表格--列
	async renderTableColumns(columns: WmlTableColumn[], parent?: HTMLElement | Element) {
		const oColGroup = createElement('colgroup');

		// 插入子元素,忽略溢出检测
		if (parent) {
			appendChildren(parent, oColGroup);
		}

		for (const col of columns) {
			const oCol = createElement('col');

			if (col.width) {
				oCol.style.width = col.width;
			}
			appendChildren(oColGroup, oCol);
		}

		return oColGroup;
	}

	// 表格--行
	async renderTableRow(elem: OpenXmlElement, parent?: HTMLElement | Element) {
		const oTableRow = createElement('tr');

		this.currentCellPosition.col = 0;

		this.renderClass(elem, oTableRow);
		this.renderStyleValues(elem.cssStyle, oTableRow);
		this.currentCellPosition.row++;
		// TODO 单行溢出页面，陷入死循环

		// 后代元素忽略溢出检测
		await this.renderChildren(elem, oTableRow);
		// 如果拥有父级
		if (parent) {
			// 作为子元素插入,针对此元素进行溢出检测
			oTableRow.dataset.overflow = await this.appendChildren(
				parent,
				oTableRow,
				elem
			);
		}

		return oTableRow;
	}

	// 表格--单元格
	async renderTableCell(elem: WmlTableCell) {
		const oTableCell = createElement('td');

		const key = this.currentCellPosition.col;

		if (elem.verticalMerge) {
			if (elem.verticalMerge == 'restart') {
				this.currentVerticalMerge[key] = oTableCell;
				oTableCell.rowSpan = 1;
			} else if (this.currentVerticalMerge[key]) {
				this.currentVerticalMerge[key].rowSpan += 1;
				oTableCell.style.display = 'none';
			}
		} else {
			this.currentVerticalMerge[key] = null;
		}

		this.renderClass(elem, oTableCell);
		this.renderStyleValues(elem.cssStyle, oTableCell);

		if (elem.span) {
			oTableCell.colSpan = elem.span;
		}

		this.currentCellPosition.col += oTableCell.colSpan;

		await this.renderChildren(elem, oTableCell);
		return oTableCell;
	}

	async renderHyperlink(elem: WmlHyperlink, parent?: HTMLElement | Element) {
		const oAnchor = createElement('a');

		this.renderStyleValues(elem.cssStyle, oAnchor);

		// 作为子元素插入，先执行溢出检测，方便对后代元素进行溢出检测
		if (parent) {
			const is_overflow: Overflow = await this.appendChildren(parent, oAnchor);
			if (is_overflow === Overflow.TRUE) {
				oAnchor.dataset.overflow = Overflow.TRUE;
				return oAnchor;
			}
		}

		if (elem.href) {
			oAnchor.href = elem.href;
		} else if (elem.id) {
			const rel = this.document.documentPart.rels.find(
				it => it.id == elem.id && it.targetMode === 'External'
			);
			oAnchor.href = rel?.target;
		}

		oAnchor.dataset.overflow = await this.renderChildren(elem, oAnchor);

		return oAnchor;
	}

	async renderDrawing(elem: WmlDrawing, parent?: HTMLElement | Element) {
		const oDrawing = createElement('span');

		oDrawing.style.textIndent = '0px';

		// TODO 外围添加一个元素清除浮动

		// TODO 标识当前环绕方式，后期可删除
		oDrawing.dataset.wrap = elem?.props.wrapType;

		this.renderStyleValues(elem.cssStyle, oDrawing);
		// 作为子元素插入，先执行溢出检测，方便对后代元素进行溢出检测
		if (parent) {
			const is_overflow: Overflow = await this.appendChildren(parent, oDrawing);
			if (is_overflow === Overflow.TRUE) {
				oDrawing.dataset.overflow = Overflow.TRUE;
				return oDrawing;
			}
		}
		// 对后代元素进行溢出检测
		oDrawing.dataset.overflow = await this.renderChildren(elem, oDrawing);

		return oDrawing;
	}

	// 渲染图片，默认转换blob--异步
	async renderImage(elem: WmlImage, parent?: HTMLElement | Element) {
		// 判断是否需要canvas转换
		const { is_clip, is_transform } = elem.props;
		// Image元素
		const oImage = new Image();
		// 渲染style样式
		this.renderStyleValues(elem.cssStyle, oImage);
		// 图片资源地址，base64/blob类型
		const source: string = await this.document.loadDocumentImage(
			elem.src,
			this.currentPart
		);
		if (is_clip || is_transform) {
			// canvas转换
			oImage.src = await this.transformImage(elem, source);
		} else {
			// 直接使用原图片
			oImage.src = source;
		}
		// 作为子元素插入，执行溢出检测
		if (parent) {
			oImage.dataset.overflow = await this.appendChildren(parent, oImage);
		}

		return oImage;
	}

	// 生成Konva框架--元素
	renderKonva() {
		// 创建konva容器元素
		const oContainer = createElement('div');
		oContainer.id = 'konva-container';
		// 插入页面底部
		document.body.appendChild(oContainer);
		// 创建Stage元素
		this.konva_stage = new Konva.Stage({ container: 'konva-container' });
		// 创建Layer元素
		this.konva_layer = new Konva.Layer({ listening: false });
		// 添加Stage元素
		this.konva_stage.add(this.konva_layer);
	}

	// canvas画布转换，处理旋转、裁剪、翻转等情况
	async transformImage(elem: WmlImage, source: string): Promise<string> {
		const { width, height, is_clip, clip, is_transform, transform } = elem.props;
		// 图片实例
		const img = new Image();
		// 设置图片源
		img.src = source;
		// 等待图片解码
		await img.decode();
		// 图片原始尺寸
		const { naturalWidth, naturalHeight } = img;
		// 显示Stage
		this.konva_stage.visible(true);
		// 设置Stage宽高
		this.konva_stage.width(naturalWidth);
		this.konva_stage.height(naturalHeight);
		// 设置Layer配置
		this.konva_layer.removeChildren();
		// 创建Group元素
		const group: Group = new Konva.Group();
		// 图片加载成功后创建Image
		const image = new Konva.Image({
			image: img,
			x: naturalWidth / 2,
			y: naturalHeight / 2,
			width: naturalWidth,
			height: naturalHeight,
			// 旋转中心
			offset: {
				x: naturalWidth / 2,
				y: naturalHeight / 2,
			},
		});
		// 计算裁剪参数
		if (is_clip) {
			const { left, right, top, bottom } = clip.path;
			const x = naturalWidth * left;
			const y = naturalHeight * top;
			const width = naturalWidth * (1 - left - right);
			const height = naturalHeight * (1 - top - bottom);
			image.crop({ x, y, width, height });
			image.size({ width, height });
		}
		// transform变换
		if (is_transform) {
			for (const key in transform) {
				switch (key) {
					case 'scaleX':
						image.scaleX(transform[key]);
						break;
					case 'scaleY':
						image.scaleY(transform[key]);
						break;
					case 'rotate':
						image.rotation(transform[key]);
						break;
				}
			}
		}
		// Group添加Image图片
		group.add(image);
		// 添加Group元素
		this.konva_layer.add(group);
		// 导出装换之后的图片
		let result: string | PromiseLike<string>;
		if (this.options.useBase64URL) {
			result = group.toDataURL();
		} else {
			const blob = (await group.toBlob()) as Blob;
			result = URL.createObjectURL(blob);
		}
		// 隐藏Stage
		this.konva_stage.visible(false);

		return result;
	}

	// 渲染书签，主要用于定位，导航
	renderBookmarkStart(elem: WmlBookmarkStart): HTMLElement {
		const oSpan = createElement('span');
		oSpan.id = elem.name;
		return oSpan;
	}

	// 渲染制表符
	async renderTab(elem: OpenXmlElement, parent?: HTMLElement | Element) {
		const tabSpan = createElement('span');

		tabSpan.innerHTML = '&emsp;'; //"&nbsp;";

		if (this.options.experimental) {
			tabSpan.className = this.tabStopClass();
			const stops = findParent<WmlParagraph>(elem, DomType.Paragraph)?.tabs;
			this.currentTabs.push({ stops, span: tabSpan });
		}

		// 作为子元素插入，执行溢出检测
		if (parent) {
			await this.appendChildren(parent, tabSpan);
		}

		return tabSpan;
	}

	async renderSymbol(elem: WmlSymbol, parent?: HTMLElement | Element) {
		const oSpan = createElement('span');
		oSpan.style.fontFamily = elem.font;
		oSpan.innerHTML = `&#x${elem.char};`;

		// 作为子元素插入，执行溢出检测
		if (parent) {
			await this.appendChildren(parent, oSpan);
		}

		return oSpan;
	}

	// 渲染换行符号
	async renderBreak(elem: WmlBreak, parent?: HTMLElement | Element) {
		let oBr: HTMLElement;

		switch (elem.break) {
			// 分页符
			case 'page':
				oBr = createElement('br');
				// 添加class
				oBr.classList.add('break', 'page');
				break;
			// 强制换行
			case 'textWrapping':
				oBr = createElement('br');
				// 添加class
				oBr.classList.add('break', 'textWrap');
				break;
			// 	TODO 分栏符
			case 'column':
				oBr = createElement('br');
				// 添加class
				oBr.classList.add('break', 'column');
				break;
			// 渲染至尾部分页
			case 'lastRenderedPageBreak':
				oBr = createElement('wbr');
				// 添加class
				oBr.classList.add('break', 'lastRenderedPageBreak');
				break;
			default:
		}
		// 作为子元素插入，忽略溢出检测
		if (parent) {
			appendChildren(parent, oBr);
		}
		return oBr;
	}

	// TODO 需要处理溢出
	renderInserted(elem: OpenXmlElement) {
		if (this.options.renderChanges) {
			return this.renderContainer(elem, 'ins');
		}

		return this.renderContainer(elem, 'span');
	}

	// TODO 需要处理溢出
	async renderDeleted(elem: OpenXmlElement) {
		if (this.options.renderChanges) {
			return await this.renderContainer(elem, 'del');
		}
		return null;
	}

	// TODO 需要处理溢出
	async renderDeletedText(elem: WmlText, parent?: HTMLElement | Element) {
		let oDeletedText: Text;

		if (this.options.renderEndnotes) {
			oDeletedText = document.createTextNode(elem.text);
			// TODO 作为子元素插入，执行溢出检测
			if (parent) {
				await this.appendChildren(parent, oDeletedText);
			}
		} else {
			oDeletedText = null;
		}
		return oDeletedText;
	}

	// 注释开始
	renderCommentRangeStart(commentStart: WmlCommentRangeStart) {
		if (!this.options.experimental) {
			return null;
		}

		return document.createComment(`start of comment #${commentStart.id}`);
	}

	// 注释结束
	renderCommentRangeEnd(commentEnd: WmlCommentRangeStart) {
		if (!this.options.experimental) {
			return null;
		}

		return document.createComment(`end of comment #${commentEnd.id}`);
	}

	// 注释
	renderCommentReference(commentRef: WmlCommentReference) {
		if (!this.options.experimental) {
			return null;
		}

		const comment = this.document.commentsPart?.commentMap[commentRef.id];

		if (!comment) return null;

		return document.createComment(
			`comment #${comment.id} by ${comment.author} on ${comment.date}`
		);
	}

	// 渲染页眉页脚
	async renderHeaderFooter(elem: OpenXmlElement, tagName: keyof HTMLElementTagNameMap) {
		const oElement: HTMLElement = createElement(tagName);
		// 渲染子元素
		await this.renderChildren(elem, oElement);
		// 渲染style样式
		this.renderStyleValues(elem.cssStyle, oElement);

		return oElement;
	}

	renderFootnoteReference(elem: WmlNoteReference) {
		const oSup = createElement('sup');
		this.currentFootnoteIds.push(elem.id);
		oSup.textContent = `${this.currentFootnoteIds.length}`;
		return oSup;
	}

	renderEndnoteReference(elem: WmlNoteReference) {
		const oSup = createElement('sup');
		this.currentEndnoteIds.push(elem.id);
		oSup.textContent = `${this.currentEndnoteIds.length}`;
		return oSup;
	}

	async renderVmlElement(elem: VmlElement, parent?: HTMLElement | Element): Promise<SVGElement> {
		const oSvg = createSvgElement('svg');

		oSvg.setAttribute('style', elem.cssStyleText);

		const oChildren = await this.renderVmlChildElement(elem);

		if (elem.imageHref?.id) {
			const source = await this.document?.loadDocumentImage(elem.imageHref.id, this.currentPart);
			oChildren.setAttribute('href', source);
		}
		// 后代元素忽略溢出检测
		appendChildren(oSvg, oChildren);

		requestAnimationFrame(() => {
			const bb = (oSvg.firstElementChild as any).getBBox();

			oSvg.setAttribute('width', `${Math.ceil(bb.x + bb.width)}`);
			oSvg.setAttribute('height', `${Math.ceil(bb.y + bb.height)}`);
		});
		// 如果拥有父级
		if (parent) {
			// 作为子元素插入,针对此元素进行溢出检测
			oSvg.dataset.overflow = await this.appendChildren(parent, oSvg);
		}
		return oSvg;
	}

	// 渲染VML中图片
	async renderVmlPicture(elem: OpenXmlElement) {
		const oPictureContainer = createElement('span');
		await this.renderChildren(elem, oPictureContainer);
		return oPictureContainer;
	}

	async renderVmlChildElement(elem: VmlElement) {
		const oVMLElement = createSvgElement(elem.tagName as any);
		// set attributes
		Object.entries(elem.attrs).forEach(([k, v]) => oVMLElement.setAttribute(k, v));

		for (const child of elem.children) {
			if (child.type == DomType.VmlElement) {
				const oChild = await this.renderVmlChildElement(child as VmlElement);
				appendChildren(oVMLElement, oChild);
			} else {
				await this.renderElement(child as any, oVMLElement);
			}
		}

		return oVMLElement;
	}

	async renderMmlRadical(elem: OpenXmlElement) {
		const base = elem.children.find(el => el.type == DomType.MmlBase);
		let oParent: HTMLElement | Element;
		if (elem.props?.hideDegree) {
			oParent = createElementNS(ns.mathML, 'msqrt', null);
			await this.renderElements([base], oParent);
			return oParent;
		}

		const degree = elem.children.find(el => el.type == DomType.MmlDegree);
		oParent = createElementNS(ns.mathML, 'mroot', null);
		await this.renderElements([base, degree], oParent);
		return oParent;
	}

	async renderMmlDelimiter(elem: OpenXmlElement): Promise<MathMLElement> {
		const oMrow = createElementNS(ns.mathML, 'mrow', null);
		// 开始Char
		let oBegin = createElementNS(ns.mathML, "mo", null, [elem.props.beginChar ?? '(']);
		appendChildren(oMrow, oBegin);
		// 子元素
		await this.renderElements(elem.children, oMrow);
		// 结束char
		let oEnd = createElementNS(ns.mathML, "mo", null, [elem.props.endChar ?? ')']);
		appendChildren(oMrow, oEnd);

		return oMrow;
	}

	async renderMmlNary(elem: OpenXmlElement): Promise<MathMLElement> {
		const children = [];
		const grouped = keyBy(elem.children, x => x.type);

		const sup = grouped[DomType.MmlSuperArgument];
		const sub = grouped[DomType.MmlSubArgument];

		let supElem = sup ? createElementNS(ns.mathML, "mo", null, asArray(await this.renderElement(sup))) : null;
		let subElem = sub ? createElementNS(ns.mathML, "mo", null, asArray(await this.renderElement(sub))) : null;

		let charElem = createElementNS(ns.mathML, "mo", null, [elem.props?.char ?? '\u222B']);

		if (supElem || subElem) {
			children.push(createElementNS(ns.mathML, "munderover", null, [charElem, subElem, supElem]));
		} else if (supElem) {
			children.push(createElementNS(ns.mathML, "mover", null, [charElem, supElem]));
		} else if (subElem) {
			children.push(createElementNS(ns.mathML, "munder", null, [charElem, subElem]));
		} else {
			children.push(charElem);
		}

		const oMrow = createElementNS(ns.mathML, 'mrow', null);

		appendChildren(oMrow, children);

		await this.renderElements(grouped[DomType.MmlBase].children, oMrow);

		return oMrow;
	}

	async renderMmlPreSubSuper(elem: OpenXmlElement) {
		const children = [];
		const grouped = keyBy(elem.children, x => x.type);

		const sup = grouped[DomType.MmlSuperArgument];
		const sub = grouped[DomType.MmlSubArgument];
		let supElem = sup ? createElementNS(ns.mathML, "mo", null, asArray(await this.renderElement(sup))) : null;
		let subElem = sub ? createElementNS(ns.mathML, "mo", null, asArray(await this.renderElement(sub))) : null;
		let stubElem = createElementNS(ns.mathML, "mo", null);

		children.push(createElementNS(ns.mathML, "msubsup", null, [stubElem, subElem, supElem]));

		const oMrow = createElementNS(ns.mathML, 'mrow', null);

		appendChildren(oMrow, children);

		await this.renderElements(grouped[DomType.MmlBase].children, oMrow);

		return oMrow;
	}

	async renderMmlGroupChar(elem: OpenXmlElement) {
		let tagName = elem.props.verticalJustification === "bot" ? "mover" : "munder";
		let oGroupChar = await this.renderContainerNS(elem, ns.mathML, tagName);

		if (elem.props.char) {
			const oMo = createElementNS(ns.mathML, 'mo', null, [elem.props.char]);
			appendChildren(oGroupChar, oMo);
		}

		return oGroupChar;
	}

	async renderMmlBar(elem: OpenXmlElement) {
		let oMrow = await this.renderContainerNS(elem, ns.mathML, "mrow") as MathMLElement;

		switch (elem.props.position) {
			case 'top':
				oMrow.style.textDecoration = 'overline';
				break;
			case 'bottom':
				oMrow.style.textDecoration = 'underline';
				break;
		}

		return oMrow;
	}

	async renderMmlRun(elem: OpenXmlElement) {
		const oMs = createElementNS(ns.mathML, 'ms') as HTMLElement;

		this.renderClass(elem, oMs);
		this.renderStyleValues(elem.cssStyle, oMs);
		await this.renderChildren(elem, oMs);

		return oMs;
	}

	async renderMllList(elem: OpenXmlElement) {
		const oMtable = createElementNS(ns.mathML, 'mtable') as HTMLElement;
		// 添加class类
		this.renderClass(elem, oMtable);
		// 渲染style样式
		this.renderStyleValues(elem.cssStyle, oMtable);

		for (const child of elem.children) {
			const oChild = await this.renderElement(child);

			const oMtd = createElementNS(ns.mathML, 'mtd', null, [oChild]);

			const oMtr = createElementNS(ns.mathML, 'mtr', null, [oMtd]);

			appendChildren(oMtable, oMtr);
		}

		return oMtable;
	}

	// 设置元素style样式
	renderStyleValues(style: Record<string, string>, output: HTMLElement) {
		for (const k in style) {
			if (k.startsWith('$')) {
				output.setAttribute(k.slice(1), style[k]);
			} else {
				output.style[k] = style[k];
			}
		}
	}

	renderRunProperties(style: any, props: RunProperties) {
		this.renderCommonProperties(style, props);
	}

	renderCommonProperties(style: any, props: CommonProperties) {
		if (props == null) return;

		if (props.color) {
			style['color'] = props.color;
		}

		if (props.fontSize) {
			style['font-size'] = props.fontSize;
		}
	}

	// 添加class类名
	renderClass(input: OpenXmlElement, output: HTMLElement | Element) {
		if (input.className) {
			output.className = input.className;
		}

		if (input.styleName) {
			output.classList.add(this.processStyleName(input.styleName));
		}
	}

	// 查找内置默认style样式
	findStyle(styleName: string) {
		return styleName && this.styleMap?.[styleName];
	}

	tabStopClass() {
		return `${this.className}-tab-stop`;
	}

	// 刷新tab制表符
	refreshTabStops() {
		if (!this.options.experimental) {
			return;
		}

		clearTimeout(this.tabsTimeout);

		this.tabsTimeout = setTimeout(() => {
			const pixelToPoint = computePixelToPoint();

			for (const tab of this.currentTabs) {
				updateTabStop(tab.span, tab.stops, this.defaultTabSize, pixelToPoint);
			}
		}, 500);
	}
}

/**
 *  操作DOM元素的函数方法
 *
 */

// 元素类型
type ChildrenType = Node[] | Node | Element[] | Element;

// 根据标签名tagName创建元素
function createElement<T extends keyof HTMLElementTagNameMap>(tagName: T, props?: Partial<Record<keyof HTMLElementTagNameMap[T], any>>): HTMLElementTagNameMap[T] {
	return createElementNS(null, tagName, props);
}

// 根据标签名tagName创建svg元素
function createSvgElement<T extends keyof SVGElementTagNameMap>(tagName: T, props?: Partial<Record<keyof SVGElementTagNameMap[T], any>>): SVGElementTagNameMap[T] {
	return createElementNS(ns.svg, tagName, props);
}

// 根据标签名tagName创建带命名空间的元素
function createElementNS<T extends keyof MathMLElementTagNameMap>(ns: string, tagName: T, props?: Partial<Record<any, any>>, children?: ChildrenType): MathMLElementTagNameMap[T];
function createElementNS<T extends keyof SVGElementTagNameMap>(ns: string, tagName: T, props?: Partial<Record<any, any>>, children?: ChildrenType): SVGElementTagNameMap[T];
function createElementNS<T extends keyof HTMLElementTagNameMap>(ns: string, tagName: T, props?: Partial<Record<any, any>>, children?: ChildrenType): HTMLElementTagNameMap[T];
function createElementNS<T>(ns: string, tagName: T, props?: Partial<Record<any, any>>, children?: ChildrenType): Element | SVGElement | MathMLElement {
	let oParent: Element | SVGElement | MathMLElement;
	switch (ns) {
		case "http://www.w3.org/1998/Math/MathML":
			oParent = document.createElementNS(ns, tagName as keyof MathMLElementTagNameMap);
			break;
		case 'http://www.w3.org/2000/svg':
			oParent = document.createElementNS(ns, tagName as keyof SVGElementTagNameMap);
			break;
		case 'http://www.w3.org/1999/xhtml':
			oParent = document.createElement(tagName as keyof HTMLElementTagNameMap);
			break;
		default:
			oParent = document.createElement(tagName as keyof HTMLElementTagNameMap);
	}

	if (props) {
		Object.assign(oParent, props);
	}

	if (children) {
		appendChildren(oParent, children);
	}

	return oParent;
}

// 清空所有子元素
function removeAllElements(elem: HTMLElement) {
	elem.innerHTML = '';
}

// 插入子元素
function appendChildren(parent: Element, children: ChildrenType): void {
	if (Array.isArray(children)) {
		parent.append(...children);
	} else if (children) {
		if (isString(children)) {
			parent.append(children);
		} else {
			parent.appendChild(children);
		}
	}
}

// 判断文本区是否溢出
function checkOverflow(el: HTMLElement) {
	// 提取原来的overflow属性值
	const current_overflow: string = getComputedStyle(el).overflow;
	//先让溢出效果为 hidden 这样才可以比较 clientHeight和scrollHeight
	if (!current_overflow || current_overflow === 'visible') {
		el.style.overflow = 'hidden';
	}
	const is_overflow: boolean = el.clientHeight < el.scrollHeight;

	// 还原overflow属性值
	el.style.overflow = current_overflow;

	return is_overflow;
}

// 删除单个或者多个子元素
function removeElements(target: Node[] | Node, parent: HTMLElement | Element): void;
function removeElements(target: Element[] | Element): void;
function removeElements(target: ChildrenType, parent?: HTMLElement | Element): void {
	if (Array.isArray(target)) {
		target.forEach(elem => {
			if (elem instanceof Element) {
				elem.remove();
			} else {
				if (parent) {
					parent.removeChild(elem);
				}
			}
		});
	} else {
		if (target instanceof Element) {
			target.remove();
		} else {
			if (target) {
				parent.removeChild(target);
			}
		}
	}
}

// 创建style标签
function createStyleElement(cssText: string) {
	return createElement('style', { innerHTML: cssText });
}

// 插入注释
function appendComment(elem: HTMLElement, comment: string) {
	elem.appendChild(document.createComment(comment));
}

// 根据元素类型，回溯元素的父级元素、祖先元素
function findParent<T extends OpenXmlElement>(elem: OpenXmlElement, type: DomType): T {
	let parent = elem.parent;

	while (parent != null && parent.type != type) {
		parent = parent.parent;
	}

	return <T>parent;
}
