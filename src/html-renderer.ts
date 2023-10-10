import { WordDocument } from './word-document';
import {
    DomType,
    WmlTable,
    IDomNumbering,
    WmlHyperlink,
    IDomImage,
    OpenXmlElement,
    WmlTableColumn,
    WmlTableCell,
    WmlText,
    WmlSymbol,
    WmlBreak,
    WmlNoteReference
} from './document/dom';
import { CommonProperties } from './document/common';
import { Options } from './docx-preview';
import { DocumentElement } from './document/document';
import { WmlParagraph } from './document/paragraph';
import { asArray, escapeClassName, isString, keyBy, mergeDeep } from './utils';
import { computePixelToPoint, updateTabStop } from './javascript';
import { FontTablePart } from './font-table/font-table';
import { FooterHeaderReference, SectionProperties, Section } from './document/section';
import { WmlRun, RunProperties } from './document/run';
import { WmlBookmarkStart } from './document/bookmarks';
import { IDomStyle } from './document/style';
import { WmlBaseNote, WmlFootnote } from './notes/elements';
import { ThemePart } from './theme/theme-part';
import { BaseHeaderFooterPart } from './header-footer/parts';
import { Part } from './common/part';
import mathMLCSS from "./mathml.scss";
import { VmlElement } from './vml/vml';

const ns = {
    svg: "http://www.w3.org/2000/svg",
    mathML: "http://www.w3.org/1998/Math/MathML"
}

interface CellPos {
    col: number;
    row: number;
}

type CellVerticalMergeType = Record<number, HTMLTableCellElement>;

// HTML渲染器

export class HtmlRenderer {

    className: string = "docx";
    rootSelector: string;
    document: WordDocument;
    options: Options;
    styleMap: Record<string, IDomStyle> = {};
    currentPart: Part = null;

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

    constructor(public htmlDocument: Document) {
    }

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
        // styleContainer== null，styleContainer = bodyContainer
        styleContainer = styleContainer || bodyContainer;

        // CSS样式生成容器，清空所有CSS样式
        removeAllElements(styleContainer);
        // HTML生成容器，清空所有HTML元素
        removeAllElements(bodyContainer);

        // 添加注释
        appendComment(styleContainer, "docxjs library predefined styles");
        // 添加默认CSS样式
        styleContainer.appendChild(this.renderDefaultStyle());

        // 数学公式CSS样式
        if (!window.MathMLElement && options.useMathMLPolyfill) {
            appendComment(styleContainer, "docxjs mathml polyfill styles");
            styleContainer.appendChild(createStyleElement(mathMLCSS));
        }
        // 主题CSS样式
        if (document.themePart) {
            appendComment(styleContainer, "docxjs document theme values");
            this.renderTheme(document.themePart, styleContainer);
        }
        // 文档默认CSS样式，包含表格、列表、段落、字体，样式存在继承顺序
        if (document.stylesPart != null) {
            this.styleMap = this.processStyles(document.stylesPart.styles);

            appendComment(styleContainer, "docxjs document styles");
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
        // 主文档--内容
        let sectionElements = await this.renderSections(document.documentPart.body);
        if (this.options.inWrapper) {
            bodyContainer.appendChild(this.renderWrapper(sectionElements));
        } else {
            appendChildren(bodyContainer, sectionElements);
        }

        // 刷新制表符
        this.refreshTabStops();
    }

    /**
     * Object对象 => HTML字符串
     *
     * @param document word文档Object对象
     * @param styleContainer CSS样式生成容器
     * @param options 渲染配置选项
     */
    async renderFragment(document: WordDocument, styleContainer: HTMLElement = null, options: Options) {
        this.document = document;
        this.options = options;
        // class类前缀
        this.className = options.className;
        // 根元素
        this.rootSelector = options.inWrapper ? `.${this.className}-wrapper` : ':root';
        // 文档CSS样式
        this.styleMap = null;
        // 生成代码片段实例
        const template = window.document.createElement('template');
        // CSS样式生成容器，清空所有CSS样式
        removeAllElements(styleContainer);

        // 添加注释
        appendComment(styleContainer, "docxjs library predefined styles");
        // 添加默认CSS样式
        styleContainer.appendChild(this.renderDefaultStyle());

        // 数学公式CSS样式
        if (!window.MathMLElement && options.useMathMLPolyfill) {
            appendComment(styleContainer, "docxjs mathml polyfill styles");
            styleContainer.appendChild(createStyleElement(mathMLCSS));
        }
        // 主题CSS样式
        if (document.themePart) {
            appendComment(styleContainer, "docxjs document theme values");
            this.renderTheme(document.themePart, styleContainer);
        }

        // 文档默认CSS样式，包含表格、列表、段落、字体，样式存在继承顺序
        if (document.stylesPart != null) {
            this.styleMap = this.processStyles(document.stylesPart.styles);

            appendComment(styleContainer, "docxjs document styles");
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
        // 主文档--section内容
        let sectionElements = await this.renderSections(document.documentPart.body);
        if (this.options.inWrapper) {
            template.appendChild(this.renderWrapper(sectionElements));
        } else {
            appendChildren(template, sectionElements);
        }
        // 刷新制表符
        this.refreshTabStops();

        return template.innerHTML;
    }

    // 文档CSS主题样式
    renderTheme(themePart: ThemePart, styleContainer: HTMLElement | DocumentFragment) {
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
            for (let [k, v] of Object.entries(colorScheme.colors)) {
                variables[`--docx-${k}-color`] = `#${v}`;
            }
        }

        const cssText = this.styleToString(`.${this.className}`, variables);
        styleContainer.appendChild(createStyleElement(cssText));
    }

    // 字体列表CSS样式
    renderFontTable(fontsPart: FontTablePart, styleContainer: HTMLElement | DocumentFragment) {
        for (let f of fontsPart.fonts) {
            for (let ref of f.embedFontRefs) {
                this.document.loadFont(ref.id, ref.key).then(fontData => {
                    const cssValues = {
                        'font-family': f.name,
                        'src': `url(${fontData})`
                    };

                    if (ref.type == "bold" || ref.type == "boldItalic") {
                        cssValues['font-weight'] = 'bold';
                    }

                    if (ref.type == "italic" || ref.type == "boldItalic") {
                        cssValues['font-style'] = 'italic';
                    }

                    appendComment(styleContainer, `docxjs ${f.name} font`);
                    const cssText = this.styleToString("@font-face", cssValues);
                    styleContainer.appendChild(createStyleElement(cssText));
                    this.refreshTabStops();
                });
            }
        }
    }

    // 计算className，小写，默认前缀："docx_"
    processStyleName(className: string): string {
        return className ? `${this.className}_${escapeClassName(className)}` : this.className;
    }

    // 处理样式继承
    processStyles(styles: IDomStyle[]) {
        //
        const stylesMap = keyBy(styles.filter(x => x.id != null), x => x.id);
        // 遍历base_on关系,合并样式
        for (const style of styles.filter(x => x.basedOn)) {
            let baseStyle = stylesMap[style.basedOn];

            if (baseStyle) {
                // 深度合并
                style.paragraphProps = mergeDeep(style.paragraphProps, baseStyle.paragraphProps);
                style.runProps = mergeDeep(style.runProps, baseStyle.runProps);

                for (const baseValues of baseStyle.styles) {
                    const styleValues = style.styles.find(x => x.target == baseValues.target);

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

        for (let style of styles) {
            style.cssName = this.processStyleName(style.id);
        }

        return stylesMap;
    }

    processNumberings(numberings: IDomNumbering[]) {
        for (let num of numberings.filter(n => n.pStyleName)) {
            const style = this.findStyle(num.pStyleName);

            if (style?.paragraphProps?.numbering) {
                style.paragraphProps.numbering.level = num.level;
            }
        }
    }

    // 递归明确元素parent父级关系
    processElement(element: OpenXmlElement) {
        if (element.children) {
            for (let e of element.children) {
                e.parent = element;
                // 判断类型
                if (e.type == DomType.Table) {
                    // 渲染表格
                    this.processTable(e);
                } else {
                    // 递归渲染
                    this.processElement(e);
                }
            }
        }
    }

    // 表格style样式
    processTable(table: WmlTable) {
        for (let r of table.children) {
            for (let c of r.children) {
                c.cssStyle = this.copyStyleProperties(table.cellStyle, c.cssStyle, [
                    "border-left", "border-right", "border-top", "border-bottom",
                    "padding-left", "padding-right", "padding-top", "padding-bottom"
                ]);

                this.processElement(c);
            }
        }
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

        for (let key of attrs) {
            if (input.hasOwnProperty(key) && !output.hasOwnProperty(key))
                output[key] = input[key];
        }

        return output;
    }

    // 创建Page Section
    createSection(className: string, props: SectionProperties) {
        let elem = this.createElement("section", { className });

        if (props) {
            if (props.pageMargins) {
                elem.style.paddingLeft = props.pageMargins.left;
                elem.style.paddingRight = props.pageMargins.right;
                elem.style.paddingTop = props.pageMargins.top;
                elem.style.paddingBottom = props.pageMargins.bottom;
            }

            if (props.pageSize) {
                if (!this.options.ignoreWidth)
                    elem.style.width = props.pageSize.width;
                if (!this.options.ignoreHeight)
                    elem.style.minHeight = props.pageSize.height;
            }

            if (props.columns && props.columns.numberOfColumns) {
                elem.style.columnCount = `${props.columns.numberOfColumns}`;
                elem.style.columnGap = props.columns.space;

                if (props.columns.separator) {
                    elem.style.columnRule = "1px solid black";
                }
            }
        }

        return elem;
    }

    // 生成Page Section
    async renderSections(document: DocumentElement): Promise<HTMLElement[]> {
        const result = [];
        // 生成页面parent父级关系
        this.processElement(document);
        // 根据options.breakPages，判断是否分页
        let sections: Section[];
        if (this.options.breakPages) {
            // 根据section切分页面
            sections = this.splitBySection(document.children);
        } else {
            // 不分页则，只有一个section
            sections = [{ sectProps: document.props, elements: document.children }];
        }

        let prevProps = null;
        // 遍历生成每一个section
        for (let i = 0, l = sections.length; i < l; i++) {
            this.currentFootnoteIds = [];

            const section = sections[i];
            const props = section.sectProps || document.props;
            // 根据sectProps，创建section
            const sectionElement = this.createSection(this.className, props);
            // 给section添加style样式
            this.renderStyleValues(document.cssStyle, sectionElement);
            // 渲染页眉
            if (this.options.renderHeaders) {
                await this.renderHeaderFooter(props.headerRefs, props, result.length, prevProps != props, sectionElement);
            }
            // 渲染Page内容
            let contentElement = this.createElement("article");
            await this.renderElements(section.elements, contentElement);
            sectionElement.appendChild(contentElement);
            // 渲染脚注
            if (this.options.renderFootnotes) {
                await this.renderNotes(this.currentFootnoteIds, this.footnoteMap, sectionElement);
            }
            // 渲染尾注
            if (this.options.renderEndnotes && i == l - 1) {
                await this.renderNotes(this.currentEndnoteIds, this.endnoteMap, sectionElement);
            }
            // 渲染页脚
            if (this.options.renderFooters) {
                await this.renderHeaderFooter(props.footerRefs, props, result.length, prevProps != props, sectionElement);
            }

            result.push(sectionElement);
            prevProps = props;
        }

        return result;
    }

    // 渲染页眉/页脚
    async renderHeaderFooter(refs: FooterHeaderReference[], props: SectionProperties, page: number, firstOfSection: boolean, into: HTMLElement) {
        if (!refs) return;
        // 查找奇数偶数的ref指向
        let ref = (props.titlePage && firstOfSection ? refs.find(x => x.type == "first") : null)
            ?? (page % 2 == 1 ? refs.find(x => x.type == "even") : null)
            ?? refs.find(x => x.type == "default");

        // 查找ref对应的part部分
        let part = ref && this.document.findPartByRelId(ref.id, this.document.documentPart) as BaseHeaderFooterPart;

        if (part) {
            this.currentPart = part;
            if (!this.usedHederFooterParts.includes(part.path)) {
                this.processElement(part.rootElement);
                this.usedHederFooterParts.push(part.path);
            }

            await this.renderElements([part.rootElement], into);
            this.currentPart = null;
        }
    }

    // 判断是否存在分页元素
    isPageBreakElement(elem: OpenXmlElement): boolean {
        // 分页符、换行符、分栏符
        if (elem.type != DomType.Break) {
            return false;
        }
        // 默认以lastRenderedPageBreak作为分页依据
        if ((elem as WmlBreak).break == "lastRenderedPageBreak") {
            return !this.options.ignoreLastRenderedPageBreak;
        }
        // 分页符
        if ((elem as WmlBreak).break === "page") {
            return true;
        }
    }

    // 根据section切分页面
    splitBySection(elements: OpenXmlElement[]): Section[] {
        // 当前操作section，elements数组包含子元素
        let current_section = { sectProps: null, elements: [] };
        // 切分出的所有sections
        let sections = [current_section];

        for (let elem of elements) {
            /* 段落基本结构：paragraph => run => text... */
            if (elem.type == DomType.Paragraph) {
                const p = elem as WmlParagraph;
                // 节属性，代表分节符
                let sectProps = p.sectionProps;

                /*
                    检测段落是否默认存在强制分页符
                */

                // 查找内置默认段落样式
                const default_paragraph_style = this.findStyle(p.styleName);

                // 段落内置样式之前存在强制分页符
                if (default_paragraph_style?.paragraphProps?.pageBreakBefore) {
                    // 保存当前section的sectionProps
                    current_section.sectProps = sectProps;
                    // 重置新的section
                    current_section = { sectProps: null, elements: [] };
                    // 添加新section
                    sections.push(current_section);
                }
            }

            // 添加elem进入当前操作section
            current_section.elements.push(elem);

            /* 段落基本结构：paragraph => run => text... */
            if (elem.type == DomType.Paragraph) {
                const p = elem as WmlParagraph;
                // 节属性
                let sectProps = p.sectionProps;
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
                            // 分页符、换行符、分栏符
                            if (t.type != DomType.Break) {
                                return false;
                            }
                            // 默认忽略lastRenderedPageBreak，
                            if ((t as WmlBreak).break == "lastRenderedPageBreak") {
                                // 判断前一个p段落，
                                // 如果含有分页符、分节符，那它们一定位于上一个section，
                                // 如果前一个段落是普通段落，则代表文字过多超过一页，需要自动分页
                                return current_section.elements.length > 2 || !this.options.ignoreLastRenderedPageBreak;
                            }
                            // 分页符
                            if ((t as WmlBreak).break === "page") {
                                return true;
                            }
                        });
                        rBreakIndex = rBreakIndex ?? -1
                        return rBreakIndex != -1;
                    });
                }

                // 段落中存在节属性sectProps/段落Break索引
                if (sectProps || pBreakIndex != -1) {
                    // 保存当前section的sectionProps
                    current_section.sectProps = sectProps;
                    // 重置新的section
                    current_section = { sectProps: null, elements: [] };
                    // 添加新section
                    sections.push(current_section);
                }

                // 根据段落Break索引，拆分Run部分
                if (pBreakIndex != -1) {
                    // 即将拆分的Run部分
                    let breakRun = p.children[pBreakIndex];
                    // 是否需要拆分Run
                    let is_split = rBreakIndex < breakRun.children.length - 1;

                    if (pBreakIndex < p.children.length - 1 || is_split) {
                        // 原始的Run
                        let origin_run = p.children;
                        // 切出Break索引后面的Run，创建新段落
                        let new_paragraph = { ...p, children: origin_run.slice(pBreakIndex) };
                        // 保存Break索引前面的Run
                        p.children = origin_run.slice(0, pBreakIndex);
                        // 添加新段落
                        current_section.elements.push(new_paragraph);

                        if (is_split) {
                            // Run下面原始的元素
                            let origin_elements = breakRun.children;
                            // 切出Run Break索引前面的元素，创建新Run
                            let newRun = { ...breakRun, children: origin_elements.slice(0, rBreakIndex) };
                            // 将新Run放入上一个section的段落
                            p.children.push(newRun);
                            // 切出Run Break索引后面的元素
                            breakRun.children = origin_elements.slice(rBreakIndex);
                        }
                    }
                }
            }

            // TODO elem元素是表格，拆分section
            if (elem.type === DomType.Table) {
                // console.log(elem)
            }
        }

        // 处理所有section的section_props
        let currentSectProps = null;
        // 倒序
        for (let i = sections.length - 1; i >= 0; i--) {

            if (sections[i].sectProps == null) {
                sections[i].sectProps = currentSectProps;
            } else {
                currentSectProps = sections[i].sectProps
            }
        }
        console.log(sections)
        return sections;
    }

    // 生成父级容器
    renderWrapper(children: HTMLElement[]) {
        return this.createElement("div", { className: `${this.className}-wrapper` }, children);
    }

    // 渲染默认样式
    renderDefaultStyle() {
        let c = this.className;
        let styleText = `
			.${c}-wrapper { background: gray; padding: 30px; padding-bottom: 0px; display: flex; flex-flow: column; align-items: center; } 
			.${c}-wrapper>section.${c} { background: white; box-shadow: 0 0 10px rgba(0, 0, 0, 0.5); margin-bottom: 30px; }
			.${c} { color: black; hyphens: auto; }
			section.${c} { box-sizing: border-box; display: flex; flex-flow: column nowrap; position: relative; overflow: hidden; }
			section.${c}>article { margin-bottom: auto; z-index: 1; }
			section.${c}>footer { z-index: 1; }
			.${c} table { border-collapse: collapse; }
			.${c} table td, .${c} table th { vertical-align: top; }
			.${c} p { margin: 0pt; min-height: 1em; }
			.${c} span { white-space: pre-wrap; overflow-wrap: break-word; }
			.${c} a { color: inherit; text-decoration: inherit; }
		`;

        return createStyleElement(styleText);
    }

    // renderNumbering2(numberingPart: NumberingPartProperties, container: HTMLElement): HTMLElement {
    //     let css = "";
    //     const numberingMap = keyBy(numberingPart.abstractNumberings, x => x.id);
    //     const bulletMap = keyBy(numberingPart.bulletPictures, x => x.id);
    //     const topCounters = [];

    //     for(let num of numberingPart.numberings) {
    //         const absNum = numberingMap[num.abstractId];

    //         for(let lvl of absNum.levels) {
    //             const className = this.numberingClass(num.id, lvl.level);
    //             let listStyleType = "none";

    //             if(lvl.text && lvl.format == 'decimal') {
    //                 const counter = this.numberingCounter(num.id, lvl.level);

    //                 if (lvl.level > 0) {
    //                     css += this.styleToString(`p.${this.numberingClass(num.id, lvl.level - 1)}`, {
    //                         "counter-reset": counter
    //                     });
    //                 } else {
    //                     topCounters.push(counter);
    //                 }

    //                 css += this.styleToString(`p.${className}:before`, {
    //                     "content": this.levelTextToContent(lvl.text, num.id),
    //                     "counter-increment": counter
    //                 });
    //             } else if(lvl.bulletPictureId) {
    //                 let pict = bulletMap[lvl.bulletPictureId];
    //                 let variable = `--${this.className}-${pict.referenceId}`.toLowerCase();

    //                 css += this.styleToString(`p.${className}:before`, {
    //                     "content": "' '",
    //                     "display": "inline-block",
    //                     "background": `var(${variable})`
    //                 }, pict.style);

    //                 this.document.loadNumberingImage(pict.referenceId).then(data => {
    //                     var text = `.${this.className}-wrapper { ${variable}: url(${data}) }`;
    //                     container.appendChild(createStyleElement(text));
    //                 });
    //             } else {
    //                 listStyleType = this.numFormatToCssValue(lvl.format);
    //             }

    //             css += this.styleToString(`p.${className}`, {
    //                 "display": "list-item",
    //                 "list-style-position": "inside",
    //                 "list-style-type": listStyleType,
    //                 //TODO
    //                 //...num.style
    //             });
    //         }
    //     }

    //     if (topCounters.length > 0) {
    //         css += this.styleToString(`.${this.className}-wrapper`, {
    //             "counter-reset": topCounters.join(" ")
    //         });
    //     }

    //     return createStyleElement(css);
    // }

    renderNumbering(numberings: IDomNumbering[], styleContainer: HTMLElement | DocumentFragment) {
        let styleText = "";
        let resetCounters = [];

        for (let num of numberings) {
            let selector = `p.${this.numberingClass(num.id, num.level)}`;
            let listStyleType = "none";

            if (num.bullet) {
                let valiable = `--${this.className}-${num.bullet.src}`.toLowerCase();

                styleText += this.styleToString(`${selector}:before`, {
                    "content": "' '",
                    "display": "inline-block",
                    "background": `var(${valiable})`
                }, num.bullet.style);

                this.document.loadNumberingImage(num.bullet.src).then(data => {
                    let text = `${this.rootSelector} { ${valiable}: url(${data}) }`;
                    styleContainer.appendChild(createStyleElement(text));
                });
            } else if (num.levelText) {
                let counter = this.numberingCounter(num.id, num.level);
                const counterReset = counter + " " + (num.start - 1);
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
                "display": "list-item",
                "list-style-position": "inside",
                "list-style-type": listStyleType,
                ...num.pStyle
            });
        }

        if (resetCounters.length > 0) {
            styleText += this.styleToString(this.rootSelector, {
                "counter-reset": resetCounters.join(" ")
            });
        }

        return createStyleElement(styleText);
    }

    renderStyles(styles: IDomStyle[]): HTMLElement {
        let styleText = "";
        const stylesMap = this.styleMap;
        const defautStyles = keyBy(styles.filter(s => s.isDefault), s => s.target);

        for (const style of styles) {
            let subStyles = style.styles;

            if (style.linked) {
                let linkedStyle = style.linked && stylesMap[style.linked];

                if (linkedStyle)
                    subStyles = subStyles.concat(linkedStyle.styles);
                else if (this.options.debug)
                    console.warn(`Can't find linked style ${style.linked}`);
            }

            for (const subStyle of subStyles) {
                //TODO temporary disable modificators until test it well
                let selector = `${style.target ?? ''}.${style.cssName}`; //${subStyle.mod ?? ''}

                if (style.target != subStyle.target)
                    selector += ` ${subStyle.target}`;

                if (defautStyles[style.target] == style)
                    selector = `.${this.className} ${style.target}, ` + selector;

                styleText += this.styleToString(selector, subStyle.values);
            }
        }

        return createStyleElement(styleText);
    }

    // 渲染脚注/尾注
    async renderNotes(noteIds: string[], notesMap: Record<string, WmlBaseNote>, parent: HTMLElement) {
        let notes = noteIds.map(id => notesMap[id]).filter(x => x);

        if (notes.length > 0) {
            let children = await this.renderElements(notes);
            let result = this.createElement("ol", null, children);
            parent.appendChild(result);
        }
    }

    async renderElement(elem: OpenXmlElement): Promise<Node | Node[]> {
        switch (elem.type) {
            case DomType.Paragraph:
                return this.renderParagraph(elem as WmlParagraph);

            case DomType.BookmarkStart:
                return this.renderBookmarkStart(elem as WmlBookmarkStart);

            case DomType.BookmarkEnd:
                return null; //ignore bookmark end

            case DomType.Run:
                return this.renderRun(elem as WmlRun);

            case DomType.Table:
                return this.renderTable(elem);

            case DomType.Row:
                return this.renderTableRow(elem);

            case DomType.Cell:
                return this.renderTableCell(elem);

            case DomType.Hyperlink:
                return this.renderHyperlink(elem);

            case DomType.Drawing:
                return this.renderDrawing(elem);

            case DomType.Image:
                return await this.renderImage(elem as IDomImage);

            case DomType.Text:
                return this.renderText(elem as WmlText);

            case DomType.DeletedText:
                return this.renderDeletedText(elem as WmlText);

            case DomType.Tab:
                return this.renderTab(elem);

            case DomType.Symbol:
                return this.renderSymbol(elem as WmlSymbol);

            case DomType.Break:
                return this.renderBreak(elem as WmlBreak);

            case DomType.Footer:
                return this.renderContainer(elem, "footer");

            case DomType.Header:
                // 修复绝对定位bug
                elem.children[0].cssStyle = { ...elem.children[0].cssStyle, position: "relative" };
                return this.renderContainer(elem, "header");

            case DomType.Footnote:
            case DomType.Endnote:
                return this.renderContainer(elem, "li");

            case DomType.FootnoteReference:
                return this.renderFootnoteReference(elem as WmlNoteReference);

            case DomType.EndnoteReference:
                return this.renderEndnoteReference(elem as WmlNoteReference);

            case DomType.NoBreakHyphen:
                return this.createElement("wbr");

            case DomType.VmlPicture:
                return this.renderVmlPicture(elem);

            case DomType.VmlElement:
                return this.renderVmlElement(elem as VmlElement);

            case DomType.MmlMath:
                return this.renderContainerNS(elem, ns.mathML, "math", { xmlns: ns.mathML });

            case DomType.MmlMathParagraph:
                return this.renderContainer(elem, "span");

            case DomType.MmlFraction:
                return this.renderContainerNS(elem, ns.mathML, "mfrac");

            case DomType.MmlNumerator:
            case DomType.MmlDenominator:
                return this.renderContainerNS(elem, ns.mathML, "mrow");

            case DomType.MmlRadical:
                return this.renderMmlRadical(elem);

            case DomType.MmlDegree:
                return this.renderContainerNS(elem, ns.mathML, "mn");

            case DomType.MmlSuperscript:
                return this.renderContainerNS(elem, ns.mathML, "msup");

            case DomType.MmlSubscript:
                return this.renderContainerNS(elem, ns.mathML, "msub");

            case DomType.MmlBase:
                return this.renderContainerNS(elem, ns.mathML, "mrow");

            case DomType.MmlSuperArgument:
                return this.renderContainerNS(elem, ns.mathML, "mn");

            case DomType.MmlSubArgument:
                return this.renderContainerNS(elem, ns.mathML, "mn");

            case DomType.MmlDelimiter:
                return this.renderMmlDelimiter(elem);

            case DomType.MmlRun:
                return this.renderMmlRun(elem);

            case DomType.MmlNary:
                return this.renderMmlNary(elem);

            case DomType.MmlEquationArray:
                return this.renderMllList(elem);

            case DomType.Inserted:
                return this.renderInserted(elem);

            case DomType.Deleted:
                return this.renderDeleted(elem);
            default:
                return null;
        }

    }

    async renderChildren(elem: OpenXmlElement, into?: Element): Promise<Node[]> {
        return await this.renderElements(elem.children, into);
    }

    // 渲染元素，深度可达到2层级
    async renderElements(elems: OpenXmlElement[], into?: Element): Promise<Node[]> {
        if (elems == null) {
            return null;
        }

        let result: Node[] = [];

        for (let i = 0; i < elems.length; i++) {
            let element = await this.renderElement(elems[i]);

            if (element) {
                result.push(element as Node);
            }
        }

        if (into) {
            appendChildren(into, result);
        }

        return result;
    }

    async renderContainer(elem: OpenXmlElement, tagName: keyof HTMLElementTagNameMap, props?: Record<string, any>) {
        return this.createElement(tagName, props, await this.renderChildren(elem));
    }

    async renderContainerNS(elem: OpenXmlElement, ns: string, tagName: string, props?: Record<string, any>) {
        return createElementNS(ns, tagName, props, await this.renderChildren(elem));
    }

    async renderParagraph(elem: WmlParagraph) {
        let result = this.createElement("p");

        const style = this.findStyle(elem.styleName);
        elem.tabs ??= style?.paragraphProps?.tabs;  //TODO

        this.renderClass(elem, result);
        await this.renderChildren(elem, result);
        this.renderStyleValues(elem.cssStyle, result);
        this.renderCommonProperties(result.style, elem);

        const numbering = elem.numbering ?? style?.paragraphProps?.numbering;

        if (numbering) {
            result.classList.add(this.numberingClass(numbering.id, numbering.level));
        }

        return result;
    }

    renderRunProperties(style: any, props: RunProperties) {
        this.renderCommonProperties(style, props);
    }

    renderCommonProperties(style: any, props: CommonProperties) {
        if (props == null)
            return;

        if (props.color) {
            style["color"] = props.color;
        }

        if (props.fontSize) {
            style["font-size"] = props.fontSize;
        }
    }

    async renderHyperlink(elem: WmlHyperlink) {
        let result = this.createElement("a");

        await this.renderChildren(elem, result);
        this.renderStyleValues(elem.cssStyle, result);

        if (elem.href) {
            result.href = elem.href;
        } else if (elem.id) {
            const rel = this.document.documentPart.rels
                .find(it => it.id == elem.id && it.targetMode === "External");
            result.href = rel?.target;
        }

        return result;
    }

    async renderDrawing(elem: OpenXmlElement) {
        let result = this.createElement("div");

        result.style.display = "inline-block";
        result.style.position = "relative";
        result.style.textIndent = "0px";

        await this.renderChildren(elem, result);
        this.renderStyleValues(elem.cssStyle, result);

        return result;
    }

    async renderImage(elem: IDomImage) {
        let result = this.createElement("img");

        this.renderStyleValues(elem.cssStyle, result);

        if (this.document) {
            result.src = await this.document.loadDocumentImage(elem.src, this.currentPart)
        }

        return result;
    }

    renderText(elem: WmlText) {
        return this.htmlDocument.createTextNode(elem.text);
    }

    renderDeletedText(elem: WmlText) {
        return this.options.renderEndnotes ? this.htmlDocument.createTextNode(elem.text) : null;
    }

    renderBreak(elem: WmlBreak) {
        if (elem.break == "textWrapping") {
            return this.createElement("br");
        }

        return null;
    }

    async renderInserted(elem: OpenXmlElement): Promise<Node | Node[]> {
        if (this.options.renderChanges) {
            return await this.renderContainer(elem, "ins");
        }

        return await this.renderChildren(elem);
    }

    async renderDeleted(elem: OpenXmlElement): Promise<Node> {
        if (this.options.renderChanges) {
            return await this.renderContainer(elem, "del");
        }

        return null;
    }

    renderSymbol(elem: WmlSymbol) {
        let span = this.createElement("span");
        span.style.fontFamily = elem.font;
        span.innerHTML = `&#x${elem.char};`
        return span;
    }

    renderFootnoteReference(elem: WmlNoteReference) {
        let result = this.createElement("sup");
        this.currentFootnoteIds.push(elem.id);
        result.textContent = `${this.currentFootnoteIds.length}`;
        return result;
    }

    renderEndnoteReference(elem: WmlNoteReference) {
        let result = this.createElement("sup");
        this.currentEndnoteIds.push(elem.id);
        result.textContent = `${this.currentEndnoteIds.length}`;
        return result;
    }

    // 渲染制表符
    renderTab(elem: OpenXmlElement) {
        let tabSpan = this.createElement("span");

        tabSpan.innerHTML = "&emsp;";//"&nbsp;";

        if (this.options.experimental) {
            tabSpan.className = this.tabStopClass();
            let stops = findParent<WmlParagraph>(elem, DomType.Paragraph)?.tabs;
            this.currentTabs.push({ stops, span: tabSpan });
        }

        return tabSpan;
    }

    renderBookmarkStart(elem: WmlBookmarkStart): HTMLElement {
        let result = this.createElement("span");
        result.id = elem.name;
        return result;
    }

    async renderRun(elem: WmlRun) {
        if (elem.fieldRun)
            return null;

        const result = this.createElement("span");

        if (elem.id)
            result.id = elem.id;

        this.renderClass(elem, result);
        this.renderStyleValues(elem.cssStyle, result);

        if (elem.verticalAlign) {
            const wrapper = this.createElement(elem.verticalAlign as any);
            await this.renderChildren(elem, wrapper);
            result.appendChild(wrapper);
        } else {
            await this.renderChildren(elem, result);
        }

        return result;
    }

    async renderTable(elem: WmlTable) {
        let result = this.createElement("table");

        this.tableCellPositions.push(this.currentCellPosition);
        this.tableVerticalMerges.push(this.currentVerticalMerge);
        this.currentVerticalMerge = {};
        this.currentCellPosition = { col: 0, row: 0 };

        if (elem.columns) {
            result.appendChild(this.renderTableColumns(elem.columns));
        }

        this.renderClass(elem, result);
        await this.renderChildren(elem, result);
        this.renderStyleValues(elem.cssStyle, result);

        this.currentVerticalMerge = this.tableVerticalMerges.pop();
        this.currentCellPosition = this.tableCellPositions.pop();
        return result;
    }

    renderTableColumns(columns: WmlTableColumn[]) {
        let result = this.createElement("colgroup");

        for (let col of columns) {
            let colElem = this.createElement("col");

            if (col.width)
                colElem.style.width = col.width;

            result.appendChild(colElem);
        }

        return result;
    }

    async renderTableRow(elem: OpenXmlElement) {
        let result = this.createElement("tr");

        this.currentCellPosition.col = 0;

        this.renderClass(elem, result);
        await this.renderChildren(elem, result);
        this.renderStyleValues(elem.cssStyle, result);

        this.currentCellPosition.row++;

        return result;
    }

    async renderTableCell(elem: WmlTableCell) {
        let result = this.createElement("td");

        const key = this.currentCellPosition.col;

        if (elem.verticalMerge) {
            if (elem.verticalMerge == "restart") {
                this.currentVerticalMerge[key] = result;
                result.rowSpan = 1;
            } else if (this.currentVerticalMerge[key]) {
                this.currentVerticalMerge[key].rowSpan += 1;
                result.style.display = "none";
            }
        } else {
            this.currentVerticalMerge[key] = null;
        }

        this.renderClass(elem, result);
        await this.renderChildren(elem, result);
        this.renderStyleValues(elem.cssStyle, result);

        if (elem.span)
            result.colSpan = elem.span;

        this.currentCellPosition.col += result.colSpan;

        return result;
    }

    async renderVmlPicture(elem: OpenXmlElement) {
        let result = createElement("div");
        await this.renderChildren(elem, result);
        return result;
    }

    async renderVmlElement(elem: VmlElement): Promise<SVGElement> {
        let container = createSvgElement("svg");

        container.setAttribute("style", elem.cssStyleText);

        const result = await this.renderVmlChildElement(elem);

        if (elem.imageHref?.id) {
            this.document?.loadDocumentImage(elem.imageHref.id, this.currentPart)
                .then(x => result.setAttribute("href", x));
        }

        container.appendChild(result);

        requestAnimationFrame(() => {
            const bb = (container.firstElementChild as any).getBBox();

            container.setAttribute("width", `${Math.ceil(bb.x + bb.width)}`);
            container.setAttribute("height", `${Math.ceil(bb.y + bb.height)}`);
        });

        return container;
    }

    async renderVmlChildElement(elem: VmlElement): Promise<any> {
        const result = createSvgElement(elem.tagName as any);
        Object.entries(elem.attrs).forEach(([k, v]) => result.setAttribute(k, v));

        for (let child of elem.children) {
            if (child.type == DomType.VmlElement) {
                result.appendChild(await this.renderVmlChildElement(child as VmlElement));
            } else {
                result.appendChild(...asArray(await this.renderElement(child as any)));
            }
        }

        return result;
    }

    async renderMmlRadical(elem: OpenXmlElement): Promise<HTMLElement> {
        const base = elem.children.find(el => el.type == DomType.MmlBase);

        if (elem.props?.hideDegree) {
            return createElementNS(ns.mathML, "msqrt", null, await this.renderElements([base]));
        }

        const degree = elem.children.find(el => el.type == DomType.MmlDegree);
        return createElementNS(ns.mathML, "mroot", null, await this.renderElements([base, degree]));
    }

    async renderMmlDelimiter(elem: OpenXmlElement): Promise<HTMLElement> {
        const children = [];

        children.push(createElementNS(ns.mathML, "mo", null, [elem.props.beginChar ?? '(']));
        children.push(...await this.renderElements(elem.children));
        children.push(createElementNS(ns.mathML, "mo", null, [elem.props.endChar ?? ')']));

        return createElementNS(ns.mathML, "mrow", null, children);
    }

    async renderMmlNary(elem: OpenXmlElement): Promise<HTMLElement> {
        const children = [];
        const grouped = keyBy(elem.children, x => x.type);

        const sup = grouped[DomType.MmlSuperArgument];
        const sub = grouped[DomType.MmlSubArgument];
        const supElem = sup ? createElementNS(ns.mathML, "mo", null, asArray(await this.renderElement(sup))) : null;
        const subElem = sub ? createElementNS(ns.mathML, "mo", null, asArray(await this.renderElement(sub))) : null;

        if (elem.props?.char) {
            const charElem = createElementNS(ns.mathML, "mo", null, [elem.props.char]);

            if (supElem || subElem) {
                children.push(createElementNS(ns.mathML, "munderover", null, [charElem, subElem, supElem]));
            } else if (supElem) {
                children.push(createElementNS(ns.mathML, "mover", null, [charElem, supElem]));
            } else if (subElem) {
                children.push(createElementNS(ns.mathML, "munder", null, [charElem, subElem]));
            } else {
                children.push(charElem);
            }
        }
        let base_children = await this.renderElements(grouped[DomType.MmlBase].children);
        children.push(...base_children);

        return createElementNS(ns.mathML, "mrow", null, children);
    }

    async renderMmlRun(elem: OpenXmlElement) {
        const result = createElementNS(ns.mathML, "ms");

        this.renderClass(elem, result);
        this.renderStyleValues(elem.cssStyle, result);
        await this.renderChildren(elem, result);
        return result;
    }

    async renderMllList(elem: OpenXmlElement) {
        const result = createElementNS(ns.mathML, "mtable");
        // 添加class类
        this.renderClass(elem, result);
        // 渲染style样式
        this.renderStyleValues(elem.cssStyle, result);

        const childern = await this.renderChildren(elem);

        for (let child of childern) {
            result.appendChild(createElementNS(ns.mathML, "mtr", null,
                [createElementNS(ns.mathML, "mtd", null, [child])]
            ));
        }

        return result;
    }


    renderStyleValues(style: Record<string, string>, ouput: HTMLElement) {
        for (let k in style) {
            if (k.startsWith("$")) {
                ouput.setAttribute(k.slice(1), style[k]);
            } else {
                ouput.style[k] = style[k];
            }
        }
    }

    // 添加class类名
    renderClass(input: OpenXmlElement, ouput: HTMLElement) {
        if (input.className)
            ouput.className = input.className;

        if (input.styleName) {
            ouput.classList.add(this.processStyleName(input.styleName));
        }
    }

    // 查找内置默认style样式
    findStyle(styleName: string) {
        return styleName && this.styleMap?.[styleName];
    }

    numberingClass(id: string, lvl: number) {
        return `${this.className}-num-${id}-${lvl}`;
    }

    tabStopClass() {
        return `${this.className}-tab-stop`;
    }

    styleToString(selectors: string, values: Record<string, string>, cssText: string = null) {
        let result = `${selectors} {\r\n`;

        for (const key in values) {
            if (key.startsWith('$'))
                continue;

            result += `  ${key}: ${values[key]};\r\n`;
        }

        if (cssText)
            result += cssText;

        return result + "}\r\n";
    }

    numberingCounter(id: string, lvl: number) {
        return `${this.className}-num-${id}-${lvl}`;
    }

    levelTextToContent(text: string, suff: string, id: string, numformat: string) {
        const suffMap = {
            "tab": "\\9",
            "space": "\\a0",
        };

        let result = text.replace(/%\d*/g, s => {
            let lvl = parseInt(s.substring(1), 10) - 1;
            return `"counter(${this.numberingCounter(id, lvl)}, ${numformat})"`;
        });

        return `"${result}${suffMap[suff] ?? ""}"`;
    }

    numFormatToCssValue(format: string) {
        let mapping = {
            none: "none",
            bullet: "disc",
            decimal: "decimal",
            lowerLetter: "lower-alpha",
            upperLetter: "upper-alpha",
            lowerRoman: "lower-roman",
            upperRoman: "upper-roman",
            decimalZero: "decimal-leading-zero", // 01,02,03,...
            // ordinal: "", // 1st, 2nd, 3rd,...
            // ordinalText: "", //First, Second, Third, ...
            // cardinalText: "", //One,Two Three,...
            // numberInDash: "", //-1-,-2-,-3-, ...
            // hex: "upper-hexadecimal",
            aiueo: "katakana",
            aiueoFullWidth: "katakana",
            chineseCounting: "simp-chinese-informal",
            chineseCountingThousand: "simp-chinese-informal",
            chineseLegalSimplified: "simp-chinese-formal", // 中文大写
            chosung: "hangul-consonant",
            ideographDigital: "cjk-ideographic",
            ideographTraditional: "cjk-heavenly-stem", // 十天干
            ideographLegalTraditional: "trad-chinese-formal",
            ideographZodiac: "cjk-earthly-branch", // 十二地支
            iroha: "katakana-iroha",
            irohaFullWidth: "katakana-iroha",
            japaneseCounting: "japanese-informal",
            japaneseDigitalTenThousand: "cjk-decimal",
            japaneseLegal: "japanese-formal",
            thaiNumbers: "thai",
            koreanCounting: "korean-hangul-formal",
            koreanDigital: "korean-hangul-formal",
            koreanDigital2: "korean-hanja-informal",
            hebrew1: "hebrew",
            hebrew2: "hebrew",
            hindiNumbers: "devanagari",
            ganada: "hangul",
            taiwaneseCounting: "cjk-ideographic",
            taiwaneseCountingThousand: "cjk-ideographic",
            taiwaneseDigital: "cjk-decimal",
        };

        return mapping[format] ?? format;
    }

    // 刷新tab制表符
    refreshTabStops() {
        if (!this.options.experimental) {
            return;
        }

        clearTimeout(this.tabsTimeout);

        this.tabsTimeout = setTimeout(() => {
            const pixelToPoint = computePixelToPoint();

            for (let tab of this.currentTabs) {
                updateTabStop(tab.span, tab.stops, this.defaultTabSize, pixelToPoint);
            }
        }, 500);
    }

    createElement = createElement;
}

type ChildType = Node | string;

function createElement<T extends keyof HTMLElementTagNameMap>(
    tagName: T,
    props?: Partial<Record<keyof HTMLElementTagNameMap[T], any>>,
    children?: ChildType[]
): HTMLElementTagNameMap[T] {
    return createElementNS(undefined, tagName, props, children);
}

function createSvgElement<T extends keyof SVGElementTagNameMap>(
    tagName: T,
    props?: Partial<Record<keyof SVGElementTagNameMap[T], any>>,
    children?: ChildType[]
): SVGElementTagNameMap[T] {
    return createElementNS(ns.svg, tagName, props, children);
}

function createElementNS(ns: string, tagName: string, props?: Partial<Record<any, any>>, children?: ChildType[]): any {
    let result = ns ? document.createElementNS(ns, tagName) : document.createElement(tagName);
    Object.assign(result, props);
    children && appendChildren(result, children);
    return result;
}

function removeAllElements(elem: HTMLElement) {
    elem.innerHTML = '';
}

// 插入子元素
function appendChildren(parent: Element | DocumentFragment, children: (Node | string)[]) {
    children.forEach(child => {
        parent.appendChild(isString(child) ? document.createTextNode(child) : child)
    });
}

// 创建style标签
function createStyleElement(cssText: string) {
    return createElement("style", { innerHTML: cssText });
}

// 插入注释
function appendComment(elem: HTMLElement | DocumentFragment, comment: string) {
    elem.appendChild(document.createComment(comment));
}

function findParent<T extends OpenXmlElement>(elem: OpenXmlElement, type: DomType): T {
    let parent = elem.parent;

    while (parent != null && parent.type != type)
        parent = parent.parent;

    return <T>parent;
}
