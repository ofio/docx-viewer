import {WordDocument} from './word-document';
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
import {CommonProperties} from './document/common';
import {Options} from './docx-preview';
import {DocumentElement} from './document/document';
import {WmlParagraph} from './document/paragraph';
import {asArray, escapeClassName, isString, keyBy, mergeDeep} from './utils';
import {computePixelToPoint, updateTabStop} from './javascript';
import {FontTablePart} from './font-table/font-table';
import {FooterHeaderReference, SectionProperties, Section} from './document/section';
import {WmlRun, RunProperties} from './document/run';
import {WmlBookmarkStart} from './document/bookmarks';
import {IDomStyle} from './document/style';
import {WmlBaseNote, WmlFootnote} from './notes/elements';
import {ThemePart} from './theme/theme-part';
import {BaseHeaderFooterPart} from './header-footer/parts';
import {Part} from './common/part';
import {VmlElement} from './vml/vml';

const ns = {
    html: "http://www.w3.org/1999/xhtml",
    svg: "http://www.w3.org/2000/svg",
    mathML: "http://www.w3.org/1998/Math/MathML"
}

interface CellPos {
    col: number;
    row: number;
}

type CellVerticalMergeType = Record<number, HTMLTableCellElement>;

interface CurrentElement {
    parent: HTMLElement | Element,
}

// HTML渲染器

export class HtmlRendererSync {

    className: string = "docx";
    rootSelector: string;
    document: WordDocument;
    options: Options;
    styleMap: Record<string, IDomStyle> = {};
    currentPart: Part = null;
    wrapper: HTMLElement;

    // 当前操作的section
    current_section: Section;
    // 当前渲染的元素
    current_element: CurrentElement = {parent: null};

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

    /**
     * Object对象 => HTML标签
     *
     * @param document word文档Object对象
     * @param bodyContainer HTML生成容器
     * @param styleContainer CSS样式生成容器
     * @param options 渲染配置选项
     */

    render(document: WordDocument, bodyContainer: HTMLElement, styleContainer: HTMLElement = null, options: Options) {
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
        appendComment(styleContainer, "docxjs library predefined styles");
        // 添加默认CSS样式
        styleContainer.appendChild(this.renderDefaultStyle());

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
        // 根据option生成wrapper
        if (this.options.inWrapper) {
            this.wrapper = this.renderWrapper();
            bodyContainer.appendChild(this.wrapper);
        } else {
            this.wrapper = bodyContainer;
        }
        // 主文档--内容
        this.renderSections(document.documentPart.body);

        // 刷新制表符
        this.refreshTabStops();
    }

    // 渲染默认样式
    renderDefaultStyle() {
        let c = this.className;
        let styleText = `
			.${c}-wrapper { background: gray; padding: 30px; padding-bottom: 0px; display: flex; flex-flow: column; align-items: center; } 
			.${c}-wrapper>section.${c} { background: white; box-shadow: 0 0 10px rgba(0, 0, 0, 0.5); margin-bottom: 30px; }
			.${c} { color: black; hyphens: auto; }
			section.${c} { box-sizing: border-box; display: flex; flex-flow: column nowrap; position: relative; overflow: hidden; }
            section.${c}>header { position: absolute; top: 0; z-index: 1; display: flex; align-items: flex-end; }
			section.${c}>article { overflow: hidden; z-index: 1; }
			section.${c}>footer { position: absolute; bottom: 0; z-index: 1; }
			.${c} table { border-collapse: collapse; }
			.${c} table td, .${c} table th { vertical-align: top; }
			.${c} p { margin: 0pt; min-height: 1em; }
			.${c} span { white-space: pre-wrap; overflow-wrap: break-word; }
			.${c} a { color: inherit; text-decoration: inherit; }
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
            for (let [k, v] of Object.entries(colorScheme.colors)) {
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
                        style.styles.push({...baseValues, values: {...baseValues.values}});
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

    renderStyles(styles: IDomStyle[]): HTMLElement {
        let styleText = "";
        const stylesMap = this.styleMap;
        const defaultStyles = keyBy(styles.filter(s => s.isDefault), s => s.target);

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

                if (defaultStyles[style.target] == style)
                    selector = `.${this.className} ${style.target}, ` + selector;

                styleText += this.styleToString(selector, subStyle.values);
            }
        }

        return createStyleElement(styleText);
    }

    processNumberings(numberings: IDomNumbering[]) {
        for (let num of numberings.filter(n => n.pStyleName)) {
            const style = this.findStyle(num.pStyleName);

            if (style?.paragraphProps?.numbering) {
                style.paragraphProps.numbering.level = num.level;
            }
        }
    }

    renderNumbering(numberings: IDomNumbering[], styleContainer: HTMLElement) {
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

    numberingClass(id: string, lvl: number) {
        return `${this.className}-num-${id}-${lvl}`;
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

    // renderNumbering2(numberingPart: NumberingPartProperties, container: HTMLElement): HTMLElement {
    //     let css = "";
    //     const numberingMap = keyBy(numberingPart.abstractNumberings, x => x.id);
    //     const bulletMap = keyBy(numberingPart.bulletPictures, x => x.id);
    //     const topCounters = [];
    //
    //     for(let num of numberingPart.numberings) {
    //         const absNum = numberingMap[num.abstractId];
    //
    //         for(let lvl of absNum.levels) {
    //             const className = this.numberingClass(num.id, lvl.level);
    //             let listStyleType = "none";
    //
    //             if(lvl.text && lvl.format == 'decimal') {
    //                 const counter = this.numberingCounter(num.id, lvl.level);
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

    // 生成父级容器
    renderWrapper() {
        return createElement("div", {className: `${this.className}-wrapper`});
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

    // 处理表格style样式
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

    // 根据section切分页面
    splitBySection(elements: OpenXmlElement[]): Section[] {
        // 当前操作section，elements数组包含子元素
        let current_section = {sectProps: null, elements: [], is_split: false,};
        // 切分出的所有sections
        let sections = [current_section];

        for (let elem of elements) {
            // 添加elem进入当前操作section
            current_section.elements.push(elem);
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
                    // 标记当前section已拆分
                    current_section.is_split = true;
                    // 保存当前section的sectionProps
                    current_section.sectProps = sectProps;
                    // 重置新的section
                    current_section = {sectProps: null, elements: [], is_split: false};
                    // 添加新section
                    sections.push(current_section);
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
                            // 分页符、换行符、分栏符
                            if (t.type != DomType.Break) {
                                return false;
                            }
                            // 默认忽略lastRenderedPageBreak，
                            if ((t as WmlBreak).break == "lastRenderedPageBreak") {
                                // 判断前一个p段落，
                                // 如果含有分页符、分节符，那它们一定位于上一个section，
                                // 如果前一个段落是普通段落，则代表文字过多超过一页，需要自动分页
                                return current_section.elements.length > 1 || !this.options.ignoreLastRenderedPageBreak;
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
                // 段落中存在节属性sectProps
                if (sectProps) {
                    // 标记当前section未拆分，需要计算拆分
                    current_section.is_split = false;
                }
                // 段落Break索引
                if (pBreakIndex != -1) {
                    // 标记当前section 已拆分
                    current_section.is_split = true;
                }
                // 段落中存在节属性sectProps/段落Break索引
                if (sectProps || pBreakIndex != -1) {
                    // 保存当前section的sectionProps
                    current_section.sectProps = sectProps;
                    // 重置新的section
                    current_section = {sectProps: null, elements: [], is_split: false};
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
                        let new_paragraph = {...p, children: origin_run.slice(pBreakIndex)};
                        // 保存Break索引前面的Run
                        p.children = origin_run.slice(0, pBreakIndex);
                        // 添加新段落
                        current_section.elements.push(new_paragraph);

                        if (is_split) {
                            // Run下面原始的元素
                            let origin_elements = breakRun.children;
                            // 切出Run Break索引前面的元素，创建新Run
                            let newRun = {...breakRun, children: origin_elements.slice(0, rBreakIndex)};
                            // 将新Run放入上一个section的段落
                            p.children.push(newRun);
                            // 切出Run Break索引后面的元素
                            breakRun.children = origin_elements.slice(rBreakIndex);
                        }
                    }
                }
            }

            // elem元素是表格，需要渲染过程中拆分section，标记:is_split
            if (elem.type === DomType.Table) {
                current_section.is_split = false;
            }

        }

        // 一个节可能分好几个页，但是节属性section_props存在当前节中最后一段对应的 paragraph 元素的子元素。即：[null,null,null,setPr];
        let currentSectProps = null;
        // 倒序给每一页填充section_props，方便后期页面渲染
        for (let i = sections.length - 1; i >= 0; i--) {
            if (sections[i].sectProps == null) {
                sections[i].sectProps = currentSectProps;
            } else {
                currentSectProps = sections[i].sectProps
            }
        }
        return sections;
    }

    // 生成所有的Page Section
    renderSections(document: DocumentElement) {
        // 生成页面parent父级关系
        this.processElement(document);
        // 根据options.breakPages，选择是否分页
        let sections: Section[];
        if (this.options.breakPages) {
            // 根据section切分页面
            sections = this.splitBySection(document.children);
        } else {
            // 不分页则，只有一个section
            sections = [{sectProps: document.props, elements: document.children, is_split: false}];
        }
        // 前一个节属性，判断分节符的第一个section
        let prevProps = null;
        // 遍历生成每一个section
        for (let i = 0, l = sections.length; i < l; i++) {
            this.currentFootnoteIds = [];

            let section: Section = sections[i];

            let {sectProps} = section;
            // section属性不存在，则使用文档级别props;
            section.sectProps = sectProps ?? document.props;
            // 是否第一个section
            section.isFirstSection = prevProps != sectProps;
            // 是否最后一个section
            section.isLastSection = i === (l - 1);
            // 页码，判断奇偶页码
            section.pageIndex = i;
            // 溢出检测默认不开启
            section.checking_overflow = false;
            // 将上述数据存储在current_section中
            this.current_section = section;
            // 渲染单个section
            this.renderSection();
            // 存储前一个节属性
            prevProps = sectProps;
        }
    }

    // 生成单个section,如果发现超出一页，递归拆分出下一个section
    renderSection() {
        // 当前操作的section
        let section: Section = this.current_section;
        // 解构section中的属性
        let {sectProps, isFirstSection, isLastSection, pageIndex} = section;
        // 根据sectProps，创建section
        const sectionElement = this.createSection(this.className, sectProps);
        // 给section添加style样式
        this.renderStyleValues(this.document.documentPart.body.cssStyle, sectionElement);
        // 渲染section页眉
        if (this.options.renderHeaders) {
            this.renderHeaderFooterRef(sectProps.headerRefs, sectProps, pageIndex, isFirstSection, sectionElement);
        }
        // 渲染section脚注
        if (this.options.renderFootnotes) {
            this.renderNotes(this.currentFootnoteIds, this.footnoteMap, sectionElement);
        }
        // 渲染section尾注，判断最后一页
        if (this.options.renderEndnotes && isLastSection) {
            this.renderNotes(this.currentEndnoteIds, this.endnoteMap, sectionElement);
        }
        // 渲染section页脚
        if (this.options.renderFooters) {
            this.renderHeaderFooterRef(sectProps.footerRefs, sectProps, pageIndex, isFirstSection, sectionElement);
        }
        // section内容Article元素
        let contentElement = createElement("article");
        // 根据options.breakPages，设置article的高度
        if (this.options.breakPages) {
            // 切分页面，高度固定
            contentElement.style.height = sectProps.contentSize.height;
        } else {
            // 不分页则，拥有最小高度
            contentElement.style.minHeight = sectProps.contentSize.height;
        }
        // 缓存当前操作的Article元素
        this.current_section.contentElement = contentElement;
        // 将Article插入section
        sectionElement.appendChild(contentElement);
        // 标识--开启溢出计算
        this.current_section.checking_overflow = true;
        // 生成article内容
        this.renderElements(section.elements, contentElement);
        // 标识--结束溢出计算
        this.current_section.checking_overflow = false;
    }

    // 创建Page Section
    createSection(className: string, props: SectionProperties) {
        let elem = createElement("section", {className});

        if (props) {
            if (props.pageMargins) {
                elem.style.paddingLeft = props.pageMargins.left;
                elem.style.paddingRight = props.pageMargins.right;
                elem.style.paddingTop = props.pageMargins.top;
                elem.style.paddingBottom = props.pageMargins.bottom;
            }

            if (props.pageSize) {
                if (!this.options.ignoreWidth) {
                    elem.style.width = props.pageSize.width;
                }
                if (!this.options.ignoreHeight) {
                    elem.style.minHeight = props.pageSize.height;
                }
            }

            if (props.columns && props.columns.numberOfColumns) {
                elem.style.columnCount = `${props.columns.numberOfColumns}`;
                elem.style.columnGap = props.columns.space;

                if (props.columns.separator) {
                    elem.style.columnRule = "1px solid black";
                }
            }
        }
        // 插入生成的section
        this.wrapper.appendChild(elem);

        return elem;
    }

    // TODO 分页不准确，页脚页码混乱
    // 渲染页眉/页脚的Ref
    renderHeaderFooterRef(refs: FooterHeaderReference[], props: SectionProperties, page: number, firstOfSection: boolean, parent: HTMLElement) {
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
            // 根据页眉页脚，设置CSS
            switch (part.rootElement.type) {
                case DomType.Header:
                    part.rootElement.cssStyle = {
                        left: props.pageMargins?.left,
                        width: props.contentSize?.width,
                        height: props.pageMargins?.top,
                    }
                    break;
                case DomType.Footer:
                    part.rootElement.cssStyle = {
                        left: props.pageMargins?.left,
                        width: props.contentSize?.width,
                        height: props.pageMargins?.bottom,
                    }
                    break;
                default:
                    console.warn('set header/footer style error', part.rootElement.type);
                    break;
            }

            this.renderElements([part.rootElement], parent);
            this.currentPart = null;
        }
    }

    // 渲染脚注/尾注
    renderNotes(noteIds: string[], notesMap: Record<string, WmlBaseNote>, parent: HTMLElement) {
        let notes = noteIds.map(id => notesMap[id]).filter(x => x);

        if (notes.length > 0) {
            let oList = createElement("ol", null);
            this.renderElements(notes, oList);
            parent.appendChild(oList);
        }
    }

    // 根据XML对象渲染出多元素，
    renderElements(elems: OpenXmlElement[], parent: HTMLElement | Element) {

        for (let i = 0; i < elems.length; i++) {
            this.renderElement(elems[i], parent);
            // 缓存当前操作元素的索引值
            this.current_section.elementIndex = i;
        }

    }

    // 根据XML对象渲染单个元素
    renderElement(elem: OpenXmlElement, parent?: HTMLElement | Element): Node | Node[] {
        let oNode: Node | Node[];

        switch (elem.type) {
            case DomType.Paragraph:
                oNode = this.renderParagraph(elem as WmlParagraph, parent);
                break;
            case DomType.BookmarkStart:
                oNode = this.renderBookmarkStart(elem as WmlBookmarkStart, parent);
                break;
            case DomType.BookmarkEnd:
                oNode = null; //ignore bookmark end
                break;
            case DomType.Run:
                oNode = this.renderRun(elem as WmlRun, parent);
                break;
            case DomType.Table:
                oNode = this.renderTable(elem, parent);
                break;
            case DomType.Row:
                oNode = this.renderTableRow(elem, parent);
                break;
            case DomType.Cell:
                oNode = this.renderTableCell(elem, parent);
                break;
            case DomType.Hyperlink:
                oNode = this.renderHyperlink(elem, parent);
                break;
            case DomType.Drawing:
                oNode = this.renderDrawing(elem, parent);
                break;
            case DomType.Image:
                oNode = this.renderImage(elem as IDomImage, parent);
                break;
            case DomType.Text:
                oNode = this.renderText(elem as WmlText, parent);
                break;
            case DomType.DeletedText:
                oNode = this.renderDeletedText(elem as WmlText, parent);
                break;
            case DomType.Tab:
                oNode = this.renderTab(elem, parent);
                break;
            case DomType.Symbol:
                oNode = this.renderSymbol(elem as WmlSymbol, parent);
                break;
            case DomType.Break:
                oNode = this.renderBreak(elem as WmlBreak, parent);
                break;
            case DomType.Footer:
                oNode = this.renderHeaderFooter(elem, "footer");
                if (parent) {
                    this.appendChildren(parent, oNode);
                }
                break;
            case DomType.Header:
                oNode = this.renderHeaderFooter(elem, "header");
                if (parent) {
                    this.appendChildren(parent, oNode);
                }
                break;
            case DomType.Footnote:
            case DomType.Endnote:
                oNode = this.renderContainer(elem, "li");
                if (parent) {
                    this.appendChildren(parent, oNode);
                }
                break;
            case DomType.FootnoteReference:
                oNode = this.renderFootnoteReference(elem as WmlNoteReference);
                if (parent) {
                    this.appendChildren(parent, oNode);
                }
                break;
            case DomType.EndnoteReference:
                oNode = this.renderEndnoteReference(elem as WmlNoteReference);
                if (parent) {
                    this.appendChildren(parent, oNode);
                }
                break;
            case DomType.NoBreakHyphen:
                oNode = createElement("wbr");
                if (parent) {
                    this.appendChildren(parent, oNode);
                }
                break;
            case DomType.VmlPicture:
                oNode = this.renderVmlPicture(elem);
                if (parent) {
                    this.appendChildren(parent, oNode);
                }
                break;
            case DomType.VmlElement:
                oNode = this.renderVmlElement(elem as VmlElement);
                if (parent) {
                    this.appendChildren(parent, oNode);
                }
                break;
            case DomType.MmlMath:
                oNode = this.renderContainerNS(elem, ns.mathML, "math", {xmlns: ns.mathML});
                if (parent) {
                    this.appendChildren(parent, oNode);
                }
                break;
            case DomType.MmlMathParagraph:
                oNode = this.renderContainer(elem, "span");
                if (parent) {
                    this.appendChildren(parent, oNode);
                }
                break;
            case DomType.MmlFraction:
                oNode = this.renderContainerNS(elem, ns.mathML, "mfrac");
                if (parent) {
                    this.appendChildren(parent, oNode);
                }
                break;
            case DomType.MmlBase:
                oNode = this.renderContainerNS(elem, ns.mathML, elem.parent.type == DomType.MmlMatrixRow ? "mtd" : "mrow");
                if (parent) {
                    this.appendChildren(parent, oNode);
                }
                break;
            case DomType.MmlNumerator:
            case DomType.MmlDenominator:
            case DomType.MmlFunction:
            case DomType.MmlLimit:
            case DomType.MmlBox:
                oNode = this.renderContainerNS(elem, ns.mathML, "mrow");
                if (parent) {
                    this.appendChildren(parent, oNode);
                }
                break;
            case DomType.MmlGroupChar:
                oNode = this.renderMmlGroupChar(elem);
                if (parent) {
                    this.appendChildren(parent, oNode);
                }
                break;
            case DomType.MmlLimitLower:
                oNode = this.renderContainerNS(elem, ns.mathML, "munder");
                if (parent) {
                    this.appendChildren(parent, oNode);
                }
                break;
            case DomType.MmlMatrix:
                oNode = this.renderContainerNS(elem, ns.mathML, "mtable");
                if (parent) {
                    this.appendChildren(parent, oNode);
                }
                break;
            case DomType.MmlMatrixRow:
                oNode = this.renderContainerNS(elem, ns.mathML, "mtr");
                if (parent) {
                    this.appendChildren(parent, oNode);
                }
                break;
            case DomType.MmlRadical:
                oNode = this.renderMmlRadical(elem);
                if (parent) {
                    this.appendChildren(parent, oNode);
                }
                break;
            case DomType.MmlSuperscript:
                oNode = this.renderContainerNS(elem, ns.mathML, "msup");
                if (parent) {
                    this.appendChildren(parent, oNode);
                }
                break;
            case DomType.MmlSubscript:
                oNode = this.renderContainerNS(elem, ns.mathML, "msub");
                if (parent) {
                    this.appendChildren(parent, oNode);
                }
                break;
            case DomType.MmlDegree:
            case DomType.MmlSuperArgument:
            case DomType.MmlSubArgument:
                oNode = this.renderContainerNS(elem, ns.mathML, "mn");
                if (parent) {
                    this.appendChildren(parent, oNode);
                }
                break;
            case DomType.MmlFunctionName:
                oNode = this.renderContainerNS(elem, ns.mathML, "ms");
                if (parent) {
                    this.appendChildren(parent, oNode);
                }
                break;
            case DomType.MmlDelimiter:
                oNode = this.renderMmlDelimiter(elem);
                if (parent) {
                    this.appendChildren(parent, oNode);
                }
                break;
            case DomType.MmlRun:
                oNode = this.renderMmlRun(elem);
                if (parent) {
                    this.appendChildren(parent, oNode);
                }
                break;
            case DomType.MmlNary:
                oNode = this.renderMmlNary(elem);
                if (parent) {
                    this.appendChildren(parent, oNode);
                }
                break;
            case DomType.MmlPreSubSuper:
                oNode = this.renderMmlPreSubSuper(elem);
                if (parent) {
                    this.appendChildren(parent, oNode);
                }
                break;
            case DomType.MmlBar:
                oNode = this.renderMmlBar(elem);
                if (parent) {
                    this.appendChildren(parent, oNode);
                }
                break;
            case DomType.MmlEquationArray:
                oNode = this.renderMllList(elem);
                if (parent) {
                    this.appendChildren(parent, oNode);
                }
                break;
            case DomType.Inserted:
                oNode = this.renderInserted(elem);
                if (parent) {
                    this.appendChildren(parent, oNode);
                }
                break;
            case DomType.Deleted:
                oNode = this.renderDeleted(elem);
                if (parent) {
                    this.appendChildren(parent, oNode);
                }
                break;
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
        if ((elem as WmlBreak).break == "lastRenderedPageBreak") {
            return !this.options.ignoreLastRenderedPageBreak;
        }
        // 分页符
        if ((elem as WmlBreak).break === "page") {
            return true;
        }
    }

    // 根据XML对象渲染子元素，并插入父级元素
    renderChildren(elem: OpenXmlElement, parent: Element) {
        this.renderElements(elem.children, parent);
    }

    // 插入子元素，针对后代元素进行溢出检测
    appendChildren(parent: HTMLElement | Element, children: ChildrenType) {
        // TODO 单个child执行溢出检测，children数组应该依次进行溢出检测
        appendChildren(parent, children);
        let {is_split, contentElement, pageIndex, elementIndex, checking_overflow, elements} = this.current_section;
        // 针对一级元素进行溢出检测
        if (is_split === false && checking_overflow) {
            let is_overflow = checkOverflow(contentElement);
            if (is_overflow) {
                console.log(elementIndex, children, is_overflow);
                // 删除DOM中导致溢出的元素
                removeElements(children, parent);
                // 删除数组前面已经渲染的元素，保留后续为渲染元素
                elements.splice(0, elementIndex);
                // 页码自增+1
                pageIndex += 1;
                // 关闭溢出检测，方便后续页脚渲染
                checking_overflow = false;
                // 覆盖current_section的属性
                this.current_section = {...this.current_section, pageIndex, checking_overflow, elements};
                // 重启新一个section的渲染
                this.renderSection();
            }
        }
    }

    renderContainer(elem: OpenXmlElement, tagName: keyof HTMLElementTagNameMap, props?: Record<string, any>) {
        let parent = createElement(tagName, props);
        this.renderChildren(elem, parent);
        return parent;
    }

    renderContainerNS(elem: OpenXmlElement, ns: string, tagName: string, props?: Record<string, any>) {
        let parent = createElementNS(ns, tagName as any, props);
        this.renderChildren(elem, parent);
        return parent;
    }

    renderParagraph(elem: WmlParagraph, parent?: HTMLElement | Element) {
        let oParagraph = createElement("p");

        const style = this.findStyle(elem.styleName);
        elem.tabs ??= style?.paragraphProps?.tabs;  //TODO

        this.renderClass(elem, oParagraph);
        this.renderStyleValues(elem.cssStyle, oParagraph);
        this.renderCommonProperties(oParagraph.style, elem);

        // 插入子元素,可针对后代元素进行溢出检测
        if (parent) {
            this.appendChildren(parent, oParagraph);
        }

        this.renderChildren(elem, oParagraph);


        const numbering = elem.numbering ?? style?.paragraphProps?.numbering;

        if (numbering) {
            oParagraph.classList.add(this.numberingClass(numbering.id, numbering.level));
        }

        return oParagraph;
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

    renderHyperlink(elem: WmlHyperlink, parent?: HTMLElement | Element) {
        let oAnchor = createElement("a");

        // 插入子元素,可针对后代元素进行溢出检测
        if (parent) {
            this.appendChildren(parent, oAnchor);
        }

        this.renderChildren(elem, oAnchor);
        this.renderStyleValues(elem.cssStyle, oAnchor);

        if (elem.href) {
            oAnchor.href = elem.href;
        } else if (elem.id) {
            const rel = this.document.documentPart.rels
                .find(it => it.id == elem.id && it.targetMode === "External");
            oAnchor.href = rel?.target;
        }

        return oAnchor;
    }

    renderDrawing(elem: OpenXmlElement, parent?: HTMLElement | Element) {
        let oDrawing = createElement("div");

        oDrawing.style.display = "inline-block";
        oDrawing.style.position = "relative";
        oDrawing.style.textIndent = "0px";

        this.renderChildren(elem, oDrawing);
        this.renderStyleValues(elem.cssStyle, oDrawing);

        // 插入子元素,可针对后代元素进行溢出检测
        if (parent) {
            this.appendChildren(parent, oDrawing);
        }

        return oDrawing;
    }

    // 渲染图片，默认转换blob--异步
    renderImage(elem: IDomImage, parent?: HTMLElement | Element) {
        let oImage = createElement("img");

        this.renderStyleValues(elem.cssStyle, oImage);

        if (this.document) {
            this.document
                .loadDocumentImage(elem.src, this.currentPart)
                .then(src => {
                    oImage.src = src;
                });
        }

        // 插入子元素,可针对后代元素进行溢出检测
        if (parent) {
            this.appendChildren(parent, oImage);
        }

        return oImage;
    }

    renderText(elem: WmlText, parent?: HTMLElement | Element) {
        let oText = document.createTextNode(elem.text);
        // 插入子元素,可针对后代元素进行溢出检测
        if (parent) {
            this.appendChildren(parent, oText);
        }
        return oText;
    }

    renderDeletedText(elem: WmlText, parent?: HTMLElement | Element) {
        let oDeletedText: Text;
        if (this.options.renderEndnotes) {
            oDeletedText = document.createTextNode(elem.text);
            // 插入子元素,可针对后代元素进行溢出检测
            if (parent) {
                this.appendChildren(parent, oDeletedText);
            }
        } else {
            oDeletedText = null;
        }
        return oDeletedText;
    }

    renderBreak(elem: WmlBreak, parent?: HTMLElement | Element) {
        if (elem.break == "textWrapping") {
            let oBr = createElement("br");
            // 插入子元素,可针对后代元素进行溢出检测
            if (parent) {
                this.appendChildren(parent, oBr);
            }
            return oBr;
        }

        return null;
    }

    renderInserted(elem: OpenXmlElement) {
        if (this.options.renderChanges) {
            return this.renderContainer(elem, "ins");
        }

        return this.renderContainer(elem, "span");
    }

    renderDeleted(elem: OpenXmlElement) {
        if (this.options.renderChanges) {
            return this.renderContainer(elem, "del");
        }

        return null;
    }

    renderSymbol(elem: WmlSymbol, parent?: HTMLElement | Element) {
        let oSpan = createElement("span");
        oSpan.style.fontFamily = elem.font;
        oSpan.innerHTML = `&#x${elem.char};`

        // 插入子元素,可针对后代元素进行溢出检测
        if (parent) {
            this.appendChildren(parent, oSpan);
        }

        return oSpan;
    }

    // 渲染页眉页脚
    renderHeaderFooter(elem: OpenXmlElement, tagName: keyof HTMLElementTagNameMap,) {
        let oElement: HTMLElement = createElement(tagName);
        // 渲染子元素
        this.renderChildren(elem, oElement);
        // 渲染style样式
        this.renderStyleValues(elem.cssStyle, oElement);

        return oElement;
    }

    renderFootnoteReference(elem: WmlNoteReference) {
        let oSup = createElement("sup");
        this.currentFootnoteIds.push(elem.id);
        oSup.textContent = `${this.currentFootnoteIds.length}`;
        return oSup;
    }

    renderEndnoteReference(elem: WmlNoteReference) {
        let oSup = createElement("sup");
        this.currentEndnoteIds.push(elem.id);
        oSup.textContent = `${this.currentEndnoteIds.length}`;
        return oSup;
    }

    // 渲染制表符
    renderTab(elem: OpenXmlElement, parent?: HTMLElement | Element) {
        let tabSpan = createElement("span");

        tabSpan.innerHTML = "&emsp;";//"&nbsp;";

        if (this.options.experimental) {
            tabSpan.className = this.tabStopClass();
            let stops = findParent<WmlParagraph>(elem, DomType.Paragraph)?.tabs;
            this.currentTabs.push({stops, span: tabSpan});
        }

        // 插入子元素,可针对后代元素进行溢出检测
        if (parent) {
            this.appendChildren(parent, tabSpan);
        }

        return tabSpan;
    }

    renderBookmarkStart(elem: WmlBookmarkStart, parent?: HTMLElement | Element): HTMLElement {
        let oSpan = createElement("span");
        oSpan.id = elem.name;

        // 插入子元素,可针对后代元素进行溢出检测
        if (parent) {
            this.appendChildren(parent, oSpan);
        }

        return oSpan;
    }

    renderRun(elem: WmlRun, parent?: HTMLElement | Element) {
        if (elem.fieldRun) {
            return null;
        }

        const oSpan = createElement("span");

        if (elem.id) {
            oSpan.id = elem.id;
        }

        this.renderClass(elem, oSpan);
        this.renderStyleValues(elem.cssStyle, oSpan);

        // 插入子元素,可针对后代元素进行溢出检测
        if (parent) {
            this.appendChildren(parent, oSpan);
        }

        if (elem.verticalAlign) {
            const wrapper = createElement(elem.verticalAlign as any);
            this.renderChildren(elem, wrapper);
            this.appendChildren(oSpan, wrapper);
        } else {
            this.renderChildren(elem, oSpan);
        }
        return oSpan;
    }

    renderTable(elem: WmlTable, parent?: HTMLElement | Element) {
        let oTable = createElement("table");

        this.tableCellPositions.push(this.currentCellPosition);
        this.tableVerticalMerges.push(this.currentVerticalMerge);
        this.currentVerticalMerge = {};
        this.currentCellPosition = {col: 0, row: 0};

        this.renderClass(elem, oTable);
        this.renderStyleValues(elem.cssStyle, oTable);
        // 插入子元素,可针对后代元素进行溢出检测
        if (parent) {
            this.appendChildren(parent, oTable);
        }

        // 渲染表格column列
        if (elem.columns) {
            let oColumns = this.renderTableColumns(elem.columns, oTable);
        }

        this.renderChildren(elem, oTable);

        this.currentVerticalMerge = this.tableVerticalMerges.pop();
        this.currentCellPosition = this.tableCellPositions.pop();
        return oTable;
    }

    renderTableColumns(columns: WmlTableColumn[], parent?: HTMLElement | Element) {
        let oColGroup = createElement("colgroup");

        // 插入子元素,可针对后代元素进行溢出检测
        if (parent) {
            this.appendChildren(parent, oColGroup);
        }

        for (let col of columns) {
            let oCol = createElement("col");

            if (col.width) {
                oCol.style.width = col.width;
            }
            this.appendChildren(oColGroup, oCol);
        }

        return oColGroup;
    }

    renderTableRow(elem: OpenXmlElement, parent?: HTMLElement | Element) {
        let oTableRow = createElement("tr");

        this.currentCellPosition.col = 0;

        // 插入子元素,可针对后代元素进行溢出检测
        if (parent) {
            this.appendChildren(parent, oTableRow);
        }

        this.renderClass(elem, oTableRow);
        this.renderChildren(elem, oTableRow);
        this.renderStyleValues(elem.cssStyle, oTableRow);

        this.currentCellPosition.row++;

        return oTableRow;
    }

    renderTableCell(elem: WmlTableCell, parent?: HTMLElement | Element) {
        let oTableCell = createElement("td");

        const key = this.currentCellPosition.col;

        if (elem.verticalMerge) {
            if (elem.verticalMerge == "restart") {
                this.currentVerticalMerge[key] = oTableCell;
                oTableCell.rowSpan = 1;
            } else if (this.currentVerticalMerge[key]) {
                this.currentVerticalMerge[key].rowSpan += 1;
                oTableCell.style.display = "none";
            }
        } else {
            this.currentVerticalMerge[key] = null;
        }

        // 插入子元素,可针对后代元素进行溢出检测
        if (parent) {
            this.appendChildren(parent, oTableCell);
        }

        this.renderClass(elem, oTableCell);
        this.renderChildren(elem, oTableCell);
        this.renderStyleValues(elem.cssStyle, oTableCell);

        if (elem.span)
            oTableCell.colSpan = elem.span;

        this.currentCellPosition.col += oTableCell.colSpan;

        return oTableCell;
    }

    renderVmlPicture(elem: OpenXmlElement) {
        let oPictureContainer = createElement("div");
        this.renderChildren(elem, oPictureContainer);
        return oPictureContainer;
    }

    renderVmlElement(elem: VmlElement): SVGElement {
        let oSvg = createSvgElement("svg");

        oSvg.setAttribute("style", elem.cssStyleText);

        const oChildren = this.renderVmlChildElement(elem);

        if (elem.imageHref?.id) {
            this.document?.loadDocumentImage(elem.imageHref.id, this.currentPart)
                .then(x => oChildren.setAttribute("href", x));
        }

        appendChildren(oSvg, oChildren);

        requestAnimationFrame(() => {
            const bb = (oSvg.firstElementChild as any).getBBox();

            oSvg.setAttribute("width", `${Math.ceil(bb.x + bb.width)}`);
            oSvg.setAttribute("height", `${Math.ceil(bb.y + bb.height)}`);
        });

        return oSvg;
    }

    renderVmlChildElement(elem: VmlElement) {
        const oSvgElement = createSvgElement(elem.tagName as any);
        // set attributes
        Object.entries(elem.attrs).forEach(([k, v]) => oSvgElement.setAttribute(k, v));

        for (let child of elem.children) {
            if (child.type == DomType.VmlElement) {
                let oChild = this.renderVmlChildElement(child as VmlElement);
                appendChildren(oSvgElement, oChild);
            } else {
                let oChild = this.renderElement(child as any);
                appendChildren(oSvgElement, oChild);
            }
        }

        return oSvgElement;
    }

    renderMmlRadical(elem: OpenXmlElement): HTMLElement | Element {
        const base = elem.children.find(el => el.type == DomType.MmlBase);
        let oParent: HTMLElement | Element;
        if (elem.props?.hideDegree) {
            oParent = createElementNS(ns.mathML, "msqrt", null);
            this.renderElements([base], oParent);
            return oParent;
        }

        const degree = elem.children.find(el => el.type == DomType.MmlDegree);
        oParent = createElementNS(ns.mathML, "mroot", null);
        this.renderElements([base, degree], oParent);
        return oParent;
    }

    renderMmlDelimiter(elem: OpenXmlElement): MathMLElement {
        let oMrow = createElementNS(ns.mathML, "mrow", null);
        // 开始Char
        let oBegin = createElementNS(ns.mathML, "mo", null, [elem.props.beginChar ?? '(']);
        appendChildren(oMrow, oBegin);
        // 子元素
        this.renderElements(elem.children, oMrow);
        // 结束char
        let oEnd = createElementNS(ns.mathML, "mo", null, [elem.props.endChar ?? ')']);
        appendChildren(oMrow, oEnd);

        return oMrow;
    }

    renderMmlNary(elem: OpenXmlElement): MathMLElement {
        const children = [];
        const grouped = keyBy(elem.children, x => x.type);

        const sup = grouped[DomType.MmlSuperArgument];
        const sub = grouped[DomType.MmlSubArgument];

        const supElem = sup ? createElementNS(ns.mathML, "mo", null, asArray(this.renderElement(sup))) : null;
        const subElem = sub ? createElementNS(ns.mathML, "mo", null, asArray(this.renderElement(sub))) : null;

        const charElem = createElementNS(ns.mathML, "mo", null, [elem.props?.char ?? '\u222B']);

        if (supElem || subElem) {
            children.push(createElementNS(ns.mathML, "munderover", null, [charElem, subElem, supElem]));
        } else if (supElem) {
            children.push(createElementNS(ns.mathML, "mover", null, [charElem, supElem]));
        } else if (subElem) {
            children.push(createElementNS(ns.mathML, "munder", null, [charElem, subElem]));
        } else {
            children.push(charElem);
        }

        let oMrow = createElementNS(ns.mathML, "mrow", null);

        appendChildren(oMrow, children);

        this.renderElements(grouped[DomType.MmlBase].children, oMrow);

        return oMrow;
    }

    renderMmlPreSubSuper(elem: OpenXmlElement) {
        const children = [];
        const grouped = keyBy(elem.children, x => x.type);

        const sup = grouped[DomType.MmlSuperArgument];
        const sub = grouped[DomType.MmlSubArgument];
        const supElem = sup ? createElementNS(ns.mathML, "mo", null, asArray(this.renderElement(sup))) : null;
        const subElem = sub ? createElementNS(ns.mathML, "mo", null, asArray(this.renderElement(sub))) : null;
        const stubElem = createElementNS(ns.mathML, "mo", null);

        children.push(createElementNS(ns.mathML, "msubsup", null, [stubElem, subElem, supElem]));

        let oMrow = createElementNS(ns.mathML, "mrow", null);

        appendChildren(oMrow, children);

        this.renderElements(grouped[DomType.MmlBase].children, oMrow);

        return oMrow;
    }

    renderMmlGroupChar(elem: OpenXmlElement) {
        const tagName = elem.props.verticalJustification === "bot" ? "mover" : "munder";
        const oGroupChar = this.renderContainerNS(elem, ns.mathML, tagName);

        if (elem.props.char) {
            let oMo = createElementNS(ns.mathML, "mo", null, [elem.props.char])
            appendChildren(oGroupChar, oMo);
        }

        return oGroupChar;
    }

    renderMmlBar(elem: OpenXmlElement) {
        const oMrow = this.renderContainerNS(elem, ns.mathML, "mrow") as MathMLElement;

        switch (elem.props.position) {
            case "top":
                oMrow.style.textDecoration = "overline";
                break
            case "bottom":
                oMrow.style.textDecoration = "underline";
                break
        }

        return oMrow;
    }

    renderMmlRun(elem: OpenXmlElement) {
        const oMs = createElementNS(ns.mathML, "ms") as HTMLElement;

        this.renderClass(elem, oMs);
        this.renderStyleValues(elem.cssStyle, oMs);
        this.renderChildren(elem, oMs);

        return oMs;
    }

    renderMllList(elem: OpenXmlElement) {
        const oMtable = createElementNS(ns.mathML, "mtable") as HTMLElement;
        // 添加class类
        this.renderClass(elem, oMtable);
        // 渲染style样式
        this.renderStyleValues(elem.cssStyle, oMtable);

        for (let child of elem.children) {

            let oChild = this.renderElement(child) as Element;

            let oMtd = createElementNS(ns.mathML, "mtd", null, [oChild]);

            let oMtr = createElementNS(ns.mathML, "mtr", null, [oMtd]);

            appendChildren(oMtable, oMtr);
        }

        return oMtable;
    }

    // 设置元素style样式
    renderStyleValues(style: Record<string, string>, output: HTMLElement) {
        for (let k in style) {
            if (k.startsWith("$")) {
                output.setAttribute(k.slice(1), style[k]);
            } else {
                output.style[k] = style[k];
            }
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

            for (let tab of this.currentTabs) {
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
        case "http://www.w3.org/2000/svg":
            oParent = document.createElementNS(ns, tagName as keyof SVGElementTagNameMap);
            break;
        case "http://www.w3.org/1999/xhtml":
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
function checkOverflow(el: Element) {
    //先让溢出效果为 hidden 这样才可以比较 clientHeight和scrollHeight
    return el.clientWidth < el.scrollWidth || el.clientHeight < el.scrollHeight;
}

// 删除单个或者多个子元素
function removeElements(target: Node[] | Node, parent: HTMLElement | Element): void;
function removeElements(target: Element[] | Element): void;
function removeElements(target: ChildrenType, parent?: HTMLElement | Element): void {
    if (Array.isArray(target)) {
        target.forEach((elem) => {
            if (elem instanceof Element) {
                elem.remove()
            } else {
                if (parent) {
                    parent.removeChild(elem)
                }
            }
        })
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
    return createElement("style", {innerHTML: cssText});
}

// 插入注释
function appendComment(elem: HTMLElement, comment: string) {
    elem.appendChild(document.createComment(comment));
}

function findParent<T extends OpenXmlElement>(elem: OpenXmlElement, type: DomType): T {
    let parent = elem.parent;

    while (parent != null && parent.type != type)
        parent = parent.parent;

    return <T>parent;
}
