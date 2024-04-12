import { DomType, IDomNumbering, NumberingPicBullet, OpenXmlElement, WmlBreak, WmlDrawing, WmlHyperlink, WmlImage, WmlNoteReference, WmlSymbol, WmlTable, WmlTableCell, WmlTableColumn, WmlTableRow, WmlText, WrapType } from './document/dom';
import { DocumentElement } from './document/document';
import { parseParagraphProperties, parseParagraphProperty, WmlParagraph } from './document/paragraph';
import { parseSectionProperties, SectionProperties } from './document/section';
import xml from './parser/xml-parser';
import { parseRunProperties, WmlRun } from './document/run';
import { parseBookmarkEnd, parseBookmarkStart } from './document/bookmarks';
import { IDomStyle, Ruleset } from './document/style';
import { WmlFieldChar, WmlFieldSimple, WmlInstructionText } from './document/fields';
import { convertLength, LengthUsage, LengthUsageType } from './document/common';
import { parseVmlElement } from './vml/vml';
import { uuid } from "./utils";
import { WmlComment, WmlCommentRangeEnd, WmlCommentRangeStart, WmlCommentReference } from './comments/elements';

export var autos = {
	shd: "inherit",
	color: "black",
	borderColor: "black",
	highlight: "transparent"
};

// TODO 支持的命名空间：wps、wpi
const supportedNamespaceURIs = [
	// "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
];

const mmlTagMap = {
	"oMath": DomType.MmlMath,
	"oMathPara": DomType.MmlMathParagraph,
	"f": DomType.MmlFraction,
	"func": DomType.MmlFunction,
	"fName": DomType.MmlFunctionName,
	"num": DomType.MmlNumerator,
	"den": DomType.MmlDenominator,
	"rad": DomType.MmlRadical,
	"deg": DomType.MmlDegree,
	"e": DomType.MmlBase,
	"sSup": DomType.MmlSuperscript,
	"sSub": DomType.MmlSubscript,
	"sPre": DomType.MmlPreSubSuper,
	"sup": DomType.MmlSuperArgument,
	"sub": DomType.MmlSubArgument,
	"d": DomType.MmlDelimiter,
	"nary": DomType.MmlNary,
	"eqArr": DomType.MmlEquationArray,
	"lim": DomType.MmlLimit,
	"limLow": DomType.MmlLimitLower,
	"m": DomType.MmlMatrix,
	"mr": DomType.MmlMatrixRow,
	"box": DomType.MmlBox,
	"bar": DomType.MmlBar,
	"groupChr": DomType.MmlGroupChar
}

export interface DocumentParserOptions {
	ignoreWidth: boolean;
	debug: boolean;
	ignoreTableWrap: boolean,
	ignoreImageWrap: boolean,
}

// 默认解析选项
export const defaultDocumentParserOptions: DocumentParserOptions = {
	ignoreWidth: false,
	debug: false,
	ignoreTableWrap: true,
	ignoreImageWrap: true,
}

export class DocumentParser {
	options: DocumentParserOptions;

	constructor(options?: Partial<DocumentParserOptions>) {
		this.options = {
			...defaultDocumentParserOptions,
			...options
		};
	}

	parseNotes(xmlDoc: Element, elemName: string, elemClass: any): any[] {
		let result = [];

		for (let el of xml.elements(xmlDoc, elemName)) {
			const node = new elemClass();
			node.id = xml.attr(el, "id");
			node.noteType = xml.attr(el, "type");
			node.children = this.parseBodyElements(el);
			result.push(node);
		}

		return result;
	}

	parseComments(xmlDoc: Element): any[] {
		let result = [];

		for (let el of xml.elements(xmlDoc, "comment")) {
			const item = new WmlComment();
			item.id = xml.attr(el, "id");
			item.author = xml.attr(el, "author");
			item.initials = xml.attr(el, "initials");
			item.date = xml.attr(el, "date");
			item.children = this.parseBodyElements(el);
			result.push(item);
		}

		return result;
	}

	parseDocumentFile(xmlDoc: Element): DocumentElement {
		let xbody = xml.element(xmlDoc, "body");
		// 背景色
		let background = xml.element(xmlDoc, "background");
		let sectPr = xml.element(xbody, "sectPr");
		// 计算节属性
		let props = {} as SectionProperties;
		if (sectPr) {
			props = parseSectionProperties(sectPr, xml);
		}
		// 生成唯一uuid标识
		props.sectionId = uuid();

		return {
			type: DomType.Document,
			children: this.parseBodyElements(xbody),
			pages: [],
			props,
			cssStyle: background ? this.parseBackground(background) : {},
		};
	}

	parseBackground(elem: Element): any {
		let result = {};
		let color = xmlUtil.colorAttr(elem, "color");

		if (color) {
			result["background-color"] = color;
		}

		return result;
	}

	parseBodyElements(element: Element): OpenXmlElement[] {
		let children = [];

		for (let elem of xml.elements(element)) {
			switch (elem.localName) {
				case "p":
					children.push(this.parseParagraph(elem));
					break;

				case "tbl":
					children.push(this.parseTable(elem));
					break;

				case "sdt":
					children.push(...this.parseSdt(elem, (e: Element) => this.parseBodyElements(e)));
					break;

				case "sectPr":
					// ignore,section property has parsed in parseDocumentFile
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Body Element：${elem.localName}`, 'color:red');
					}

			}
		}

		return children;
	}

	parseStylesFile(xstyles: Element): IDomStyle[] {
		let result = [];

		xmlUtil.foreach(xstyles, n => {
			switch (n.localName) {
				case "style":
					result.push(this.parseStyle(n));
					break;

				case "docDefaults":
					result.push(this.parseDefaultStyles(n));
					break;
				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Style File：${n.localName}`, 'color:#f75607');
					}
			}
		});

		return result;
	}

	parseDefaultStyles(node: Element): IDomStyle {
		let result = <IDomStyle>{
			basedOn: null,
			id: null,
			name: null,
			rulesets: [],
			type: null
		};

		xmlUtil.foreach(node, c => {
			switch (c.localName) {
				case "rPrDefault":
					let rPr = xml.element(c, "rPr");

					if (rPr) {
						result.rulesets.push({
							target: "span",
							declarations: this.parseDefaultProperties(rPr, {})
						});
					}
					break;

				case "pPrDefault":
					let pPr = xml.element(c, "pPr");

					if (pPr)
						result.rulesets.push({
							target: "p",
							declarations: this.parseDefaultProperties(pPr, {})
						});
					break;
				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Default Style：${c.localName}`, 'color:#f75607');
					}
			}
		});

		return result;
	}

	parseStyle(node: Element): IDomStyle {
		let result: IDomStyle = <IDomStyle>{
			basedOn: null,
			id: null,
			name: null,
			rulesets: [],
			type: null,
		};
		for (const attr of xml.attrs(node)) {
			switch (attr.localName) {
				// User-Defined Style
				case "customStyle":
					result.customStyle = xml.boolAttr(node, "customStyle", false);
					break;

				// Default Style
				case "default":
					result.isDefault = xml.boolAttr(node, "default", false);
					break;

				// Style ID
				case "styleId":
					result.id = xml.attr(node, "styleId");
					break;

				// Style Type
				case "type":
					result.type = xml.attr(node, "type");
					const typeToLabelMap = {
						"paragraph": "p",
						"table": "table",
						"character": "span",
						"numbering": "p",
					};
					// 检查result.type是否在映射中
					if (typeToLabelMap.hasOwnProperty(result.type)) {
						result.label = typeToLabelMap[result.type];
					} else {
						// 未知类型处理，确保在options.debug为false时也能处理
						if (this.options && this.options.debug) {
							console.warn(`DOCX:%c Unknown Style Type：${result.type}`, 'color:#f75607');
						}
					}
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Style Property：${attr.localName}`, 'color:#f75607');
					}
			}
		}

		xmlUtil.foreach(node, n => {
			switch (n.localName) {
				// Alternate Style Names
				case "aliases":
					result.aliases = xml.attr(n, "val").split(",");
					break;

				// Automatically Merge User Formatting Into Style Definition.
				// that change is stored on the style and therefore propagated to all locations where the style is in use.
				case "autoRedefine":
					result.autoRedefine = true;
					break;

				// Parent Style ID
				case "basedOn":
					result.basedOn = xml.attr(n, "val");
					break;

				// Hide Style From User Interface
				case "hidden":
					result.hidden = true;
					break;

				// Linked Style Reference
				case "link":
					result.linked = xml.attr(n, "val");
					break;

				// Style Cannot Be Applied
				case "locked":
					result.locked = true;
					break;

				// Primary Style Name
				case "name":
					result.name = xml.attr(n, "val");
					break;

				// Style For Next Paragraph
				case "next":
					result.next = xml.attr(n, "val");
					break;

				// E-Mail Message Text Style
				case "personal":
					result.personal = xml.boolAttr(n, "val");
					break;

				// E-Mail Message Composition Style
				case "personalCompose":
					result.personalCompose = xml.boolAttr(n, "val");
					break;

				// E-Mail Message Reply Style
				case "personalReply":
					result.personalReply = xml.boolAttr(n, "val");
					break;

				// Style Paragraph Properties
				case "pPr":
					result.rulesets.push({
						target: "p",
						declarations: this.parseDefaultProperties(n, {})
					});
					result.paragraphProps = parseParagraphProperties(n, xml);
					break;

				// Specifies Primary Style
				case "qFormat":
					result.primaryStyle = true;
					break;

				// Run Properties
				case "rPr":
					result.rulesets.push({
						target: "span",
						declarations: this.parseDefaultProperties(n, {})
					});
					result.runProps = parseRunProperties(n, xml);
					break;

				// Revision Identifier for Style Definition.Single Session Revision Save ID.
				case "rsid":
					result.rsid = xml.hexAttr(n, "val");
					break;

				// 	Hide Style From Main User Interface.
				// 	This setting is intended to define a style property which allows styles to be seen and modified in an advanced user interface, without exposing the style in a less advanced setting
				case "semiHidden":
					result.semiHidden = true;
					break;

				// Style Table Properties
				case "tblPr":
					result.rulesets.push({
						target: "td",
						declarations: this.parseDefaultProperties(n, {})
					});
					break;

				// Style Table Row Properties
				case "trPr":
					//TODO: maybe move to processor
					result.rulesets.push({
						target: "tr",
						declarations: this.parseDefaultProperties(n, {})
					});
					break;

				// Style Table Cell Properties
				case "tcPr":
					result.rulesets.push({
						target: "td",
						declarations: this.parseDefaultProperties(n, {})
					});
					break;

				// Style Conditional Table Formatting Properties
				case "tblStylePr":
					for (let s of this.parseTableStyle(n)) {
						result.rulesets.push(s);
					}
					break;

				// Optional User Interface Sorting Order
				// This element specifies a number which can be used to sort the set of style definitions in a user interface when this document is loaded by an application
				// If this element is omitted, then the style shall have more or less an Infinity value and shall be sorted to the end of the list of style definitions
				case "uiPriority":
					result.uiPriority = xml.intAttr(n, "val", Infinity);
					break;

				// 	Remove Semi-Hidden Property When Style Is Used
				case "unhideWhenUsed":
					result.unhideWhenUsed = true;
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Style element：${n.localName}`, 'color:blue');
					}
			}
		});

		return result;
	}

	// TODO 表格style样式规则未生效
	parseTableStyle(node: Element): Ruleset[] {
		let result = [];

		let type = xml.attr(node, "type");
		let selector = "";
		let modificator = "";

		switch (type) {
			case "firstRow":
				modificator = ".first-row";
				selector = "tr.first-row td";
				break;
			case "lastRow":
				modificator = ".last-row";
				selector = "tr.last-row td";
				break;
			case "firstCol":
				modificator = ".first-col";
				selector = "td.first-col";
				break;
			case "lastCol":
				modificator = ".last-col";
				selector = "td.last-col";
				break;
			case "band1Vert":
				modificator = ":not(.no-vband)";
				selector = "td.odd-col";
				break;
			case "band2Vert":
				modificator = ":not(.no-vband)";
				selector = "td.even-col";
				break;
			case "band1Horz":
				modificator = ":not(.no-hband)";
				selector = "tr.odd-row";
				break;
			case "band2Horz":
				modificator = ":not(.no-hband)";
				selector = "tr.even-row";
				break;
			default:
				return [];
		}

		xmlUtil.foreach(node, n => {
			switch (n.localName) {
				case "pPr":
					result.push({
						target: `${selector} p`,
						mod: modificator,
						values: this.parseDefaultProperties(n, {})
					});
					break;

				case "rPr":
					result.push({
						target: `${selector} span`,
						mod: modificator,
						values: this.parseDefaultProperties(n, {})
					});
					break;

				case "tblPr":
				case "tcPr":
					result.push({
						target: selector, //TODO: maybe move to processor
						mod: modificator,
						values: this.parseDefaultProperties(n, {})
					});
					break;
				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Table Style：${n.localName}`, 'color:#f75607');
					}
			}
		});

		return result;
	}

	parseNumberingFile(xnums: Element): IDomNumbering[] {
		let result = [];
		const mapping = {};
		let bullets = [];

		xmlUtil.foreach(xnums, n => {
			switch (n.localName) {
				case "abstractNum":
					this.parseAbstractNumbering(n, bullets)
						.forEach(x => result.push(x));
					break;

				case "numPicBullet":
					bullets.push(this.parseNumberingPicBullet(n));
					break;

				case "num":
					let numId = xml.attr(n, "numId");
					let abstractNumId = xml.elementAttr(n, "abstractNumId", "val");
					mapping[abstractNumId] = numId;
					break;
				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Numbering File：${n.localName}`, 'color:#f75607');
					}
			}
		});

		result.forEach(x => x.id = mapping[x.id]);

		return result;
	}

	parseNumberingPicBullet(elem: Element): NumberingPicBullet {
		let pict = xml.element(elem, "pict");
		let shape = pict && xml.element(pict, "shape");
		let imagedata = shape && xml.element(shape, "imagedata");

		return imagedata ? {
			id: xml.intAttr(elem, "numPicBulletId"),
			src: xml.attr(imagedata, "id"),
			style: xml.attr(shape, "style")
		} : null;
	}

	parseAbstractNumbering(node: Element, bullets: any[]): IDomNumbering[] {
		let result = [];
		let id = xml.attr(node, "abstractNumId");

		xmlUtil.foreach(node, n => {
			switch (n.localName) {
				case "lvl":
					result.push(this.parseNumberingLevel(id, n, bullets));
					break;
				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Abstract Numbering：${n.localName}`, 'color:#f75607');
					}
			}
		});

		return result;
	}

	parseNumberingLevel(id: string, node: Element, bullets: any[]): IDomNumbering {
		let result: IDomNumbering = {
			id: id,
			level: xml.intAttr(node, "ilvl"),
			start: 1,
			pStyleName: undefined,
			pStyle: {},
			rStyle: {},
			suff: "tab"
		};

		xmlUtil.foreach(node, n => {
			switch (n.localName) {
				case "start":
					result.start = xml.intAttr(n, "val");
					break;

				case "pPr":
					this.parseDefaultProperties(n, result.pStyle);
					break;

				case "rPr":
					this.parseDefaultProperties(n, result.rStyle);
					break;

				case "lvlPicBulletId":
					let id = xml.intAttr(n, "val");
					result.bullet = bullets.find(x => x.id == id);
					break;

				case "lvlText":
					result.levelText = xml.attr(n, "val");
					break;

				case "pStyle":
					result.pStyleName = xml.attr(n, "val");
					break;

				case "numFmt":
					result.format = xml.attr(n, "val");
					break;

				case "suff":
					result.suff = xml.attr(n, "val");
					break;
				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Numbering Level：${n.localName}`, 'color:#f75607');
					}
			}
		});

		return result;
	}

	parseSdt(node: Element, parser: Function): OpenXmlElement[] {
		const sdtContent = xml.element(node, "sdtContent");
		return sdtContent ? parser(sdtContent) : [];
	}

	parseInserted(node: Element, parentParser: Function): OpenXmlElement {
		return <OpenXmlElement>{
			type: DomType.Inserted,
			children: parentParser(node)?.children ?? []
		};
	}

	parseDeleted(node: Element, parentParser: Function): OpenXmlElement {
		return <OpenXmlElement>{
			type: DomType.Deleted,
			children: parentParser(node)?.children ?? []
		};
	}

	parseParagraph(node: Element): OpenXmlElement {
		let wmlParagraph = <WmlParagraph>{ type: DomType.Paragraph, children: [] };

		for (let el of xml.elements(node)) {
			switch (el.localName) {
				case "pPr":
					this.parseParagraphProperties(el, wmlParagraph);
					break;

				case "r":
					wmlParagraph.children.push(this.parseRun(el, wmlParagraph));
					break;

				case "hyperlink":
					wmlParagraph.children.push(this.parseHyperlink(el, wmlParagraph));
					break;

				case "bookmarkStart":
					wmlParagraph.children.push(parseBookmarkStart(el, xml));
					break;

				case "bookmarkEnd":
					wmlParagraph.children.push(parseBookmarkEnd(el, xml));
					break;

				case "commentRangeStart":
					wmlParagraph.children.push(new WmlCommentRangeStart(xml.attr(el, "id")));
					break;

				case "commentRangeEnd":
					wmlParagraph.children.push(new WmlCommentRangeEnd(xml.attr(el, "id")));
					break;

				case "oMath":
				case "oMathPara":
					wmlParagraph.children.push(this.parseMathElement(el));
					break;

				case "sdt":
					wmlParagraph.children.push(...this.parseSdt(el, (e: Element) => this.parseParagraph(e).children));
					break;

				case "ins":
					wmlParagraph.children.push(this.parseInserted(el, (e: Element) => this.parseParagraph(e)));
					break;

				case "del":
					wmlParagraph.children.push(this.parseDeleted(el, (e: Element) => this.parseParagraph(e)));
					break;
				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Paragraph Element：${el.localName}`, 'color:#f75607');
					}
			}
		}
		// when paragraph is empty, a br tag needs to be added to work with the rich text editor and generate line height
		// 当段落children为空，需要添加一个br标签，配合富文本编辑器，同时产生行高
		// TODO 实体符号来替换空行
		if (wmlParagraph.children.length === 0) {
			let wmlBreak: WmlBreak = { type: DomType.Break, "break": "textWrapping" };
			let wmlRun = { type: DomType.Run, children: [wmlBreak as OpenXmlElement] } as WmlRun;
			wmlParagraph.children = [wmlRun];
		}

		return wmlParagraph;
	}

	parseParagraphProperties(elem: Element, paragraph: WmlParagraph) {
		this.parseDefaultProperties(elem, paragraph.cssStyle = {}, null, c => {
			if (parseParagraphProperty(c, paragraph, xml)) {
				return true;
			}

			switch (c.localName) {
				case "pStyle":
					paragraph.styleName = xml.attr(c, "val");
					break;

				case "cnfStyle":
					paragraph.className = values.classNameOfCnfStyle(c);
					break;

				case "framePr":
					this.parseFrame(c, paragraph);
					break;

				case "rPr":
					//TODO ignore
					break;

				default:
					// pass other properties to parseDefaultProperties function
					return false;
			}

			return true;
		});
	}

	parseFrame(node: Element, paragraph: WmlParagraph) {
		let dropCap = xml.attr(node, "dropCap");

		if (dropCap == "drop")
			paragraph.cssStyle["float"] = "left";
	}

	parseHyperlink(node: Element, parent?: OpenXmlElement): WmlHyperlink {
		let result: WmlHyperlink = <WmlHyperlink>{ type: DomType.Hyperlink, parent: parent, children: [] };
		let anchor = xml.attr(node, "anchor");
		let relId = xml.attr(node, "id");

		if (anchor)
			result.href = "#" + anchor;

		if (relId)
			result.id = relId;

		xmlUtil.foreach(node, c => {
			switch (c.localName) {
				case "r":
					result.children.push(this.parseRun(c, result));
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Hyperlink Element：${c.localName}`, 'color:#f75607');
					}
			}
		});

		return result;
	}

	parseRun(node: Element, parent?: OpenXmlElement): WmlRun {
		let result: WmlRun = <WmlRun>{ type: DomType.Run, parent: parent, children: [] };

		xmlUtil.foreach(node, c => {
			c = this.checkAlternateContent(c);

			switch (c.localName) {
				case "t":
					let textContent = c.textContent;
					// 是否保留空格
					let is_preserve_space = xml.attr(c, "xml:space") === "preserve";
					if (is_preserve_space) {
						// \u00A0 = 不间断空格，英文应该一个空格，中文两个空格。受到font-family影响。
						textContent = textContent.split(/\s/).join("\u00A0");
					}
					result.children.push(<WmlText>{
						type: DomType.Text,
						text: textContent
					});
					break;

				case "delText":
					result.children.push(<WmlText>{
						type: DomType.DeletedText,
						text: c.textContent
					});
					break;

				case "commentReference":
					result.children.push(new WmlCommentReference(xml.attr(c, "id")));
					break;

				case "fldSimple":
					result.children.push(<WmlFieldSimple>{
						type: DomType.SimpleField,
						instruction: xml.attr(c, "instr"),
						lock: xml.boolAttr(c, "lock", false),
						dirty: xml.boolAttr(c, "dirty", false)
					});
					break;

				case "instrText":
					result.fieldRun = true;
					result.children.push(<WmlInstructionText>{
						type: DomType.Instruction,
						text: c.textContent
					});
					break;

				case "fldChar":
					result.fieldRun = true;
					result.children.push(<WmlFieldChar>{
						type: DomType.ComplexField,
						charType: xml.attr(c, "fldCharType"),
						lock: xml.boolAttr(c, "lock", false),
						dirty: xml.boolAttr(c, "dirty", false)
					});
					break;

				case "noBreakHyphen":
					result.children.push({ type: DomType.NoBreakHyphen });
					break;

				case "br":
					result.children.push(<WmlBreak>{
						type: DomType.Break,
						break: xml.attr(c, "type") || "textWrapping",
						props: {
							clear: xml.attr(c, "clear")
						}
					});
					break;

				case "lastRenderedPageBreak":
					result.children.push(<WmlBreak>{
						type: DomType.Break,
						break: "lastRenderedPageBreak"
					});
					break;
				// SymbolChar：符号字符
				case "sym":
					result.children.push(<WmlSymbol>{
						type: DomType.Symbol,
						font: xml.attr(c, "font"),
						char: xml.attr(c, "char")
					});
					break;

				case "tab":
					result.children.push({ type: DomType.Tab });
					break;

				case "footnoteReference":
					result.children.push(<WmlNoteReference>{
						type: DomType.FootnoteReference,
						id: xml.attr(c, "id")
					});
					break;

				case "endnoteReference":
					result.children.push(<WmlNoteReference>{
						type: DomType.EndnoteReference,
						id: xml.attr(c, "id")
					});
					break;

				case "drawing":
					let d = this.parseDrawing(c);

					if (d)
						result.children = [d];
					break;

				case "pict":
					result.children.push(this.parseVmlPicture(c));
					break;

				case "rPr":
					this.parseRunProperties(c, result);
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Run Element：${c.localName}`, 'color:#f75607');
					}
			}
		});

		return result;
	}

	parseMathElement(elem: Element): OpenXmlElement {
		const propsTag = `${elem.localName}Pr`;
		const result = { type: mmlTagMap[elem.localName], children: [] } as OpenXmlElement;

		for (const el of xml.elements(elem)) {
			const childType = mmlTagMap[el.localName];

			if (childType) {
				result.children.push(this.parseMathElement(el));
			} else if (el.localName == "r") {
				let run = this.parseRun(el);
				run.type = DomType.MmlRun;
				result.children.push(run);
			} else if (el.localName == propsTag) {
				result.props = this.parseMathProperties(el);
			}
		}

		return result;
	}

	parseMathProperties(elem: Element): Record<string, any> {
		const result: Record<string, any> = {};

		for (const el of xml.elements(elem)) {
			switch (el.localName) {
				case "chr":
					result.char = xml.attr(el, "val");
					break;

				case "vertJc":
					result.verticalJustification = xml.attr(el, "val");
					break;

				case "pos":
					result.position = xml.attr(el, "val");
					break;

				case "degHide":
					result.hideDegree = xml.boolAttr(el, "val");
					break;

				case "begChr":
					result.beginChar = xml.attr(el, "val");
					break;

				case "endChr":
					result.endChar = xml.attr(el, "val");
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Math Property：${el.localName}`, 'color:#f75607');
					}
			}
		}

		return result;
	}

	parseRunProperties(elem: Element, run: WmlRun) {
		this.parseDefaultProperties(elem, run.cssStyle = {}, null, c => {
			switch (c.localName) {
				// Referenced Character Style
				case "rStyle":
					run.styleName = xml.attr(c, "val");
					break;

				// Subscript/Superscript Text
				case "vertAlign":
					run.verticalAlign = values.valueOfVertAlign(c, true);
					break;

				default:
					// pass other properties to parseDefaultProperties function
					return false;
			}

			return true;
		});
	}

	parseVmlPicture(elem: Element): OpenXmlElement {
		const result = { type: DomType.VmlPicture, children: [] };

		for (const el of xml.elements(elem)) {
			const child = parseVmlElement(el, this);
			child && result.children.push(child);
		}

		return result;
	}

	// 检测备选内容
	checkAlternateContent(elem: Element): Element {
		if (elem.localName != 'AlternateContent') {
			return elem;
		}

		let choice = xml.element(elem, "Choice");
		// 备选项
		if (choice) {
			let requires = xml.attr(choice, "Requires");
			let namespaceURI = elem.lookupNamespaceURI(requires);

			if (supportedNamespaceURIs.includes(namespaceURI)) {
				return choice.firstElementChild;
			}
		}
		// 回退
		return xml.element(elem, "Fallback")?.firstElementChild;
	}

	parseDrawing(node: Element): OpenXmlElement {
		for (let n of xml.elements(node)) {
			switch (n.localName) {
				case "inline":
				case "anchor":
					return this.parseDrawingWrapper(n);
				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Drawing Element：${n.localName}`, 'color:#f75607');
					}
			}
		}
	}

	// TODO 图片旋转、裁剪之后，文字环绕计算错误
	// DrawingML对象有两种状态：内联（inline）-- 对象与文本对齐，浮动（anchor）--对象在文本中浮动，但可以相对于页面进行绝对定位
	parseDrawingWrapper(node: Element): OpenXmlElement {
		// 是否布局在表格中
		let layoutInCell = xml.boolAttr(node, "layoutInCell");
		// 是否锁定
		let locked = xml.boolAttr(node, "locked");
		// 是否在文字后面显示
		let behindDoc = xml.boolAttr(node, "behindDoc");
		// 是否允许重叠
		let allowOverlap = xml.boolAttr(node, "allowOverlap");
		// 是否简单定位
		let simplePos = xml.boolAttr(node, "simplePos");
		// 层叠数值
		let relativeHeight = xml.intAttr(node, "relativeHeight", 1);
		// 计算DrawML对象相对于文字的上下左右间距；仅在浮动、文字环绕模式下有效；
		let distance = {
			left: xml.lengthAttr(node, "distL", LengthUsage.Emu),
			right: xml.lengthAttr(node, "distR", LengthUsage.Emu),
			top: xml.lengthAttr(node, "distT", LengthUsage.Emu),
			bottom: xml.lengthAttr(node, "distB", LengthUsage.Emu),
			distL: xml.intAttr(node, "distL", 0),
			distR: xml.intAttr(node, "distR", 0),
			distT: xml.intAttr(node, "distT", 0),
			distB: xml.intAttr(node, "distB", 0),
		}

		let result: WmlDrawing = {
			type: DomType.Drawing,
			children: [],
			cssStyle: {},
			props: {
				localName: node.localName,
				wrapType: null,
				layoutInCell,
				locked,
				behindDoc,
				allowOverlap,
				simplePos,
				relativeHeight,
				distance,
				extent: {},
			},
		};

		interface Position {
			relative: string;
			align: string;
			offset: string;
			origin: number;
		}

		// 横轴定位
		let posX: Position = { relative: "page", align: "left", offset: "0pt", origin: 0, };
		// 纵轴定位
		let posY: Position = { relative: "page", align: "top", offset: "0pt", origin: 0, };

		for (let n of xml.elements(node)) {
			switch (n.localName) {
				case "simplePos":
					// 简单定位
					if (simplePos) {
						posX.offset = xml.lengthAttr(n, "x", LengthUsage.Emu);
						posY.offset = xml.lengthAttr(n, "y", LengthUsage.Emu);
						posX.origin = xml.intAttr(n, "x", 0);
						posY.origin = xml.intAttr(n, "y", 0);
					}
					break;

				case "positionH":
					if (!simplePos) {
						let alignNode = xml.element(n, "align");
						let offsetNode = xml.element(n, "posOffset");

						posX.relative = xml.attr(n, "relativeFrom") ?? posX.relative;

						if (alignNode) {
							posX.align = alignNode.textContent;
						}

						if (offsetNode) {
							posX.offset = xmlUtil.sizeValue(offsetNode, LengthUsage.Emu);
							posX.origin = xmlUtil.parseTextContent(offsetNode);
						}
						// 设置横轴的属性
						result.props.posX = posX;
					}
					break;

				case "positionV":
					if (!simplePos) {
						let alignNode = xml.element(n, "align");
						let offsetNode = xml.element(n, "posOffset");

						posY.relative = xml.attr(n, "relativeFrom") ?? posY.relative;

						if (alignNode) {
							posY.align = alignNode.textContent;
						}

						if (offsetNode) {
							posY.offset = xmlUtil.sizeValue(offsetNode, LengthUsage.Emu);
							posY.origin = xmlUtil.parseTextContent(offsetNode);
						}
						// 设置纵轴的属性
						result.props.posY = posY;
					}
					break;

				// drawing外框尺寸
				case "extent":
					result.props.extent = {
						width: xml.lengthAttr(n, "cx", LengthUsage.Emu),
						height: xml.lengthAttr(n, "cy", LengthUsage.Emu),
						origin_width: xml.intAttr(n, "cx", 0),
						origin_height: xml.intAttr(n, "cy", 0),
					};
					break;

				// 特效占据空间
				case "effectExtent":
					result.props.effectExtent = {
						top: xml.lengthAttr(n, "t", LengthUsage.Emu),
						bottom: xml.lengthAttr(n, "b", LengthUsage.Emu),
						left: xml.lengthAttr(n, "l", LengthUsage.Emu),
						right: xml.lengthAttr(n, "r", LengthUsage.Emu),
						origin_top: xml.intAttr(n, "t", 0),
						origin_bottom: xml.intAttr(n, "b", 0),
						origin_left: xml.intAttr(n, "l", 0),
						origin_right: xml.intAttr(n, "r", 0),
					};
					break;

				// 图片
				case "graphic":
					let g = this.parseGraphic(n);

					if (g) {
						result.children.push(g);
					}
					break;
				case "wrapTopAndBottom":
					result.props.wrapType = WrapType.TopAndBottom;
					break;

				case "wrapNone":
					result.props.wrapType = WrapType.None;
					break;

				case "wrapSquare":
					result.props.wrapType = WrapType.Square;
					// 文本环绕位置：bothSides、largest、left、right
					result.props.wrapText = xml.attr(n, "wrapText");
					break;

				case "wrapThrough":
				case "wrapTight":
					result.props.wrapType = WrapType.Tight;
					// 文本环绕位置：bothSides、largest、left、right
					result.props.wrapText = xml.attr(n, "wrapText");
					// 多边形数据
					let polygonNode = xml.element(n, "wrapPolygon");
					this.parsePolygon(polygonNode, result);
					break;
				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Drawing Property：${n.localName}`, 'color:#f75607');
					}
			}
		}
		// 重新计算DrawWrapper的空间
		let { extent, effectExtent } = result.props;
		let real_width = extent.origin_width + effectExtent.origin_left + effectExtent.origin_right;
		let real_height = extent.origin_height + effectExtent.origin_top + effectExtent.origin_bottom;
		result.cssStyle["width"] = convertLength(real_width, LengthUsage.Emu);
		result.cssStyle["height"] = convertLength(real_height, LengthUsage.Emu);
		// 内联（inline）--嵌入型环绕
		if (node.localName === "inline") {
			result.props.wrapType = WrapType.Inline;
		}
		// 浮动（anchor）--其他环绕
		if (node.localName === "anchor") {
			// 根据relativeHeight设置z-index
			result.cssStyle["position"] = "relative";
			// 根据behindDoc判断，衬于文字下方、浮于文字上方
			if (behindDoc) {
				result.cssStyle["z-index"] = -1;
			} else {
				result.cssStyle["z-index"] = relativeHeight;
			}
			// 图片文字环绕默认采用wrapTopAndBottom
			if (this.options.ignoreImageWrap) {
				result.props.wrapType = WrapType.TopAndBottom;
			}
			// 文本环绕位置：bothSides、largest、left、right
			let { wrapText, wrapType, extent } = result.props;

			switch (wrapType) {
				// 顶部底部文字环绕
				case WrapType.TopAndBottom:
					result.cssStyle['float'] = 'left';
					result.cssStyle['width'] = "100%";
					// 水平对齐方式，目前仅支持left、right、center
					result.cssStyle['text-align'] = posX.align;
					// 横轴位移补偿
					result.cssStyle["transform"] = `translate(${posX.offset},0)`;
					// 垂直方向，纵轴位移
					result.cssStyle["margin-top"] = `calc(${posY.offset} - ${distance.top})`;
					// 计算距离顶部的inset
					result.cssStyle["shape-outside"] = `inset(calc(${posY.offset} - ${distance.top}) 0 0 0)`;
					// TODO 图片位于文字中间，定位计算错误
					// DrawML对象与文字的上下间距
					result.cssStyle["box-sizing"] = "content-box";
					result.cssStyle["padding-top"] = distance.top;
					result.cssStyle["padding-bottom"] = distance.bottom;
					break;

				// 衬于文字下方、浮于文字上方
				case WrapType.None:
					result.cssStyle['position'] = 'absolute';
					// 水平对齐方式，目前仅支持left、right、center
					switch (posX.align) {
						case "left":
						case "right":
							result.cssStyle[posX.align] = posX.offset;
							break;
						case "center":
							result.cssStyle["left"] = "50%";
							result.cssStyle["transform"] = "translateX(-50%)";
					}
					// 垂直方向，纵轴位移
					result.cssStyle["top"] = posY.offset;

					break;

				// 矩形（四周型）环绕
				case WrapType.Square:
					// TODO 环绕位置bothSides、largest无法实现，目前仅支持left、right
					result.cssStyle["float"] = wrapText === 'left' ? "right" : "left";
					// 垂直方向，纵轴位移
					result.cssStyle["margin-top"] = `calc(${posY.offset} - ${distance.top})`;
					// 计算距离顶部的inset
					result.cssStyle["shape-outside"] = `inset(calc(${posY.offset} - ${distance.top}) 0 0 0)`;
					// wrapText：文字所在的一侧
					switch (wrapText) {
						case "left":
							// 水平对齐方式，目前仅支持left、right、center
							switch (posX.align) {
								case "left":
									// 计算公式：段落width - posX.offset - Drawing对象width - Drawing对象padding-right
									result.cssStyle["margin-right"] = `calc(100% - ${extent.width} - ${posX.offset} - ${distance.right})`;
									break;
								case "right":
									result.cssStyle["margin-right"] = `calc(${posX.offset} - ${distance.right})`;
									break;
								case "center":
									result.cssStyle["margin-right"] = `calc( 50% - (${extent.width} - ${posX.offset}) / 2 - ${distance.right} )`;
							}
							break;
						case "right":
							// 水平对齐方式，目前仅支持left、right、center
							switch (posX.align) {
								case "left":
									result.cssStyle["margin-left"] = `calc(${posX.offset} - ${distance.left})`;
									break;
								case "right":
									// 计算公式：段落width - posX.offset - Drawing对象width - Drawing对象padding-right
									result.cssStyle["margin-left"] = `calc(100% - ${extent.width} - ${posX.offset} - ${distance.left})`;
									result.cssStyle["margin-right"] = `calc(${posX.offset} - ${distance.right})`;
									break;
								case "center":
									result.cssStyle["margin-left"] = `calc( 50% - (${extent.width} - ${posX.offset} ) / 2 - ${distance.left} )`;
							}

							break;
						default:
							console.error(`text wrap picture on ${wrapText} is not supported！`)
							break;
					}
					// DrawML对象与文字的上下间距
					result.cssStyle["box-sizing"] = "content-box";
					result.cssStyle["padding-top"] = distance.top;
					result.cssStyle["padding-bottom"] = distance.bottom;
					result.cssStyle["padding-left"] = distance.left;
					result.cssStyle["padding-right"] = distance.right;

					break;

				// 穿越型环绕
				case WrapType.Through:
				// 紧密型环绕
				case WrapType.Tight:
					// TODO 环绕位置bothSides、largest无法实现，目前仅支持left、right
					result.cssStyle["float"] = wrapText === 'left' ? "right" : "left";
					// 根据多边形设置环绕
					let { polygonData } = result.props;
					result.cssStyle["shape-outside"] = `polygon(${polygonData})`;

					// TODO shape-margin目前4个方位只能设置统一的数值.暂时无法采用

					// 垂直方向，纵轴位移
					// TODO 存在上下padding时，定位错误
					result.cssStyle["margin-top"] = posY.offset;

					switch (wrapText) {
						case "left":
							// 水平对齐方式，目前仅支持left、right、center
							switch (posX.align) {
								case "left":
									// 计算公式：段落width - posX.offset - Drawing对象width
									result.cssStyle["margin-right"] = `calc(100% - ${extent.width} - ${posX.offset})`;
									break;
								case "right":
									result.cssStyle["margin-right"] = posX.offset;
									break;
								case "center":
									result.cssStyle["margin-right"] = `calc( 50% - (${extent.width} - ${posX.offset}) / 2 )`;
							}
							break;
						case "right":
							// 水平对齐方式，目前仅支持left、right、center
							switch (posX.align) {
								case "left":
									result.cssStyle["margin-left"] = posX.offset;
									break;
								case "right":
									// 计算公式：段落width - posX.offset - Drawing对象width
									result.cssStyle["margin-left"] = `calc(100% - ${extent.width} - ${posX.offset})`;
									break;
								case "center":
									result.cssStyle["margin-left"] = `calc( 50% - (${extent.width} - ${posX.offset} ) / 2 )`;
							}
							break;
						default:
							console.error(`text wrap picture on ${wrapText} is not supported！`)
							break;
					}
					break;
			}
		}

		return result;
	}

	/*
	* 多边形端点数据
	* Office Open XML将X和Y属性解释为固定坐标空间（21600x21600）中的坐标，每个坐标点在x轴和y轴上都有对应的值，范围从0到21599。
	* 固定坐标空间 => 实际坐标空间：
	* 实际坐标X = 固定坐标X(EMU) * 图形的Width / 21600
	* 实际坐标Y = 固定坐标Y(EMU) * 图形的Height / 21600
	*/
	parsePolygon(node: Element, target: OpenXmlElement) {
		let polygon = [];
		let { wrapText, distance, extent, posX, posY } = target.props;

		xmlUtil.foreach(node, (elem) => {
			// 原始值，单位：EMU
			let origin_x = xml.intAttr(elem, 'x', 0);
			let origin_y = xml.intAttr(elem, 'y', 0);
			// 实际坐标，单位EMU
			let real_x: number, real_y: number;
			// Point坐标，单位pt
			let point_x: string | number, point_y: string | number;
			// 修正坐标，补偿横向位移
			let revise_x: string | number, revise_y: string | number;
			/*
			* 根据wrapText，转换坐标
			* TODO 多边形：纵轴外边距暂时忽略，横轴补偿distance。当多边形超出DrawWrapper的范围时，补偿会被忽略，导致不准确
			*/
			switch (wrapText) {
				case "left":
					// 水平对齐方式，目前仅支持left、right、center
					switch (posX.align) {
						case "left":
							// 实际坐标
							real_x = origin_x * extent.origin_width / 21600 - distance.distL;
							real_y = origin_y * extent.origin_height / 21600 + posY.origin;
							// 修正坐标
							revise_x = convertLength(real_x, LengthUsage.Emu) ?? "0pt";
							revise_y = convertLength(real_y, LengthUsage.Emu) ?? "0pt";
							break;
						case "right":
							// 实际坐标
							real_x = origin_x * extent.origin_width / 21600 + posX.origin - distance.distL;
							real_y = origin_y * extent.origin_height / 21600 + posY.origin;
							// 修正坐标
							revise_x = convertLength(real_x, LengthUsage.Emu) ?? "0pt";
							revise_y = convertLength(real_y, LengthUsage.Emu) ?? "0pt";
							break;
						case "center":
							// 实际坐标
							real_x = origin_x * extent.origin_width / 21600 + posX.origin - distance.distL;
							real_y = origin_y * extent.origin_height / 21600 + posY.origin;
							// 修正坐标
							revise_x = convertLength(real_x, LengthUsage.Emu) ?? "0pt";
							revise_y = convertLength(real_y, LengthUsage.Emu) ?? "0pt";
					}
					break;
				case "right":
					// 水平对齐方式，目前仅支持left、right、center
					switch (posX.align) {
						case "left":
							// 实际坐标
							real_x = origin_x * extent.origin_width / 21600 + posX.origin + distance.distR;
							real_y = origin_y * extent.origin_height / 21600 + posY.origin;
							// 修正坐标
							revise_x = convertLength(real_x, LengthUsage.Emu) ?? "0pt";
							revise_y = convertLength(real_y, LengthUsage.Emu) ?? "0pt";
							break;
						case "right":
							// 实际坐标
							real_x = origin_x * extent.origin_width / 21600 + posX.origin + distance.distR;
							real_y = origin_y * extent.origin_height / 21600 + posY.origin;
							// Point坐标
							point_x = convertLength(real_x, LengthUsage.Emu) ?? "0pt";
							point_y = convertLength(real_y, LengthUsage.Emu) ?? "0pt";
							// 修正坐标，横轴补偿distance
							revise_x = `calc(100% + ${point_x} - ${extent.width})`;
							revise_y = point_y;

							break;
						case "center":
							// 实际坐标
							real_x = origin_x * extent.origin_width / 21600 + posX.origin + distance.distR;
							real_y = origin_y * extent.origin_height / 21600 + posY.origin;
							// Point坐标
							point_x = convertLength(real_x, LengthUsage.Emu) ?? "0pt";
							point_y = convertLength(real_y, LengthUsage.Emu) ?? "0pt";
							// 修正坐标，横轴补偿distance
							revise_x = `calc(50% + ${point_x})`;
							revise_y = point_y;
					}

					break;
				default:
					console.error(`text wrap picture on ${wrapText} is not supported！`)
					break;
			}

			let point = `${revise_x} ${revise_y}`;
			polygon.push(point);
		});
		target.props.polygonData = polygon.join(',');
	}

	parseGraphic(elem: Element): OpenXmlElement {
		let graphicData = xml.element(elem, "graphicData");

		for (let n of xml.elements(graphicData)) {
			switch (n.localName) {
				// TODO DrawML其他元素
				// shape图形
				case "wsp":
					return this.parseShape(n);

				// 图片
				case "pic":
					return this.parsePicture(n);

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Graphic Element：${n.localName}`, 'color:#f75607');
					}
			}
		}

		return null;
	}

	// 解析图形shape
	parseShape(node: Element) {
		let shape: OpenXmlElement = { type: DomType.Shape, cssStyle: {} }
		// TODO 预制图形
		for (let n of xml.elements(node)) {
			switch (n.localName) {
				case "cNvPr":
				case "cNvSpPr":
				case "cNvCnPr":
				// 图形属性
				case "spPr":
					return this.parseShapeProperties(n, shape);
				// 图形样式
				case "style":

				case "txbx":
				case "linkedTxbx":
				// 指定形状中文本正文的正文属性。
				case "bodyPr":

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Shape Element：${n.localName}`, 'color:#f75607');
					}
			}
		}
		return null;
	}

	// 图形属性
	parseShapeProperties(node: Element, target: OpenXmlElement) {

		for (let n of xml.elements(node)) {
			switch (n.localName) {
				case "xfrm":
					// 注意：存在多种变换组合的情况,需要统一合并处理
					// 水平翻转
					let flipH = xml.boolAttr(n, "flipH");
					if (flipH) {
						target.props.is_transform = true;
						target.props.transform.scaleX = -1;
					}
					// 垂直翻转
					let flipV = xml.boolAttr(n, "flipV");
					if (flipV) {
						target.props.is_transform = true;
						target.props.transform.scaleY = -1;
					}
					// 旋转角度
					let degree = xml.numberAttr(n, "rot", LengthUsage.degree, 0);
					if (degree) {
						target.props.is_transform = true;
						target.props.transform.rotate = degree;
					}
					// 子元素
					this.parseTransform2D(n, target);
					break;
				case "custGeom":
				case "prstGeom":
				case "noFill":
				case "solidFill":
				case "gradFill":
				case "blipFill":
				case "pattFill":
				case "grpFill":
				case "ln":
				case "effectLst":
				case "effectDag":
				case "scene3d":
				case "sp3d":
				case "extLst":
				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Shape Property：${n.localName}`, 'color:#f75607');
					}
			}
		}
		return null;
	}

	// 解析图片
	parsePicture(elem: Element): WmlImage {
		let result: WmlImage = {
			type: DomType.Image,
			src: "",
			cssStyle: {},
			props: {
				is_clip: false,
				clip: {},
				is_transform: false,
				transform: {},
			}
		};
		for (let n of xml.elements(elem)) {
			switch (n.localName) {
				case "nvPicPr":
					break;
				case "blipFill":
					this.parseBlipFill(n, result);
					break;

				case "spPr":
					this.parseShapeProperties(n, result)
					break;
				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Picture Element：${n.localName}`, 'color:#f75607');
					}
			}
		}

		return result;
	}

	// 2D变换
	parseTransform2D(node: Element, target: OpenXmlElement) {
		for (let n of xml.elements(node)) {
			switch (n.localName) {
				// 变换之前的宽高，实际上无效
				case "ext":
					let { transform } = target.props;
					let origin_width = xml.intAttr(n, "cx", 0);
					let origin_height = xml.intAttr(n, "cy", 0);
					// 实际的宽高，单位emu
					let width: number;
					let height: number;
					// 根据旋转角度，重新计算宽高
					if (transform?.rotate) {
						// 换算为数字角度，单位：弧度，注意可能产生负值，-1
						let angel = Math.PI * transform.rotate / 180;
						width = Math.abs(origin_width * Math.cos(angel) + origin_height * Math.sin(angel));
						height = Math.abs(origin_width * Math.sin(angel) + origin_height * Math.cos(angel));
					} else {
						// 无旋转
						width = origin_width;
						height = origin_height;
					}
					target.props.width = convertLength(width, LengthUsage.Px, false);
					target.props.height = convertLength(height, LengthUsage.Px, false);
					target.cssStyle["width"] = convertLength(width, LengthUsage.Emu, true);
					target.cssStyle["height"] = convertLength(height, LengthUsage.Emu, true);
					break;

				// 变换之后的偏移量，实际上无效
				case "off":
					target.cssStyle["left"] = xml.lengthAttr(n, "x", LengthUsage.Emu);
					target.cssStyle["top"] = xml.lengthAttr(n, "y", LengthUsage.Emu);
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Transform2D Element：${n.localName}`, 'color:#f75607');
					}
			}
		}
	}

	// 图像填充
	parseBlipFill(node: Element, target: WmlImage) {
		// 图像填充
		for (let n of xml.elements(node)) {
			switch (n.localName) {
				// 填充效果
				case "blip":
					// embed属性：图片地址
					target.src = xml.attr(n, "embed");
					// 图片填充效果
					this.parseBlip(n, target);
					break;
				// 源矩形裁剪
				case "srcRect":
					// 距离源图片的4方位间距，单位百分比（%）
					let left = xml.numberAttr(n, "l", LengthUsage.RelativeRect, 0);
					let right = xml.numberAttr(n, "r", LengthUsage.RelativeRect, 0);
					let top = xml.numberAttr(n, "t", LengthUsage.RelativeRect, 0);
					let bottom = xml.numberAttr(n, "b", LengthUsage.RelativeRect, 0);
					// 裁剪路径
					target.props.is_clip = [left, right, top, bottom].some((item) => item !== 0);
					target.props.clip.type = 'inset';
					target.props.clip.path = { top, right, bottom, left };
					break;
				case "stretch":
					break;
				// 平铺
				case "tile":
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Blip Fill Element：${n.localName}`, 'color:#f75607');
					}
			}
		}
	}

	// 图片填充效果
	parseBlip(node: Element, target: OpenXmlElement) {

		for (let n of xml.elements(node)) {
			switch (n.localName) {
				case "alphaBiLevel":

					break;
				case "alphaCeiling":

					break;
				case "alphaFloor":

					break;
				case "alphaInv":

					break;
				case "alphaMod":

					break;
				// 透明度
				case "alphaModFix":
					let opacity = xml.lengthAttr(n, 'amt', LengthUsage.Opacity);
					target.cssStyle["opacity"] = opacity;
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Blip Element：${n.localName}`, 'color:#f75607');
					}
					break;
			}
		}

	}

	parseTable(node: Element): WmlTable {
		let result: WmlTable = { type: DomType.Table, children: [] };

		xmlUtil.foreach(node, c => {
			switch (c.localName) {
				case "tblPr":
					this.parseTableProperties(c, result);
					break;

				case "tblGrid":
					result.columns = this.parseTableColumns(c);
					break;

				case "tr":
					result.children.push(this.parseTableRow(c));
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Table Element：${c.localName}`, 'color:#f75607');
					}
			}
		});

		return result;
	}

	parseTableColumns(node: Element): WmlTableColumn[] {
		let result = [];

		xmlUtil.foreach(node, n => {
			switch (n.localName) {
				case "gridCol":
					result.push({ width: xml.lengthAttr(n, "w") });
					break;

				// TODO 网格修订信息
				case "tblGridChange":
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Table Columns Element：${n.localName}`, 'color:#f75607');
					}
			}
		});

		return result;
	}

	parseTableProperties(elem: Element, table: WmlTable) {
		table.cssStyle = {};
		table.cellStyle = {};

		this.parseDefaultProperties(elem, table.cssStyle, table.cellStyle, c => {
			switch (c.localName) {
				case "tblStyle":
					table.styleName = xml.attr(c, "val");
					break;

				case "tblLook":
					table.className = values.classNameOftblLook(c);
					break;

				case "tblpPr":
					// 浮动表格位置
					this.parseTablePosition(c, table);
					break;

				case "tblStyleColBandSize":
					table.colBandSize = xml.intAttr(c, "val");
					break;

				case "tblStyleRowBandSize":
					table.rowBandSize = xml.intAttr(c, "val");
					break;

				default:
					// pass other properties to parseDefaultProperties function
					return false;
			}

			return true;
		});

		switch (table.cssStyle["text-align"]) {
			case "center":
				delete table.cssStyle["text-align"];
				table.cssStyle["margin-left"] = "auto";
				table.cssStyle["margin-right"] = "auto";
				break;

			case "right":
				delete table.cssStyle["text-align"];
				table.cssStyle["margin-left"] = "auto";
				break;

			default:
				if (this.options.debug) {
					console.warn(`DOCX:%c Unknown Table Align：${table.cssStyle["text-align"]}`, 'color:#f75607');
				}
		}
	}

	// 浮动表格，实现文字环绕
	parseTablePosition(node: Element, table: WmlTable) {
		// 由于浮动，导致后续元素错乱，默认忽略。
		if (this.options.ignoreTableWrap) {
			return false;
		}
		let topFromText = xml.lengthAttr(node, "topFromText");
		let bottomFromText = xml.lengthAttr(node, "bottomFromText");
		let rightFromText = xml.lengthAttr(node, "rightFromText");
		let leftFromText = xml.lengthAttr(node, "leftFromText");

		table.cssStyle["float"] = 'left';
		table.cssStyle["margin-bottom"] = values.addSize(table.cssStyle["margin-bottom"], bottomFromText);
		table.cssStyle["margin-left"] = values.addSize(table.cssStyle["margin-left"], leftFromText);
		table.cssStyle["margin-right"] = values.addSize(table.cssStyle["margin-right"], rightFromText);
		table.cssStyle["margin-top"] = values.addSize(table.cssStyle["margin-top"], topFromText);
	}

	parseTableRow(node: Element): WmlTableRow {
		let result: WmlTableRow = { type: DomType.Row, children: [] };

		xmlUtil.foreach(node, c => {
			switch (c.localName) {
				case "tc":
					result.children.push(this.parseTableCell(c));
					break;

				case "trPr":
					this.parseTableRowProperties(c, result);
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Table Row Element：${c.localName}`, 'color:#f75607');
					}
			}
		});

		return result;
	}

	parseTableRowProperties(elem: Element, row: WmlTableRow) {
		row.cssStyle = this.parseDefaultProperties(elem, {}, null, c => {
			switch (c.localName) {
				case "cnfStyle":
					row.className = values.classNameOfCnfStyle(c);
					break;
				// 	tblHeader attribute is boolean attribute
				case "tblHeader":
					row.isHeader = xml.boolAttr(c, "val", true);
					break;

				default:
					// pass other properties to parseDefaultProperties function
					return false;
			}

			return true;
		});
	}

	parseTableCell(node: Element): OpenXmlElement {
		let result: WmlTableCell = { type: DomType.Cell, children: [] };

		xmlUtil.foreach(node, c => {
			switch (c.localName) {
				case "tbl":
					result.children.push(this.parseTable(c));
					break;

				case "p":
					result.children.push(this.parseParagraph(c));
					break;

				case "tcPr":
					this.parseTableCellProperties(c, result);
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Table Cell Element：${c.localName}`, 'color:#f75607');
					}
			}
		});

		return result;
	}

	parseTableCellProperties(elem: Element, cell: WmlTableCell) {
		cell.cssStyle = this.parseDefaultProperties(elem, {}, null, c => {
			switch (c.localName) {
				case "gridSpan":
					cell.span = xml.intAttr(c, "val", null);
					break;

				case "vMerge":
					cell.verticalMerge = xml.attr(c, "val") ?? "continue";
					break;

				case "cnfStyle":
					cell.className = values.classNameOfCnfStyle(c);
					break;

				default:
					// pass other properties to parseDefaultProperties function
					return false;
			}

			return true;
		});
	}

	parseDefaultProperties(elem: Element, style: Record<string, string> = null, childStyle: Record<string, string> = null, handler: (prop: Element) => boolean = null): Record<string, string> {
		style = style || {};

		xmlUtil.foreach(elem, c => {
			/**
			 * 根据提供的handler处理函数和条件执行逻辑。
			 * 如果handler处理函数存在并且调用处理函数返回真值，则终止当前逻辑。
			 *
			 * @param handler 可选的处理函数，接受一个参数 c，并返回一个布尔值。
			 * @param c 传递给处理函数的参数。
			 */
			if (handler?.(c)) {
				return;
			}

			switch (c.localName) {
				// Bold
				case "b":
					style["font-weight"] = xml.boolAttr(c, "val", true) ? "bold" : "normal";
					break;

				// Complex Script Bold
				case "bCs":
					break;

				// Text Border
				case "bdr":
					style["border"] = values.valueOfBorder(c);
					break;

				// Display All Characters As Capital Letters
				case "caps":
					style["text-transform"] = xml.boolAttr(c, "val", true) ? "uppercase" : "none";
					break;

				// Run Content Color
				case "color":
					style["color"] = xmlUtil.colorAttr(c, "val", null, autos.color);
					break;

				// Use Complex Script Formatting on Run
				case "cs":
					break;

				// Double Strikethrough
				case "dstrike":
					break;

				// East Asian Typography Settings
				case "eastAsianLayout":
					break;

				// Animated Text Effect
				case "effect":
					break;

				// Emphasis Mark
				case "em":
					break;

				// Embossing
				case "emboss":
					break;

				// Manual Run Width
				case "fitText":
					break;

				// Text Highlighting
				case "highlight":
					style["background-color"] = xmlUtil.colorAttr(c, "val", null, autos.highlight);
					break;

				// Italics
				case "i":
					style["font-style"] = xml.boolAttr(c, "val", true) ? "italic" : "normal";
					break;

				// Complex Script Italics
				case "iCs":
					break;

				// Imprinting
				case "imprint":
					break;

				// TODO Font Kerning
				case "kern":
					//style['letter-spacing'] = xml.lengthAttr(elem, 'val', LengthUsage.FontSize);
					break;

				// Languages for Run Content,check spelling and grammar
				case "lang":
					style["$lang"] = xml.attr(c, "val");
					break;

				// TODO Do Not Check Spelling or Grammar
				case "noProof":
					break;

				// TODO Display Character Outline
				case "outline":
					break;

				// Vertically Raised or Lowered Text
				case "position":
					style.verticalAlign = xml.lengthAttr(c, "val", LengthUsage.FontSize);
					break;

				// Run Fonts
				case "rFonts":
					this.parseFont(c, style);
					break;

				// TODO Revision Information for Run Properties
				case "rPrChange":
					break;

				//TODO Right To Left Text
				case "rtl":
					break;

				//TODO Shadow
				case "shadow":
					break;

				// Run Shading
				case "shd":
					style["background-color"] = xmlUtil.colorAttr(c, "fill", null, autos.shd);
					break;

				// Small Caps
				case "smallCaps":
					style["font-variant"] = xml.boolAttr(c, "val", true) ? "small-caps" : "none";
					break;

				// Use Document Grid Settings For Inter-Character Spacing
				case "snapToGrid":
					break;

				// Character Spacing Adjustment
				case "spacing":
					// Paragraph
					if (elem.localName == "pPr") {
						this.parseSpacingBetweenLines(c, style);
					}
					// Run
					if (elem.localName == "rPr") {
						this.parseSpacing(c, style);
					}
					break;

				// Paragraph Mark Is Always Hidden
				case "specVanish":
					break;

				// Single Strikethrough
				case "strike":
					style["text-decoration"] = xml.boolAttr(c, "val", true) ? "line-through" : "none"
					break;

				// Non-Complex Script Font Size
				case "sz":
					// TODO 通过字符编码库或API来判断字符的编码范围，从而确定字符类型，字符类型决定字体大小
					style["font-size"] = style["min-height"] = xml.lengthAttr(c, "val", LengthUsage.FontSize);
					// style["font-size"] = xml.lengthAttr(c, "val", LengthUsage.FontSize);
					break;

				// Complex Script Font Size
				case "szCs":
					// TODO 通过字符编码库或API来判断字符的编码范围，从而确定字符类型，字符类型决定字体大小
					// style["font-size"] = style["min-height"] = xml.lengthAttr(c, "val", LengthUsage.FontSize);
					break;

				// Underline
				case "u":
					this.parseUnderline(c, style);
					break;

				// Hidden Text
				case "vanish":
					if (xml.boolAttr(c, "val", true))
						style["display"] = "none";
					break;

				// TODO	Subscript/Superscript Text
				case "vertAlign":
					// style.verticalAlign = values.valueOfVertAlign(c);
					break;

				// Expanded/Compressed Text
				case "w":
					break;

				// TODO Web Hidden Text
				case "webHidden":
					break;

				case "jc":
					style["text-align"] = values.valueOfJc(c);
					break;

				case "textAlignment":
					style["vertical-align"] = values.valueOfTextAlignment(c);
					break;

				case "tcW":
					if (this.options.ignoreWidth) {
					}
					break;

				case "tblW":
					style["width"] = values.valueOfSize(c, "w");
					break;

				case "trHeight":
					this.parseTrHeight(c, style);
					break;

				case "ind":
				case "tblInd":
					this.parseIndentation(c, style);
					break;

				case "tblBorders":
					this.parseBorderProperties(c, childStyle || style);
					break;

				case "tblCellSpacing":
					style["border-spacing"] = values.valueOfMargin(c);
					style["border-collapse"] = "separate";
					break;

				case "pBdr":
					this.parseBorderProperties(c, style);
					break;

				case "tcBorders":
					this.parseBorderProperties(c, style);
					break;

				case "noWrap":
					//TODO
					//style["white-space"] = "nowrap";
					break;

				case "tblCellMar":
				case "tcMar":
					this.parseMarginProperties(c, childStyle || style);
					break;

				case "tblLayout":
					style["table-layout"] = values.valueOfTblLayout(c);
					break;

				case "vAlign":
					style["vertical-align"] = values.valueOfTextAlignment(c);
					break;

				case "wordWrap":
					if (xml.boolAttr(c, "val")) //TODO: test with examples
						style["overflow-wrap"] = "break-word";
					break;

				case "suppressAutoHyphens":
					style["hyphens"] = xml.boolAttr(c, "val", true) ? "none" : "auto";
					break;

				case "tabs": //ignore - tabs is parsed by other parser
				case "outlineLvl": //TODO
				case "contextualSpacing": //TODO
				case "tblStyleColBandSize": //TODO
				case "tblStyleRowBandSize": //TODO
				case "pageBreakBefore": //TODO - maybe ignore
				case "suppressLineNumbers": //TODO - maybe ignore
				case "keepLines": //TODO - maybe ignore
				case "keepNext": //TODO - maybe ignore
				case "widowControl": //TODO - maybe ignore
				case "bidi": //TODO - maybe ignore

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Property Element：${elem.localName}.${c.localName}`, 'color:green');
					}
					break;
			}
		});

		return style;
	}

	parseUnderline(node: Element, style: Record<string, string>) {
		let val = xml.attr(node, "val");

		if (val == null)
			return;

		switch (val) {
			case "dash":
			case "dashDotDotHeavy":
			case "dashDotHeavy":
			case "dashedHeavy":
			case "dashLong":
			case "dashLongHeavy":
			case "dotDash":
			case "dotDotDash":
				style["text-decoration"] = "underline dashed";
				break;

			case "dotted":
			case "dottedHeavy":
				style["text-decoration"] = "underline dotted";
				break;

			case "double":
				style["text-decoration"] = "underline double";
				break;

			case "single":
			case "thick":
				style["text-decoration"] = "underline";
				break;

			case "wave":
			case "wavyDouble":
			case "wavyHeavy":
				style["text-decoration"] = "underline wavy";
				break;

			case "words":
				style["text-decoration"] = "underline";
				break;

			case "none":
				style["text-decoration"] = "none";
				break;

			default:
				if (this.options.debug) {
					console.warn(`DOCX:%c Unknown Underline Property：${val}`, 'color:#f75607');
				}
		}

		let col = xmlUtil.colorAttr(node, "color");

		if (col) {
			style["text-decoration-color"] = col;
		}

	}

	// 转换Run字体，包含四种，ascii，eastAsia，ComplexScript，高 ANSI Font
	// TODO 通过字符编码库或API来判断字符的编码范围，从而确定字符类型，字符类型决定字体大小
	parseFont(node: Element, style: Record<string, string>) {
		// 字体
		let fonts = new Set();
		// ascii字体
		let ascii = xml.attr(node, "ascii");
		let ascii_theme = values.themeValue(node, "asciiTheme");
		fonts.add(ascii).add(ascii_theme);
		// eastAsia
		let east_Asia = xml.attr(node, "eastAsia");
		let east_Asia_theme = values.themeValue(node, "eastAsiaTheme");
		fonts.add(east_Asia).add(east_Asia_theme);
		// ComplexScript
		let complex_script = xml.attr(node, "cs");
		let complex_script_theme = values.themeValue(node, "cstheme");
		fonts.add(complex_script).add(complex_script_theme);
		// 高 ANSI Font
		let high_ansi = xml.attr(node, "hAnsi");
		let high_ansi_theme = values.themeValue(node, "hAnsiTheme");
		fonts.add(high_ansi).add(high_ansi_theme);
		// 去除重复字体，去除null
		let unique_fonts = [...fonts].filter(x => x);
		if (unique_fonts.length > 0) {
			// 合并成一个字体配置
			style["font-family"] = unique_fonts.join(', ');
		}

		// 字体提示：hint，拥有三种值：ComplexScript（cs）、Default（default）、EastAsia（eastAsia）
		style["_hint"] = xml.attr(node, "hint");
	}

	parseIndentation(node: Element, style: Record<string, string>) {
		let firstLine = xml.lengthAttr(node, "firstLine");
		let hanging = xml.lengthAttr(node, "hanging");
		let left = xml.lengthAttr(node, "left");
		let start = xml.lengthAttr(node, "start");
		let right = xml.lengthAttr(node, "right");
		let end = xml.lengthAttr(node, "end");

		if (firstLine) style["text-indent"] = firstLine;
		if (hanging) style["text-indent"] = `-${hanging}`;
		// 段落缩进，通过padding实现
		if (left || start) style["padding-left"] = left || start;
		if (right || end) style["padding-right"] = right || end;
	}

	// the additional amount of character pitch to the contents of a run
	parseSpacing(node: Element, style: Record<string, string>) {
		for (const attr of xml.attrs(node)) {
			switch (attr.localName) {
				// Character Spacing Adjustment
				case "val":
					style["margin-bottom"] = xml.lengthAttr(node, "val");
					break;
				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Spacing Property：${attr.localName}`, 'color:#f75607');
					}
			}
		}
	}

	// Spacing Between Lines and Above/Below Paragraph
	parseSpacingBetweenLines(node: Element, style: Record<string, string>) {
		// line-height
		let line: number;

		for (const attr of xml.attrs(node)) {
			switch (attr.localName) {
				// Spacing after the last line in each paragraph
				case "after":
					style["margin-bottom"] = xml.lengthAttr(node, "after");
					break;

				// Automatically Determine Spacing after the last line in each paragraph
				case "afterAutospacing":
					break;

				// Spacing Below Paragraph in Line Units
				case "afterLines":
					style["margin-bottom"] = xml.lengthAttr(node, "afterLines");
					break;

				// Spacing before the first line in each paragraph
				case "before":
					style["margin-top"] = xml.lengthAttr(node, "before");
					break;

				// Automatically Determine Spacing before the first line in each paragraph
				case "beforeAutospacing":
					break;

				// Spacing Above Paragraph in Line Units
				case "beforeLines":
					style["margin-top"] = xml.lengthAttr(node, "beforeLines");
					break;

				//  the amount of vertical spacing between lines of text within this paragraph.
				case "line":
					line = xml.intAttr(node, "line", null);
					break;

				// Type of Spacing Between Lines
				case "lineRule":
					let lineRule = xml.attr(node, "lineRule");
					switch (lineRule) {
						// Automatically Determined Line Height.
						case "auto":
							style["line-height"] = `${(line / 240).toFixed(2)}`;
							break;

						// Minimum Line Height.
						case "atLeast":
							style["line-height"] = `calc(100% + ${line / 20}pt)`;
							break;

						// Exact Line Height.
						case "Exact":
							style["line-height"] = `${line / 20}pt`;
							break;

						default:
							style["line-height"] = style["min-height"] = `${line / 20}pt`
							break;
					}
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Spacing Property：${attr.localName}`, 'color:#f75607');
					}
			}
		}
	}

	parseMarginProperties(node: Element, output: Record<string, string>) {
		xmlUtil.foreach(node, c => {
			switch (c.localName) {
				case "left":
					output["padding-left"] = values.valueOfMargin(c);
					break;

				case "right":
					output["padding-right"] = values.valueOfMargin(c);
					break;

				case "top":
					output["padding-top"] = values.valueOfMargin(c);
					break;

				case "bottom":
					output["padding-bottom"] = values.valueOfMargin(c);
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Margin Property：${c.localName}`, 'color:#f75607');
					}
			}
		});
	}

	parseTrHeight(node: Element, output: Record<string, string>) {
		switch (xml.attr(node, "hRule")) {
			case "exact":
				output["height"] = xml.lengthAttr(node, "val");
				break;

			case "atLeast":
			default:
				output["height"] = xml.lengthAttr(node, "val");
				// min-height doesn't work for tr
				//output["min-height"] = xml.sizeAttr(node, "val");
				break;
		}
	}

	parseBorderProperties(node: Element, output: Record<string, string>) {
		xmlUtil.foreach(node, c => {
			switch (c.localName) {
				case "start":
				case "left":
					output["border-left"] = values.valueOfBorder(c);
					break;

				case "end":
				case "right":
					output["border-right"] = values.valueOfBorder(c);
					break;

				case "top":
					output["border-top"] = values.valueOfBorder(c);
					break;

				case "bottom":
					output["border-bottom"] = values.valueOfBorder(c);
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Border Property：${c.localName}`, 'color:#f75607');
					}
			}
		});
	}
}

const knownColors = ['black', 'blue', 'cyan', 'darkBlue', 'darkCyan', 'darkGray', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'green', 'lightGray', 'magenta', 'none', 'red', 'white', 'yellow'];

class xmlUtil {
	static foreach(node: Element, cb: (n: Element) => void) {
		for (let i = 0; i < node.childNodes.length; i++) {
			let n = node.childNodes[i];

			if (n.nodeType == Node.ELEMENT_NODE) {
				cb(<Element>n);
			}
		}
	}

	static colorAttr(node: Element, attrName: string, defValue: string = null, autoColor: string = 'black') {
		let v = xml.attr(node, attrName);

		if (v) {
			if (v == "auto") {
				return autoColor;
			} else if (knownColors.includes(v)) {
				return v;
			}

			return `#${v}`;
		}

		let themeColor = xml.attr(node, "themeColor");

		return themeColor ? `var(--docx-${themeColor}-color)` : defValue;
	}

	static sizeValue(node: Element, type: LengthUsageType = LengthUsage.Dxa): string {
		return convertLength(node.textContent, type) as string;
	}

	static parseTextContent(node: Element, defaultValue: number = 0): number {
		let textContent: string = node.textContent;
		return textContent ? parseInt(textContent) : defaultValue;
	}
}

// TODO 此处方法存在重复，XmlParser Class 中已存在类似的方法，需要统一
class values {
	static themeValue(c: Element, attr: string) {
		let val = xml.attr(c, attr);
		return val ? `var(--docx-${val}-font)` : null;
	}

	static valueOfSize(c: Element, attr: string) {
		let type = LengthUsage.Dxa;

		switch (xml.attr(c, "type")) {
			case "dxa":
				break;
			case "pct":
				type = LengthUsage.TablePercent;
				break;
			case "auto":
				return "auto";
		}

		return xml.lengthAttr(c, attr, type);
	}

	static valueOfMargin(c: Element) {
		return xml.lengthAttr(c, "w");
	}

	static valueOfBorder(c: Element) {
		let type = xml.attr(c, "val");

		if (type == "nil") {
			return "none";
		}

		let color = xmlUtil.colorAttr(c, "color");
		let size = xml.lengthAttr(c, "sz", LengthUsage.Border);

		return `${size} solid ${color == "auto" ? autos.borderColor : color}`;
	}

	static valueOfTblLayout(c: Element) {
		let type = xml.attr(c, "type");
		return type == "fixed" ? "fixed" : "auto";
	}

	static classNameOfCnfStyle(c: Element) {
		const val = xml.attr(c, "val");
		const classes = [
			'first-row', 'last-row', 'first-col', 'last-col',
			'odd-col', 'even-col', 'odd-row', 'even-row',
			'ne-cell', 'nw-cell', 'se-cell', 'sw-cell'
		];

		return classes.filter((_, i) => val[i] == '1').join(' ');
	}

	static valueOfJc(c: Element) {
		let type = xml.attr(c, "val");

		switch (type) {
			case "start":
			case "left":
				return "left";
			case "center":
				return "center";
			case "end":
			case "right":
				return "right";
			case "both":
				return "justify";
		}

		return type;
	}

	static valueOfVertAlign(c: Element, asTagName: boolean = false) {
		let type = xml.attr(c, "val");

		switch (type) {
			case "subscript":
				return "sub";
			case "superscript":
				return asTagName ? "sup" : "super";
		}

		return asTagName ? null : type;
	}

	static valueOfTextAlignment(c: Element) {
		let type = xml.attr(c, "val");

		switch (type) {
			case "auto":
			case "baseline":
				return "baseline";
			case "top":
				return "top";
			case "center":
				return "middle";
			case "bottom":
				return "bottom";
		}

		return type;
	}

	static addSize(a: string, b: string): string {
		if (a == null) return b;
		if (b == null) return a;

		return `calc(${a} + ${b})`; //TODO
	}

	static classNameOftblLook(c: Element) {
		const val = xml.hexAttr(c, "val", 0);
		let className = "";

		if (xml.boolAttr(c, "firstRow") || (val & 0x0020)) className += " first-row";
		if (xml.boolAttr(c, "lastRow") || (val & 0x0040)) className += " last-row";
		if (xml.boolAttr(c, "firstColumn") || (val & 0x0080)) className += " first-col";
		if (xml.boolAttr(c, "lastColumn") || (val & 0x0100)) className += " last-col";
		if (xml.boolAttr(c, "noHBand") || (val & 0x0200)) className += " no-hband";
		if (xml.boolAttr(c, "noVBand") || (val & 0x0400)) className += " no-vband";

		return className.trim();
	}
}
