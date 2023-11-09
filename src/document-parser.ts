import {DomType, WmlDrawing, IDomImage, IDomNumbering, NumberingPicBullet, OpenXmlElement, WmlBreak, WmlHyperlink, WmlNoteReference, WmlSymbol, WmlTable, WmlTableCell, WmlTableColumn, WmlTableRow, WmlText} from './document/dom';
import {DocumentElement} from './document/document';
import {parseParagraphProperties, parseParagraphProperty, WmlParagraph} from './document/paragraph';
import {parseSectionProperties, SectionProperties} from './document/section';
import xml from './parser/xml-parser';
import {parseRunProperties, WmlRun} from './document/run';
import {parseBookmarkEnd, parseBookmarkStart} from './document/bookmarks';
import {IDomStyle, IDomSubStyle} from './document/style';
import {WmlFieldChar, WmlFieldSimple, WmlInstructionText} from './document/fields';
import {convertLength, LengthUsage, LengthUsageType} from './document/common';
import {parseVmlElement} from './vml/vml';

export var autos = {
	shd: "inherit",
	color: "black",
	borderColor: "black",
	highlight: "transparent"
};

const supportedNamespaceURIs = [];

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

	parseDocumentFile(xmlDoc: Element): DocumentElement {
		let xbody = xml.element(xmlDoc, "body");
		let background = xml.element(xmlDoc, "background");
		let sectPr = xml.element(xbody, "sectPr");

		return {
			type: DomType.Document,
			children: this.parseBodyElements(xbody),
			props: sectPr ? parseSectionProperties(sectPr, xml) : {} as SectionProperties,
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
			}
		});

		return result;
	}

	parseDefaultStyles(node: Element): IDomStyle {
		let result = <IDomStyle>{
			id: null,
			name: null,
			target: null,
			basedOn: null,
			styles: []
		};

		xmlUtil.foreach(node, c => {
			switch (c.localName) {
				case "rPrDefault":
					let rPr = xml.element(c, "rPr");

					if (rPr)
						result.styles.push({
							target: "span",
							values: this.parseDefaultProperties(rPr, {})
						});
					break;

				case "pPrDefault":
					let pPr = xml.element(c, "pPr");

					if (pPr)
						result.styles.push({
							target: "p",
							values: this.parseDefaultProperties(pPr, {})
						});
					break;
			}
		});

		return result;
	}

	parseStyle(node: Element): IDomStyle {
		let result = <IDomStyle>{
			id: xml.attr(node, "styleId"),
			isDefault: xml.boolAttr(node, "default"),
			name: null,
			target: null,
			basedOn: null,
			styles: [],
			linked: null
		};

		switch (xml.attr(node, "type")) {
			case "paragraph":
				result.target = "p";
				break;
			case "table":
				result.target = "table";
				break;
			case "character":
				result.target = "span";
				break;
			//case "numbering": result.target = "p"; break;
		}

		xmlUtil.foreach(node, n => {
			switch (n.localName) {
				case "basedOn":
					result.basedOn = xml.attr(n, "val");
					break;

				case "name":
					result.name = xml.attr(n, "val");
					break;

				case "link":
					result.linked = xml.attr(n, "val");
					break;

				case "next":
					result.next = xml.attr(n, "val");
					break;

				case "aliases":
					result.aliases = xml.attr(n, "val").split(",");
					break;

				case "pPr":
					result.styles.push({
						target: "p",
						values: this.parseDefaultProperties(n, {})
					});
					result.paragraphProps = parseParagraphProperties(n, xml);
					break;

				case "rPr":
					result.styles.push({
						target: "span",
						values: this.parseDefaultProperties(n, {})
					});
					result.runProps = parseRunProperties(n, xml);
					break;

				case "tblPr":
				case "tcPr":
					result.styles.push({
						target: "td", //TODO: maybe move to processor
						values: this.parseDefaultProperties(n, {})
					});
					break;

				case "tblStylePr":
					for (let s of this.parseTableStyle(n))
						result.styles.push(s);
					break;

				case "rsid":
				case "qFormat":
				case "hidden":
				case "semiHidden":
				case "unhideWhenUsed":
				case "autoRedefine":
				case "uiPriority":
					//TODO: ignore
					break;

				default:
					this.options.debug && console.warn(`DOCX: Unknown style element: ${n.localName}`);
			}
		});

		return result;
	}

	parseTableStyle(node: Element): IDomSubStyle[] {
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
		let result = <WmlParagraph>{type: DomType.Paragraph, children: []};

		for (let el of xml.elements(node)) {
			switch (el.localName) {
				case "pPr":
					this.parseParagraphProperties(el, result);
					break;

				case "r":
					result.children.push(this.parseRun(el, result));
					break;

				case "hyperlink":
					result.children.push(this.parseHyperlink(el, result));
					break;

				case "bookmarkStart":
					result.children.push(parseBookmarkStart(el, xml));
					break;

				case "bookmarkEnd":
					result.children.push(parseBookmarkEnd(el, xml));
					break;

				case "oMath":
				case "oMathPara":
					result.children.push(this.parseMathElement(el));
					break;

				case "sdt":
					result.children.push(...this.parseSdt(el, (e: Element) => this.parseParagraph(e).children));
					break;

				case "ins":
					result.children.push(this.parseInserted(el, (e: Element) => this.parseParagraph(e)));
					break;

				case "del":
					result.children.push(this.parseDeleted(el, (e: Element) => this.parseParagraph(e)));
					break;
			}
		}
		// when paragraph is empty, a br tag needs to be added to work with the rich text editor and generate line height
		// 当段落children为空，需要添加一个br标签，配合富文本编辑器，同时产生行高
		if (result.children.length === 0) {
			let br: WmlBreak = {type: DomType.Break, "break": "textWrapping"};
			result.children = [br];
		}

		return result;
	}

	parseParagraphProperties(elem: Element, paragraph: WmlParagraph) {
		this.parseDefaultProperties(elem, paragraph.cssStyle = {}, null, c => {
			if (parseParagraphProperty(c, paragraph, xml))
				return true;

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
		let result: WmlHyperlink = <WmlHyperlink>{type: DomType.Hyperlink, parent: parent, children: []};
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
			}
		});

		return result;
	}

	parseRun(node: Element, parent?: OpenXmlElement): WmlRun {
		let result: WmlRun = <WmlRun>{type: DomType.Run, parent: parent, children: []};

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
					result.children.push({type: DomType.NoBreakHyphen});
					break;

				case "br":
					result.children.push(<WmlBreak>{
						type: DomType.Break,
						break: xml.attr(c, "type") || "textWrapping"
					});
					break;

				case "lastRenderedPageBreak":
					result.children.push(<WmlBreak>{
						type: DomType.Break,
						break: "lastRenderedPageBreak"
					});
					break;

				case "sym":
					result.children.push(<WmlSymbol>{
						type: DomType.Symbol,
						font: xml.attr(c, "font"),
						char: xml.attr(c, "char")
					});
					break;

				case "tab":
					result.children.push({type: DomType.Tab});
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
			}
		});

		return result;
	}

	parseMathElement(elem: Element): OpenXmlElement {
		const propsTag = `${elem.localName}Pr`;
		const result = {type: mmlTagMap[elem.localName], children: []} as OpenXmlElement;

		for (const el of xml.elements(elem)) {
			const childType = mmlTagMap[el.localName];

			if (childType) {
				result.children.push(this.parseMathElement(el));
			} else if (el.localName == "r") {
				let run = this.parseRun(el);
				run.type = DomType.MmlRun;
				result.children.push(run);
			} else if (el.localName == propsTag) {
				result.props = this.parseMathProperies(el);
			}
		}

		return result;
	}

	parseMathProperies(elem: Element): Record<string, any> {
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
			}
		}

		return result;
	}

	parseRunProperties(elem: Element, run: WmlRun) {
		this.parseDefaultProperties(elem, run.cssStyle = {}, null, c => {
			switch (c.localName) {
				case "rStyle":
					run.styleName = xml.attr(c, "val");
					break;

				case "vertAlign":
					run.verticalAlign = values.valueOfVertAlign(c, true);
					break;

				default:
					return false;
			}

			return true;
		});
	}

	parseVmlPicture(elem: Element): OpenXmlElement {
		const result = {type: DomType.VmlPicture, children: []};

		for (const el of xml.elements(elem)) {
			const child = parseVmlElement(el, this);
			child && result.children.push(child);
		}

		return result;
	}

	checkAlternateContent(elem: Element): Element {
		if (elem.localName != 'AlternateContent')
			return elem;

		let choice = xml.element(elem, "Choice");

		if (choice) {
			let requires = xml.attr(choice, "Requires");
			let namespaceURI = elem.lookupNamespaceURI(requires);

			if (supportedNamespaceURIs.includes(namespaceURI))
				return choice.firstElementChild;
		}

		return xml.element(elem, "Fallback")?.firstElementChild;
	}

	parseDrawing(node: Element): OpenXmlElement {
		for (let n of xml.elements(node)) {
			switch (n.localName) {
				case "inline":
				case "anchor":
					return this.parseDrawingWrapper(n);
			}
		}
	}

	parseDrawingWrapper(node: Element): OpenXmlElement {
		let result: WmlDrawing = {
			type: DomType.Drawing,
			children: [],
			cssStyle: {},
			localName: node.localName,
			wrapType: null
		};
		// DrawingML对象有两种状态：内联（inline）-- 对象与文本对齐，浮动（anchor）--对象在文本中浮动，但可以相对于页面进行绝对定位
		let isAnchor = node.localName === "anchor";

		//TODO 计算DrawML对象相对于文字的上下左右间距；
		// result.cssStyle["margin-left"] = xml.lengthAttr(node, "distL", LengthUsage.Emu);
		// result.cssStyle["margin-right"] = xml.lengthAttr(node, "distR", LengthUsage.Emu);
		result.cssStyle["margin-top"] = xml.lengthAttr(node, "distT", LengthUsage.Emu);
		result.cssStyle["margin-bottom"] = xml.lengthAttr(node, "distB", LengthUsage.Emu);

		// 是否简单定位
		let simplePos = xml.boolAttr(node, "simplePos");

		// 根据relativeHeight设置z-index
		result.cssStyle["z-index"] = xml.intAttr(node, "relativeHeight", 1);

		let posX = {relative: "page", align: "left", offset: "0"};
		let posY = {relative: "page", align: "top", offset: "0"};

		for (let n of xml.elements(node)) {
			switch (n.localName) {
				case "simplePos":
					if (simplePos) {
						posX.offset = xml.lengthAttr(n, "x", LengthUsage.Emu);
						posY.offset = xml.lengthAttr(n, "y", LengthUsage.Emu);
					}
					break;

				case "extent":
					result.cssStyle["width"] = xml.lengthAttr(n, "cx", LengthUsage.Emu);
					result.cssStyle["height"] = xml.lengthAttr(n, "cy", LengthUsage.Emu);
					break;

				case "positionH":
				case "positionV":
					if (!simplePos) {
						let pos = n.localName == "positionH" ? posX : posY;
						let alignNode = xml.element(n, "align");
						let offsetNode = xml.element(n, "posOffset");

						pos.relative = xml.attr(n, "relativeFrom") ?? pos.relative;

						if (alignNode)
							pos.align = alignNode.textContent;

						if (offsetNode)
							pos.offset = xmlUtil.sizeValue(offsetNode, LengthUsage.Emu);
					}
					break;

				case "wrapTopAndBottom":
					result.wrapType = "wrapTopAndBottom";
					break;

				case "wrapNone":
					result.wrapType = "wrapNone";
					break;

				case "graphic":
					let g = this.parseGraphic(n);

					if (g) {
						result.children.push(g);
					}
					break;
			}
		}
		// 图片文字环绕默认采用wrapTopAndBottom
		if (this.options.ignoreImageWrap) {
			result.wrapType = "wrapTopAndBottom";
		}

		switch (result.wrapType) {
			case "wrapTopAndBottom":
				// 顶部底部文字环绕
				result.cssStyle['display'] = 'block';

				if (posX.align) {
					result.cssStyle['text-align'] = posX.align;
					result.cssStyle['width'] = "100%";
				}
				break;
			case "wrapNone":
				// 衬于文字下方、浮于文字上方
				result.cssStyle['display'] = 'block';
				result.cssStyle['position'] = 'relative';
				result.cssStyle["width"] = "0px";
				result.cssStyle["height"] = "0px";

				if (posX.offset) {
					result.cssStyle["left"] = posX.offset;
				}

				if (posY.offset) {
					result.cssStyle["top"] = posY.offset;
				}
				break;
			case "wrapTight":
				// TODO 紧密型环绕
				break;
			case "wrapThrough":
				// TODO 穿越型环绕
				break;
			case "wrapSquare":
				// TODO 矩形环绕
				break;
			case "wrapPolygon":
				// TODO 多边形环绕
				break;
			default:
				// 默认四周文字环绕
				if (isAnchor && (posX.align == 'left' || posX.align == 'right')) {
					result.cssStyle["float"] = posX.align;
				}
		}

		return result;
	}

	parseGraphic(elem: Element): OpenXmlElement {
		let graphicData = xml.element(elem, "graphicData");

		for (let n of xml.elements(graphicData)) {
			switch (n.localName) {
				case "pic":
					return this.parsePicture(n);
			}
		}

		return null;
	}

	parsePicture(elem: Element): IDomImage {
		let result = <IDomImage>{type: DomType.Image, src: "", cssStyle: {}};
		let blipFill = xml.element(elem, "blipFill");
		let blip = xml.element(blipFill, "blip");

		result.src = xml.attr(blip, "embed");

		let spPr = xml.element(elem, "spPr");
		let xfrm = xml.element(spPr, "xfrm");

		// 图片旋转角度
		let degree = xml.lengthAttr(xfrm, "rot", LengthUsage.degree);
		if (degree) {
			result.cssStyle["transform"] = `rotate(${degree})`;
		}
		result.cssStyle["position"] = "relative";

		for (let n of xml.elements(xfrm)) {
			switch (n.localName) {
				case "ext":
					result.cssStyle["width"] = xml.lengthAttr(n, "cx", LengthUsage.Emu);
					result.cssStyle["height"] = xml.lengthAttr(n, "cy", LengthUsage.Emu);
					break;

				case "off":
					result.cssStyle["left"] = xml.lengthAttr(n, "x", LengthUsage.Emu);
					result.cssStyle["top"] = xml.lengthAttr(n, "y", LengthUsage.Emu);
					break;
			}
		}

		return result;
	}

	parseTable(node: Element): WmlTable {
		let result: WmlTable = {type: DomType.Table, children: []};

		xmlUtil.foreach(node, c => {
			switch (c.localName) {
				case "tr":
					result.children.push(this.parseTableRow(c));
					break;

				case "tblGrid":
					result.columns = this.parseTableColumns(c);
					break;

				case "tblPr":
					this.parseTableProperties(c, result);
					break;
			}
		});

		return result;
	}

	parseTableColumns(node: Element): WmlTableColumn[] {
		let result = [];

		xmlUtil.foreach(node, n => {
			switch (n.localName) {
				case "gridCol":
					result.push({width: xml.lengthAttr(n, "w")});
					break;
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
		let result: WmlTableRow = {type: DomType.Row, children: []};

		xmlUtil.foreach(node, c => {
			switch (c.localName) {
				case "tc":
					result.children.push(this.parseTableCell(c));
					break;

				case "trPr":
					this.parseTableRowProperties(c, result);
					break;
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
					return false;
			}

			return true;
		});
	}

	parseTableCell(node: Element): OpenXmlElement {
		let result: WmlTableCell = {type: DomType.Cell, children: []};

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
					return false;
			}

			return true;
		});
	}

	parseDefaultProperties(elem: Element, style: Record<string, string> = null, childStyle: Record<string, string> = null, handler: (prop: Element) => boolean = null): Record<string, string> {
		style = style || {};

		xmlUtil.foreach(elem, c => {
			if (handler?.(c))
				return;

			switch (c.localName) {
				case "jc":
					style["text-align"] = values.valueOfJc(c);
					break;

				case "textAlignment":
					style["vertical-align"] = values.valueOfTextAlignment(c);
					break;

				case "color":
					style["color"] = xmlUtil.colorAttr(c, "val", null, autos.color);
					break;

				case "sz":
					style["font-size"] = style["min-height"] = xml.lengthAttr(c, "val", LengthUsage.FontSize);
					// style["font-size"] = xml.lengthAttr(c, "val", LengthUsage.FontSize);
					break;

				case "shd":
					style["background-color"] = xmlUtil.colorAttr(c, "fill", null, autos.shd);
					break;

				case "highlight":
					style["background-color"] = xmlUtil.colorAttr(c, "val", null, autos.highlight);
					break;

				case "vertAlign":
					//TODO
					// style.verticalAlign = values.valueOfVertAlign(c);
					break;

				case "position":
					style.verticalAlign = xml.lengthAttr(c, "val", LengthUsage.FontSize);
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

				case "strike":
					style["text-decoration"] = xml.boolAttr(c, "val", true) ? "line-through" : "none"
					break;

				case "b":
					style["font-weight"] = xml.boolAttr(c, "val", true) ? "bold" : "normal";
					break;

				case "i":
					style["font-style"] = xml.boolAttr(c, "val", true) ? "italic" : "normal";
					break;

				case "caps":
					style["text-transform"] = xml.boolAttr(c, "val", true) ? "uppercase" : "none";
					break;

				case "smallCaps":
					style["text-transform"] = xml.boolAttr(c, "val", true) ? "lowercase" : "none";
					break;

				case "u":
					this.parseUnderline(c, style);
					break;

				case "ind":
				case "tblInd":
					this.parseIndentation(c, style);
					break;

				case "rFonts":
					this.parseFont(c, style);
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

				case "bdr":
					style["border"] = values.valueOfBorder(c);
					break;

				case "tcBorders":
					this.parseBorderProperties(c, style);
					break;

				case "vanish":
					if (xml.boolAttr(c, "val", true))
						style["display"] = "none";
					break;

				case "kern":
					//TODO
					//style['letter-spacing'] = xml.lengthAttr(elem, 'val', LengthUsage.FontSize);
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

				case "spacing":
					if (elem.localName == "pPr")
						this.parseSpacing(c, style);
					break;

				case "wordWrap":
					if (xml.boolAttr(c, "val")) //TODO: test with examples
						style["overflow-wrap"] = "break-word";
					break;

				case "suppressAutoHyphens":
					style["hyphens"] = xml.boolAttr(c, "val", true) ? "none" : "auto";
					break;

				case "lang":
					style["$lang"] = xml.attr(c, "val");
					break;

				case "bCs":
				case "iCs":
				case "szCs":
				case "tabs": //ignore - tabs is parsed by other parser
				case "outlineLvl": //TODO
				case "contextualSpacing": //TODO
				case "tblStyleColBandSize": //TODO
				case "tblStyleRowBandSize": //TODO
				case "webHidden": //TODO - maybe web-hidden should be implemented
				case "pageBreakBefore": //TODO - maybe ignore
				case "suppressLineNumbers": //TODO - maybe ignore
				case "keepLines": //TODO - maybe ignore
				case "keepNext": //TODO - maybe ignore
				case "widowControl": //TODO - maybe ignore
				case "bidi": //TODO - maybe ignore
				case "rtl": //TODO - maybe ignore
				case "noProof": //ignore spellcheck
					//TODO ignore
					break;

				default:
					if (this.options.debug)
						console.warn(`DOCX: Unknown document element: ${elem.localName}.${c.localName}`);
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
				style["text-decoration-style"] = "dashed";
				break;

			case "dotted":
			case "dottedHeavy":
				style["text-decoration-style"] = "dotted";
				break;

			case "double":
				style["text-decoration-style"] = "double";
				break;

			case "single":
			case "thick":
				style["text-decoration"] = "underline";
				break;

			case "wave":
			case "wavyDouble":
			case "wavyHeavy":
				style["text-decoration-style"] = "wavy";
				break;

			case "words":
				style["text-decoration"] = "underline";
				break;

			case "none":
				style["text-decoration"] = "none";
				break;
		}

		let col = xmlUtil.colorAttr(node, "color");

		if (col)
			style["text-decoration-color"] = col;
	}

	// 转换Run字体，包含四种，ascii，eastAsia，ComplexScript，高 ANSI Font
	parseFont(node: Element, style: Record<string, string>) {
		// 字体
		let fonts = [];
		// ascii字体
		let ascii = xml.attr(node, "ascii");
		let ascii_theme = values.themeValue(node, "asciiTheme");
		fonts.push(ascii, ascii_theme);
		// eastAsia
		let east_Asia = xml.attr(node, "eastAsia");
		let east_Asia_theme = values.themeValue(node, "eastAsiaTheme");
		fonts.push(east_Asia, east_Asia_theme);
		// ComplexScript
		let complex_script = xml.attr(node, "cs");
		let complex_script_theme = values.themeValue(node, "cstheme");
		fonts.push(complex_script, complex_script_theme);
		// 高 ANSI Font
		let high_ansi = xml.attr(node, "hAnsi");
		let high_ansi_theme = values.themeValue(node, "hAnsiTheme");
		fonts.push(high_ansi, high_ansi_theme);

		// 去除重复字体，合并成一个字体配置
		let fonts_value = [...new Set(fonts)].filter(x => x).join(', ');

		if (fonts.length > 0) {
			style["font-family"] = fonts_value;
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
		if (left || start) style["margin-left"] = left || start;
		if (right || end) style["margin-right"] = right || end;
	}

	parseSpacing(node: Element, style: Record<string, string>) {
		let before = xml.lengthAttr(node, "before");
		let after = xml.lengthAttr(node, "after");
		let line = xml.intAttr(node, "line", null);
		let lineRule = xml.attr(node, "lineRule");

		if (before) style["margin-top"] = before;
		if (after) style["margin-bottom"] = after;

		if (line !== null) {
			switch (lineRule) {
				case "auto":
					style["line-height"] = `${(line / 240).toFixed(2)}`;
					break;

				case "atLeast":
					style["line-height"] = `calc(100% + ${line / 20}pt)`;
					break;

				case "Exact":
					style["line-height"] = `${line / 20}pt`;
					break;

				default:
					style["line-height"] = style["min-height"] = `${line / 20}pt`
					break;
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

	static sizeValue(node: Element, type: LengthUsageType = LengthUsage.Dxa) {
		return convertLength(node.textContent, type);
	}
}

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
				type = LengthUsage.Percent;
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

		if (type == "nil")
			return "none";

		let color = xmlUtil.colorAttr(c, "color");
		let size = xml.lengthAttr(c, "sz", LengthUsage.Border);

		return `${size} solid ${color == "auto" ? autos.borderColor : color}`;
	}

	static valueOfTblLayout(c: Element) {
		let type = xml.attr(c, "val");
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
