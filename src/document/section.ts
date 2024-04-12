import globalXmlParser, { XmlParser } from "../parser/xml-parser";
import { Borders, parseBorders } from "./border";
import { Length, convertLength } from "./common";

export interface Column {
	space: Length;
	width: Length;
}

export interface Columns {
	space: Length;
	count: number;
	separator: boolean;
	equalWidth: boolean;
	columns: Column[];
}

export interface ContentSize {
	width: Length,
	height: Length,
}

export interface PageSize extends ContentSize {
	orientation: "landscape" | string
}

export interface PageNumber {
	start: number;
	chapSep: "colon" | "emDash" | "endash" | "hyphen" | "period" | string;
	chapStyle: string;
	format: "none" | "cardinalText" | "decimal" | "decimalEnclosedCircle" | "decimalEnclosedFullstop"
		| "decimalEnclosedParen" | "decimalZero" | "lowerLetter" | "lowerRoman"
		| "ordinalText" | "upperLetter" | "upperRoman" | string;
}

export interface PageMargins {
	top: Length;
	right: Length;
	bottom: Length;
	left: Length;
	header: Length;
	footer: Length;
	gutter: Length;
}

export enum SectionType {
	Continuous = "continuous",
	NextPage = "nextPage",
	NextColumn = "nextColumn",
	EvenPage = "evenPage",
	OddPage = "oddPage",
}

export interface FooterHeaderReference {
	id: string;
	type: string | "first" | "even" | "default";
}

export interface SectionProperties {
	sectionId: string;
	type: SectionType | string;
	pageSize: PageSize,
	pageMargins: PageMargins,
	pageBorders: Borders;
	pageNumber: PageNumber;
	columns: Columns;
	footerRefs: FooterHeaderReference[];
	headerRefs: FooterHeaderReference[];
	titlePage: boolean;
	contentSize: ContentSize;
}

// 原始尺寸数据，单位：dxa
interface OriginSize {
	pageSize: {
		width: number,
		height: number,
	},
	pageMargins: {
		top: number;
		right: number;
		bottom: number;
		left: number;
		header: number;
		footer: number;
		gutter: number;
	}
}

export function parseSectionProperties(elem: Element, xml: XmlParser = globalXmlParser): SectionProperties {
	let section = <SectionProperties>{};
	// 原始尺寸，单位：dxa
	let origin = <OriginSize>{};

	for (let e of xml.elements(elem)) {
		switch (e.localName) {
			case "pgSz":
				section.pageSize = {
					width: xml.lengthAttr(e, "w"),
					height: xml.lengthAttr(e, "h"),
					orientation: xml.attr(e, "orient")
				}
				// 记录原始尺寸
				origin.pageSize = {
					width: xml.intAttr(e, "w"),
					height: xml.intAttr(e, "h"),
				}
				break;

			case "type":
				section.type = xml.attr(e, "val");
				break;

			case "pgMar":
				section.pageMargins = {
					left: xml.lengthAttr(e, "left"),
					right: xml.lengthAttr(e, "right"),
					top: xml.lengthAttr(e, "top"),
					bottom: xml.lengthAttr(e, "bottom"),
					header: xml.lengthAttr(e, "header"),
					footer: xml.lengthAttr(e, "footer"),
					gutter: xml.lengthAttr(e, "gutter"),
				};
				// 记录原始尺寸
				origin.pageMargins = {
					left: xml.intAttr(e, "left"),
					right: xml.intAttr(e, "right"),
					top: xml.intAttr(e, "top"),
					bottom: xml.intAttr(e, "bottom"),
					header: xml.intAttr(e, "header"),
					footer: xml.intAttr(e, "footer"),
					gutter: xml.intAttr(e, "gutter"),
				}
				break;

			case "cols":
				section.columns = parseColumns(e, xml);
				break;

			case "headerReference":
				(section.headerRefs ?? (section.headerRefs = [])).push(parseFooterHeaderReference(e, xml));
				break;

			case "footerReference":
				(section.footerRefs ?? (section.footerRefs = [])).push(parseFooterHeaderReference(e, xml));
				break;

			case "titlePg":
				section.titlePage = xml.boolAttr(e, "val", true);
				break;

			case "pgBorders":
				section.pageBorders = parseBorders(e, xml);
				break;

			case "pgNumType":
				section.pageNumber = parsePageNumber(e, xml);
				break;

			// TODO 文档网格线
			case "docGrid":

				break;
			default:
				if (this.options.debug) {
					console.warn(`DOCX:%c Unknown Section Property：${elem.localName}`, 'color:#f75607');
				}
		}
	}

	// 根据原始尺寸，计算内容区域的宽高
	let { width, height } = origin.pageSize;
	let { left, right, top, bottom } = origin.pageMargins;

	section.contentSize = {
		width: convertLength(width - left - right) as string,
		height: convertLength(height - top - bottom) as string,
	}

	return section;
}

function parseColumns(elem: Element, xml: XmlParser): Columns {
	return {
		count: xml.intAttr(elem, "num"),
		space: xml.lengthAttr(elem, "space"),
		separator: xml.boolAttr(elem, "sep"),
		equalWidth: xml.boolAttr(elem, "equalWidth", true),
		columns: xml.elements(elem, "col")
			.map(e => <Column>{
				width: xml.lengthAttr(e, "w"),
				space: xml.lengthAttr(e, "space")
			})
	};
}

function parsePageNumber(elem: Element, xml: XmlParser): PageNumber {
	return {
		chapSep: xml.attr(elem, "chapSep"),
		chapStyle: xml.attr(elem, "chapStyle"),
		format: xml.attr(elem, "fmt"),
		start: xml.intAttr(elem, "start")
	};
}

function parseFooterHeaderReference(elem: Element, xml: XmlParser): FooterHeaderReference {
	return {
		id: xml.attr(elem, "id"),
		type: xml.attr(elem, "type"),
	}
}
