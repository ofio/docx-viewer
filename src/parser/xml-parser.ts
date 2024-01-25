import { Length, LengthUsage, LengthUsageType, convertLength, convertBoolean } from "../document/common";

export function parseXmlString(xmlString: string, trimXmlDeclaration: boolean = false): Document {
	if (trimXmlDeclaration)
		xmlString = xmlString.replace(/<[?].*[?]>/, "");

	xmlString = removeUTF8BOM(xmlString);

	const result = new DOMParser().parseFromString(xmlString, "application/xml");
	const errorText = hasXmlParserError(result);

	if (errorText)
		throw new Error(errorText);

	return result;
}

function hasXmlParserError(doc: Document) {
	return doc.getElementsByTagName("parsererror")[0]?.textContent;
}

function removeUTF8BOM(data: string) {
	return data.charCodeAt(0) === 0xFEFF ? data.substring(1) : data;
}

export function serializeXmlString(elem: Node): string {
	return new XMLSerializer().serializeToString(elem);
}

export class XmlParser {
	// 查找xml元素的多个子元素
	elements(elem: Element, localName: string = null): Element[] {
		const result = [];

		for (let i = 0, l = elem.childNodes.length; i < l; i++) {
			let c = elem.childNodes.item(i);

			if (c.nodeType == 1 && (localName == null || (c as Element).localName == localName))
				result.push(c);
		}

		return result;
	}

	element(elem: Element, localName: string): Element {
		for (let i = 0, l = elem.childNodes.length; i < l; i++) {
			let c = elem.childNodes.item(i);

			if (c.nodeType == 1 && (c as Element).localName == localName)
				return c as Element;
		}

		return null;
	}

	elementAttr(elem: Element, localName: string, attrLocalName: string): string {
		let el = this.element(elem, localName);
		return el ? this.attr(el, attrLocalName) : undefined;
	}

	attrs(elem: Element) {
		return Array.from(elem.attributes);
	}

	attr(elem: Element, localName: string): string {
		for (let i = 0, l = elem.attributes.length; i < l; i++) {
			let a = elem.attributes.item(i);

			if (a.localName == localName)
				return a.value;
		}

		return null;
	}

	intAttr(node: Element, attrName: string, defaultValue: number = null): number {
		let val = this.attr(node, attrName);
		return val ? parseInt(val) : defaultValue;
	}

	hexAttr(node: Element, attrName: string, defaultValue: number = null): number {
		let val = this.attr(node, attrName);
		return val ? parseInt(val, 16) : defaultValue;
	}

	floatAttr(node: Element, attrName: string, defaultValue: number = null): number {
		let val = this.attr(node, attrName);
		return val ? parseFloat(val) : defaultValue;
	}

	boolAttr(node: Element, attrName: string, defaultValue: boolean = null) {
		return convertBoolean(this.attr(node, attrName), defaultValue);
	}

	lengthAttr(node: Element, attrName: string, usage: LengthUsageType = LengthUsage.Dxa, defaultValue?: string): Length {
		let val = this.attr(node, attrName);
		return convertLength(val, usage) as Length ?? defaultValue;
	}

	numberAttr(node: Element, attrName: string, usage: LengthUsageType = LengthUsage.Dxa, defaultValue: number = 0): number {
		let val = this.attr(node, attrName);
		return convertLength(val, usage, false) as number ?? defaultValue;
	}
}

const globalXmlParser = new XmlParser();

export default globalXmlParser;
