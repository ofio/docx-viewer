import { WordDocument } from './word-document';

import { DocumentParser } from './document-parser';

// HTML Render Asynchronously
import { HtmlRenderer } from './html-renderer';

// HTML Render Synchronously
import { HtmlRendererSync } from "./html-renderer-sync";

export interface Options {
	className: string;                      //class name/prefix for default and document style classes
	inWrapper: boolean;                     //enables rendering of wrapper around document content

	ignoreWidth: boolean;                   //disables rendering width of page
	ignoreHeight: boolean;                  //disables rendering height of page
	ignoreFonts: boolean;                   //disables fonts rendering
	ignoreTableWrap: boolean;               //disables table's text wrap setting
	ignoreImageWrap: boolean;               //disables image text wrap setting
	ignoreLastRenderedPageBreak: boolean;   //disables page breaking on lastRenderedPageBreak elements
	breakPages: boolean;                    //enables page breaking on page breaks

	trimXmlDeclaration: boolean;            //if true, xml declaration will be removed from xml documents before parsing
	useBase64URL: boolean;                  //if true, images, fonts, etc. will be converted to base 64 URL, otherwise URL.createObjectURL is used

	renderHeaders: boolean;                 //enables headers rendering
	renderFooters: boolean;                 //enables footers rendering
	renderFootnotes: boolean;               //enables footnotes rendering
	renderEndnotes: boolean;                //enables endnotes rendering
	renderChanges: boolean;                 //enables experimental rendering of document changes (inserions/deletions)

	experimental: boolean;                  //enables experimental features (tab stops calculation)
	debug: boolean;                         //enables additional logging
}

export const defaultOptions: Options = {
	className: "docx",
	inWrapper: true,

	ignoreHeight: false,
	ignoreWidth: false,
	ignoreFonts: false,
	ignoreTableWrap: true,
	ignoreImageWrap: false,
	ignoreLastRenderedPageBreak: true,
	breakPages: true,

	trimXmlDeclaration: true,
	useBase64URL: false,

	renderHeaders: true,
	renderFooters: true,
	renderFootnotes: true,
	renderEndnotes: true,
	renderChanges: false,

	experimental: false,
	debug: false,
}

// Document Parser
export function parseAsync(data: Blob | any, userOptions: Partial<Options> = null): Promise<any> {
	// assign defaultOptions
	const ops = { ...defaultOptions, ...userOptions };
	// 加载blob对象，根据DocumentParser转换规则，blob对象 => Object对象
	return WordDocument.load(data, new DocumentParser(ops), ops);
}

// Document Render
export async function renderDocument(document: any, bodyContainer: HTMLElement, styleContainer?: HTMLElement, sync: boolean = true, userOptions?: Partial<Options>): Promise<any> {
	// assign defaultOptions
	const ops = { ...defaultOptions, ...userOptions };
	// HTML渲染器实例
	const renderer = sync ? new HtmlRendererSync() : new HtmlRenderer();
	// Object对象 => HTML标签
	await renderer.render(document, bodyContainer, styleContainer, ops);
}

// Render Synchronously
export async function renderSync(data: Blob | any, bodyContainer: HTMLElement, styleContainer: HTMLElement = null, userOptions: Partial<Options> = null): Promise<any> {
	// parse document data
	const doc = await parseAsync(data, userOptions);
	// render document
	await renderDocument(doc, bodyContainer, styleContainer, true, userOptions);

	return doc;
}

// Render Asynchronously
export async function renderAsync(data: Blob | any, bodyContainer: HTMLElement, styleContainer?: HTMLElement, userOptions?: Partial<Options>): Promise<any> {
	const doc = await parseAsync(data, userOptions);
	await renderDocument(doc, bodyContainer, styleContainer, false, userOptions);
	return doc;
}
