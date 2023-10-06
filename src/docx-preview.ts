import { WordDocument } from './word-document';

import { DocumentParser } from './document-parser';

import { HtmlRenderer } from './html-renderer';

export interface Options {
    className: string;                      //class name/prefix for default and document style classes
    inWrapper: boolean;                     //enables rendering of wrapper around document content
    ignoreWidth: boolean;                   //disables rendering width of page
    ignoreHeight: boolean;                  //disables rendering height of page
    ignoreFonts: boolean;                   //disables fonts rendering
    breakPages: boolean;                    //enables page breaking on page breaks
    ignoreLastRenderedPageBreak: boolean;   //disables page breaking on lastRenderedPageBreak elements
    experimental: boolean;                  //enables experimental features (tab stops calculation)
    trimXmlDeclaration: boolean;            //if true, xml declaration will be removed from xml documents before parsing
    useBase64URL: boolean;                  //if true, images, fonts, etc. will be converted to base 64 URL, otherwise URL.createObjectURL is used
    useMathMLPolyfill: boolean;             //@deprecated includes MathML polyfills for chrome, edge, etc.
    renderChanges: boolean;                 //enables experimental rendering of document changes (inserions/deletions)
    renderHeaders: boolean;                 //enables headers rendering
    renderFooters: boolean;                 //enables footers rendering
    renderFootnotes: boolean;               //enables footnotes rendering
    renderEndnotes: boolean;                //enables endnotes rendering
    debug: boolean;                         //enables additional logging
}

export const defaultOptions: Options = {
    ignoreHeight: false,
    ignoreWidth: false,
    ignoreFonts: false,
    breakPages: true,
    debug: false,
    experimental: false,
    className: "docx",
    inWrapper: true,
    trimXmlDeclaration: true,
    ignoreLastRenderedPageBreak: true,
    renderHeaders: true,
    renderFooters: true,
    renderFootnotes: true,
    renderEndnotes: true,
    useBase64URL: false,
    useMathMLPolyfill: false,
    renderChanges: false
}

export function praseAsync(data: Blob | any, userOptions: Partial<Options> = null): Promise<any> {
    const ops = { ...defaultOptions, ...userOptions };
    return WordDocument.load(data, new DocumentParser(ops), ops);
}

export async function renderAsync(data: Blob | any, bodyContainer: HTMLElement, styleContainer: HTMLElement = null, userOptions: Partial<Options> = null): Promise<any> {
    const ops = { ...defaultOptions, ...userOptions };
    // HTML渲染器实例
    const renderer = new HtmlRenderer(window.document);
    // 加载blob对象，根据DocumentParser转换规则，blob对象 => Object对象
    const doc = await WordDocument.load(data, new DocumentParser(ops), ops)
    // Object对象 => HTML标签
    await renderer.render(doc, bodyContainer, styleContainer, ops);

    return doc;
}

export async function docx2html(data: Blob | any, styleContainer: HTMLElement = null, userOptions: Partial<Options> = null): Promise<any> {
    const ops = { ...defaultOptions, ...userOptions };
    const renderer = new HtmlRenderer(window.document);
    const doc = await WordDocument.load(data, new DocumentParser(ops), ops)

    return renderer.renderFragment(doc, styleContainer, ops);
}