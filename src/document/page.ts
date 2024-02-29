import { OpenXmlElement } from "./dom";
import { SectionProperties } from "./section";
import { uuid } from "../utils";

export interface PageProps {
	// section属性
	sectProps?: SectionProperties,
	// 页面子元素
	elements: OpenXmlElement[],
	// 已分页标识
	isSplit?: boolean,
	// 是否第一页
	isFirstPage?: boolean;
	// 是否最后一页
	isLastPage?: boolean;
	// 顶层元素拆分索引
	breakIndex?: number;
	// 渲染所有内容的元素
	contentElement?: HTMLElement;
	// 溢出检测开关
	checkingOverflow?: boolean,
}

export class Page {
	pageId: string;
	// section属性
	sectProps?: SectionProperties;
	// 页面子元素
	elements: OpenXmlElement[];
	// 已分页标识
	isSplit: boolean;
	// 是否第一页
	isFirstPage?: boolean;
	// 是否最后一页
	isLastPage?: boolean;
	// 顶层元素拆分索引
	breakIndex?: number;
	// 渲染所有内容的元素
	contentElement?: HTMLElement;
	// 溢出检测开关，header/footer不检测
	checkingOverflow?: boolean;

	constructor({ sectProps, elements = [], isSplit = false, isFirstPage = false, isLastPage = false, breakIndex = 0, contentElement, checkingOverflow = false, }: PageProps) {
		this.pageId = uuid();
		this.sectProps = sectProps;
		this.elements = elements;
		this.isSplit = isSplit;
		this.isFirstPage = isFirstPage;
		this.isLastPage = isLastPage;
		this.breakIndex = breakIndex;
		this.contentElement = contentElement;
		this.checkingOverflow = checkingOverflow;
	}
}
