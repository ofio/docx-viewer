import { DomType, OpenXmlElement } from "./dom";
import { SectionProperties } from "./section";
import { uuid } from "../utils";
// Tree节点
export interface TreeNode extends OpenXmlElement {
	// 前一个兄弟指针
	prev?: TreeNode | null;
	// 下一个兄弟指针
	next?: TreeNode | null;
}

export interface PageProps {
	// section属性
	sectProps?: SectionProperties,
	// 页面子元素
	children: OpenXmlElement[],
	// 元素栈
	stack?: TreeNode[],
	// 已分页标识
	isSplit?: boolean,
	// 是否第一页
	isFirstPage?: boolean;
	// 是否最后一页
	isLastPage?: boolean;
	// 顶层元素拆分索引
	breakIndex?: Set<number>;
	// 渲染所有内容的元素
	contentElement?: HTMLElement;
	// 溢出检测开关
	checkingOverflow?: boolean,
}

export class Page implements OpenXmlElement {
	type: DomType;
	// id
	pageId: string;
	// section属性
	sectProps?: SectionProperties;
	// 页面子元素
	children: OpenXmlElement[];
	// 元素栈
	stack: TreeNode[];
	// 元素层级
	level?: number;
	// 已分页标识
	isSplit: boolean;
	// 是否第一页
	isFirstPage?: boolean;
	// 是否最后一页
	isLastPage?: boolean;
	// 顶层元素拆分索引
	breakIndex?: Set<number>;
	// 渲染所有内容的元素
	contentElement?: HTMLElement;
	// 溢出检测开关，header/footer不检测
	checkingOverflow?: boolean;

	constructor({ sectProps, children = [], stack = [], isSplit = false, isFirstPage = false, isLastPage = false, breakIndex = new Set(), contentElement, checkingOverflow = false, }: PageProps) {
		this.type = DomType.Page;
		this.level = 1;
		this.pageId = uuid();
		this.sectProps = sectProps;
		this.children = children;
		this.stack = stack;
		this.isSplit = isSplit;
		this.isFirstPage = isFirstPage;
		this.isLastPage = isLastPage;
		this.breakIndex = breakIndex;
		this.contentElement = contentElement;
		this.checkingOverflow = checkingOverflow;
	}

}
