import { OpenXmlElement } from "./dom";
import { SectionProperties } from "./section";
import { uuid } from "../utils";

export interface PageProps {
	sectProps?: SectionProperties,
	elements: OpenXmlElement[],
	isSplit?: boolean,
	isFirstPage?: boolean;
	isLastPage?: boolean;
	elementIndex?: number;
	contentElement?: HTMLElement;
	checkingOverflow?: boolean,
}

export class Page {
	pageId: string;
	sectProps: SectionProperties;
	elements: OpenXmlElement[];
	isSplit: boolean;
	isFirstPage?: boolean;
	isLastPage?: boolean;
	elementIndex?: number;
	contentElement?: HTMLElement;
	checkingOverflow?: boolean;

	constructor({ sectProps, elements = [], isSplit = false, isFirstPage = false, isLastPage = false, elementIndex = 0, contentElement, checkingOverflow = false, }: PageProps) {
		this.pageId = uuid();
		this.sectProps = sectProps;
		this.elements = elements;
		this.isSplit = isSplit;
		this.isFirstPage = isFirstPage;
		this.isLastPage = isLastPage;
		this.elementIndex = elementIndex;
		this.contentElement = contentElement;
		this.checkingOverflow = checkingOverflow;
	}
}
