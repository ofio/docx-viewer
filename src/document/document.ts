import { OpenXmlElement } from "./dom";
import { SectionProperties } from "./section";
import { Page } from "./page";

export interface DocumentElement extends OpenXmlElement {
	pages: Page[];
    props: SectionProperties;
}
