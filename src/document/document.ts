import { OpenXmlElement } from "./dom";
import { Section, SectionProperties } from "./section";

export interface DocumentElement extends OpenXmlElement {
	sections: Section[];
    props: SectionProperties;
}
