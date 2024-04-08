import { ParagraphProperties } from "./paragraph";
import { RunProperties } from "./run";

export interface IDomStyle {
	aliases?: string[];
	autoRedefine?: boolean;
	basedOn?: string;
	cssName?: string;
	hidden?: boolean;
	id: string;
	isDefault?: boolean;
	linked?: string;
	locked: boolean;
	name?: string;
	next?: string;
	paragraphProps: ParagraphProperties;
	personal?: boolean;
	personalCompose?: boolean;
	personalReply?: boolean;
	primaryStyle?: boolean;
	rsid?: number;
	runProps: RunProperties;
	semiHidden?: boolean;
	styles: IDomSubStyle[];
	target: string;
	uiPriority?: number;
	unhideWhenUsed?: boolean;
}

export interface IDomSubStyle {
	target: string;
	mod?: string;
	values: Record<string, string>;
}
