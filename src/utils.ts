export function escapeClassName(className: string) {
	return className?.replace(/[ .]+/g, '-').replace(/[&]+/g, 'and').toLowerCase();
}

export function splitPath(path: string): [string, string] {
	let si = path.lastIndexOf('/') + 1;
	let folder = si == 0 ? "" : path.substring(0, si);
	let fileName = si == 0 ? path : path.substring(si);

	return [folder, fileName];
}

export function resolvePath(path: string, base: string): string {
	try {
		const prefix = "http://docx/";
		const url = new URL(path, prefix + base).toString();
		return url.substring(prefix.length);
	} catch {
		return `${base}${path}`;
	}
}

export function blobToBase64(blob: Blob): Promise<string> {
	return new Promise((resolve, reject) => {
		const reader = new FileReader();
		reader.onloadend = () => resolve(reader.result as string);
		reader.onerror = () => reject();
		reader.readAsDataURL(blob);
	});
}

export function parseCssRules(text: string): Record<string, string> {
	const result: Record<string, string> = {};

	for (const rule of text.split(';')) {
		const [key, val] = rule.split(':');
		result[key] = val;
	}

	return result
}

export function formatCssRules(style: Record<string, string>): string {
	return Object.entries(style).map((k, v) => `${k}: ${v}`).join(';');
}

// 转化为数组
export function asArray<T>(val: T | T[]): T[] {
	return Array.isArray(val) ? val : [val];
}

// 生成UUID
export function uuid(): string {
	if (typeof crypto === 'object') {
		if (typeof crypto.randomUUID === 'function') {
			// https://developer.mozilla.org/en-US/docs/Web/API/Crypto/randomUUID
			return crypto.randomUUID();
		}
		if (typeof crypto.getRandomValues === 'function' && typeof Uint8Array === 'function') {
			// https://stackoverflow.com/questions/105034/how-to-create-a-guid-uuid
			const callback = (c: any) => {
				const num = Number(c);
				return (num ^ (crypto.getRandomValues(new Uint8Array(1))[0] & (15 >> (num / 4)))).toString(16);
			};
			return '10000000-1000-4000-8000-100000000000'.replace(/[018]/g, callback);
		}
	}
	let timestamp = new Date().getTime();
	let perforNow = (typeof performance !== 'undefined' && performance.now && performance.now() * 1000) || 0;
	return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
		let random = Math.random() * 16;
		if (timestamp > 0) {
			random = (timestamp + random) % 16 | 0;
			timestamp = Math.floor(timestamp / 16);
		} else {
			random = (perforNow + random) % 16 | 0;
			perforNow = Math.floor(perforNow / 16);
		}
		return (c === 'x' ? random : (random & 0x3) | 0x8).toString(16);
	});
}
