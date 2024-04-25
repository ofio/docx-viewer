# docx-preview-sync
The docx document **synchronously** rendering library

[![npm version](https://badge.fury.io/js/docx-preview-sync.svg)](https://www.npmjs.com/package/docx-preview-sync)

Introduction
----
This library is inspired by the [docx-preview](https://github.com/VolodymyrBaydalka/docxjs) library. Thanks to the author [VolodymyrBaydalka](https://github.com/VolodymyrBaydalka) for his hard work.
I forked the project and modified it to support break pages as much as possible.
To achieve this goal, I changed the rendering library from asynchronous to synchronous, so that the rendering process can be completed in a synchronous manner, then I can detect all HTML elements that need to be broken and split them into different pages.

Based on the synchronous rendering process, This library will cause **lower** performance.

**This library is still in the development stage, so it is not recommended to use it in production.**

Goals
----
* Render/convert DOCX document into HTML document with keeping HTML semantic as much as possible.
* Support break pages strictly

That means this library is limited by HTML capabilities.If you need to render document on canvas,try the onlyOffice library.

Usage
-----
#### Package managers
Install library in your Node.js powered apps with the npm package:

```shell
npm install docx-preview-sync
```
```typescript
import { renderSync } from 'docx-preview-sync';

// fectch document Blob,maybe from input with type = file
let docData: Blob = document.querySelector('input').files[0];

// synchronously rendering function
let wordDocument = await renderSync(docData, document.getElementById("container"));

// if you need to get the word document object
console.log("docx document object", wordDocument);
```

#### Static HTML without a build step

```html
<!--dependencies-->
<script src="https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/lodash@4.17.21/lodash.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/konva@9.3.6/konva.min.js"></script>

<script src="./js/docx-preview.min.js"></script>

<script>
	// fectch document Blob,maybe from input width type = file
	let docData = document.querySelector('input').files[0];

	// synchronously rendering function
	docx.renderSync(docData, document.getElementById("container"))
            .then(wordDocument => {
                // if you need to get the Word document object
                console.log("docx document object", wordDocument);
            });
</script>
<body>
...
<div id="container"></div>
...
</body>
```

API
---

```typescript
// renders document into specified element
renderSync(
    document: Blob | ArrayBuffer | Uint8Array, // could be any type that supported by JSZip.loadAsync
    bodyContainer: HTMLElement, //element to render document content,
    styleContainer: HTMLElement, //element to render document styles, numbeings, fonts. If null, bodyContainer will be used.
    options: {
        className: string = "docx", //class name/prefix for default and document style classes
        inWrapper: boolean = true, //enables rendering of wrapper around document content
        ignoreWidth: boolean = false, //disables rendering width of page
        ignoreHeight: boolean = false, //disables rendering height of page
        ignoreFonts: boolean = false, //disables fonts rendering
        breakPages: boolean = true, //enables page breaking on page breaks
        ignoreLastRenderedPageBreak: boolean = true, //disables page breaking on lastRenderedPageBreak elements
        experimental: boolean = false, //enables experimental features (tab stops calculation)
        trimXmlDeclaration: boolean = true, //if true, xml declaration will be removed from xml documents before parsing
        useBase64URL: boolean = false, //if true, images, fonts, etc. will be converted to base 64 URL, otherwise URL.createObjectURL is used
        renderChanges: false, //enables experimental rendering of document changes (inserions/deletions)
        renderHeaders: true, //enables headers rendering
        renderFooters: true, //enables footers rendering
        renderFootnotes: true, //enables footnotes rendering
        renderEndnotes: true, //enables endnotes rendering
        debug: boolean = false, //enables additional logging
    }): Promise<WordDocument>

/// ==== experimental / internal API ===
// this API could be used to modify document before rendering
// renderSync = praseAsync + renderDocument

// parse document and return internal document object
praseAsync(
    document: Blob | ArrayBuffer | Uint8Array,
    options: Options
): Promise<WordDocument>

// render internal document object into specified container
renderDocument(
    wordDocument: WordDocument,
    bodyContainer: HTMLElement,
    styleContainer: HTMLElement,
    options: Options
): Promise<void>
```
Partially Supported Namespaces
------------------
1. [x] DocumentFormat.OpenXml.Wordprocessing
2. [x] DocumentFormat.OpenXml.Math

Not Supported Namespaces
------------------------
1. [ ] DocumentFormat.OpenXml.Drawing
2. [ ] DocumentFormat.OpenXml.Drawing.Charts
3. [ ] DocumentFormat.OpenXml.InkML
4. [ ] DocumentFormat.OpenXml.Vml

Breaks
------

Currently, library does break pages:

- if user/manual page break `<w:br w:type="page"/>` is inserted - when user insert page break
- if application page break `<w:lastRenderedPageBreak/>` is inserted - could be inserted by editor application like MS Word (`ignoreLastRenderedPageBreak` should be set to false)
- if page settings for paragraph is changed - ex: user change settings from portrait to landscape page

Realtime page breaking is not implemented because it's requires re-calculation of sizes on each insertion and that could affect performance a lot.

If page breaking is crutual for you, I would recommend:

- try to insert manual break point as much as you could
- try to use editors like MS Word, that inserts `<w:lastRenderedPageBreak/>` break points
