# html-docx-kit

English | [简体中文](README.md)

A TypeScript library for converting HTML to Word DOCX in Node.js and modern browsers. It follows the **altChunk + MHTML** approach used by [html-docx-js](https://github.com/evidenceprime/html-docx-js). Microsoft Word performs the actual HTML conversion when it opens the document.

Install:

```sh
npm install html-docx-kit
```

## Node.js

```ts
import { writeFileSync } from 'node:fs';
import { asBuffer } from 'html-docx-kit/node';

const html = `<!doctype html>
<html><head><style>
  body { font-family: Arial, sans-serif; }
  h1 { color: #245da8; }
  table { border-collapse: collapse; }
  td, th { border: 1px solid #aaa; padding: 6px; }
</style></head><body>
  <h1>Project report</h1>
  <p>Supports English, Chinese, headings, lists, tables, and inline styles.</p>
  <table><tr><th>Item</th><th>Status</th></tr>
    <tr><td>HTML export</td><td>Done</td></tr></table>
</body></html>`;

writeFileSync('report.docx', asBuffer(html, {
  pageSize: 'A4',
  orientation: 'portrait',
  margins: { top: 720, bottom: 720 },
  properties: { title: 'Project report', creator: 'Product team' },
}));
```

CommonJS is also supported:

```js
const { asBuffer } = require('html-docx-kit/node');
require('node:fs').writeFileSync('report.docx', asBuffer('<h1>Hello</h1>'));
```

## Browser

Import the package through a bundler:

```ts
import { asBlob } from 'html-docx-kit';

const blob = asBlob('<h1>Hello, Word</h1><p>HTML fragments work too.</p>');
const url = URL.createObjectURL(blob);
const link = document.createElement('a');
link.href = url;
link.download = 'report.docx';
document.body.appendChild(link);
link.click();
link.remove();
setTimeout(() => URL.revokeObjectURL(url), 1000);
```

Without a bundler, load `dist/html-docx-kit.global.js` and call `htmlDocx.asBlob(html, options)` through the global `htmlDocx` object. Open `examples/browser.html` for a demo. Generation runs locally in the browser.

## API

| API | Return value | Purpose |
| --- | --- | --- |
| `htmlToDocx(html, options?)` | `Uint8Array` | General entry point; synchronously generates a DOCX |
| `asBlob(html, options?)` | `Blob` | Browser download; always returns a Blob in modern Node.js too |
| `asBuffer(html, options?)` | `Buffer` | Import only from `html-docx-kit/node` |
| `DOCX_MIME_TYPE` | `string` | Standard DOCX MIME type |

The main entry point also has a default export `{ htmlToDocx, asBlob }`. All output functions are synchronous and do not require `await`. The input may be a complete HTML document or a fragment; the parser fills in `html`, `head`, and `body` as needed and normalizes the document to UTF-8. An empty string throws an error.

| Option | Default | Description |
| --- | --- | --- |
| `orientation` | `'portrait'` | `'portrait'` or `'landscape'` |
| `pageSize` | `'letter'` | `'letter'`, `'A4'`, or `{ width, height }`; custom dimensions are in twips, and orientation is determined by the longer and shorter sides |
| `margins` | See below | Override any subset of the page margins |
| `externalImages` | `'error'` | External or relative image URLs throw by default; `'preserve'` keeps the URL as-is without downloading or embedding it |
| `inlineStyles` | `true` | Inline static rules from `<style>` elements before export; `false` keeps stylesheets without inlining them |
| `compression` | `6` | ZIP compression level, an integer from `0` to `9` |
| `properties` | Empty strings | `{ title?, creator?, description? }`, written to the document metadata with XML escaping |

Length units match the reference project: **1 inch = 1440 twips, 1 point = 20 twips**. Default margins are `top/right/bottom/left: 1440`, `header/footer: 720`, and `gutter: 0`. Dimensions and margins must be integers, and margins must leave a valid content area. `header` and `footer` set the distances for headers and footers; they do not create header or footer content.

## Images

Place images in `<img src="data:image/png;base64,...">`. The parser supports single-quoted, double-quoted, and unquoted attributes; case-insensitive tag names; data URIs with parameters; and percent-encoded image data URIs. Images with identical MIME types and bytes are packaged only once.

```ts
const html = `<p>Diagram:</p><img src="data:image/png;base64,${base64}" width="200">`;
const bytes = htmlToDocx(html);
```

The Word importer determines how images such as PNG, JPEG, and GIF are displayed. The library packages `image/*` data but does not convert formats such as SVG or WebP to PNG. The caller must download remote images and convert them to data URIs first; in browsers, the image server's CORS policy applies.

An absent or empty `src` does not create an image part. A non-empty `srcset` throws an error; choose one image and put it in `src` first. Sources in `<picture>`, CSS background images, external stylesheets, and similar resources are not packaged automatically. Prefer `<img src>` and inline `<style>` elements.

## Implementation and compatibility

```text
HTML → static CSS inlining → parse5 parsing/document normalization → image extraction → UTF-8/base64 MHTML
     → OOXML relationships and page settings → fflate ZIP → .docx
                                                           ↓ opened in Word
                                                   converted to Word content
```

The generated package contains `[Content_Types].xml`, the root relationships, `word/document.xml`, `word/_rels/document.xml.rels`, `word/afchunk.mht`, and `docProps/core.xml`. The `w:altChunk` relationship uses the case-sensitive `aFChunk` type required by Microsoft Word; the MHTML type is `message/rfc822`. See the [Microsoft specification](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/c391c28f-1b03-4a21-a4f8-4d9cddd4a95c).

HTML and images are fully base64-encoded and folded to 76 characters with CRLF line endings as specified by [RFC 2045](https://www.rfc-editor.org/rfc/rfc2045). MHTML resource references follow [RFC 2557](https://www.rfc-editor.org/rfc/rfc2557). A fixed ZIP timestamp makes identical input and options produce identical bytes.

- The target is Microsoft Word with `altChunk` support. The reference project documents that LibreOffice and Google Docs do not support this approach; older Mac Word 2008 does not support it either. Compatibility with other office applications is not guaranteed.
- By default, the synchronous Juice client inlines static CSS into element styles. This helps avoid incomplete Word imports of multiple classes and descendant selectors. Word determines which CSS properties it supports; flexbox, grid, complex web layouts, and browser visual effects are not guaranteed. The library does not request external stylesheets, execute scripts, or calculate styles after page scripts run.
- This is an `altChunk` document wrapper, not a converter that maps every HTML node to native WordprocessingML. For broad office-suite compatibility, a native OOXML conversion engine would be needed.
- The converter does not execute JavaScript in HTML and is not an HTML sanitizer. It preserves the supplied content; callers can sanitize it first if their application requires that.
- Node.js 18.17 or later is required; Node.js 20 is supported. Browsers need a modern environment with `atob`; `asBlob` also requires Blob support. Processing is synchronous and occupies the calling thread, so large documents can be handled in a Worker.

## Table cell padding and vertical alignment

Style inlining is enabled by default so descendant rules can apply cell padding when Word imports combinations such as `class="class-table small-table"`. More styles may affect row heights and pagination. To keep the original stylesheet without inlining, pass `{ inlineStyles: false }`.

Padding and vertical alignment are separate settings. To center content vertically, set both the cell's `vertical-align` and the paragraph's `margin` explicitly:

```html
<td style="padding:4pt; vertical-align:middle; background-color:#ffe4e6;">
  <p style="margin:0;">Content with padding and vertical centering</p>
</td>
```

`examples/table-padding/padding.html` demonstrates multiple class selectors, three background styles, and top/middle/bottom alignment. Run `node examples/table-padding/export.mjs` to generate two comparison DOCX files with inlining enabled and disabled. The library preserves alignment from the source styles and does not force all content to be centered.

## Development and verification

```sh
npm ci
npm run typecheck
npm test
npm run example
npm pack
```

`dist` contains ESM, CommonJS, browser IIFE, and `.d.ts`/`.d.cts` type declarations. Runtime code is bundled, so no runtime dependencies need to be installed separately. The project retains its source, build dependencies, and lockfile.

The tests check ZIP and OOXML relationships, MIME decoding, Chinese characters and emoji, image bytes and deduplication, page options, error handling, ESM/CJS, and the browser bundle without Node globals. Format checks cannot replace layout review in a real copy of Word; use `examples/sample.docx` to check the target Word version.

The 42 automated tests and the TypeScript build have passed. The npm tarball was installed offline in a separate project and checked with ESM/CJS, Blob/Buffer, and NodeNext type declarations. On macOS Microsoft Word 16.113.4, multi-class tables were verified for 4 pt padding on all four sides, zero paragraph spacing, and top/middle/bottom alignment; see `examples/table-padding/TEST-RESULTS.md`. Chinese text, headings, lists, tables, and embedded PNGs were also checked in that Word version. Other Word versions should be checked for the intended use case.

## Relationship to html-docx-js

This project follows its technical approach but reimplements the code in TypeScript; it does not copy the reference project's CoffeeScript source. Compared with the reference project, it adds type declarations, explicit output types, HTML fragment support, standard MIME encoding, more complete image attribute parsing, image deduplication, A4/custom page sizes, parameter validation, and reproducible output.

MIT licensed. Licenses for bundled dependencies are listed in `THIRD_PARTY_NOTICES.txt`.
