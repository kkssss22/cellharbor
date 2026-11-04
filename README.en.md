# CellHarbor · 格子助手

**Turn pasted text into a table. Turn repetitive edits into a few clicks.**

A beginner-friendly Excel helper with a Chinese interface: a guided paste-to-table wizard, 28 data tools, and 8 formula templates. Try the editable browser demo before installing an add-in.

[中文](README.md) · [Live demo](https://kkssss22.github.io/cellharbor/) · [Setup guide](docs/getting-started.md) · [Add-in manifest](https://kkssss22.github.io/cellharbor/manifest.xml)

## Features

**Paste-to-table wizard:** paste text, choose a literal cell separator and optional row separator, specify the number of rows, columns, or both, then preview row-major or column-major filling. Short tables are padded with blanks; overflowing an explicit capacity produces an error rather than dropping data. Excel output always goes to a new worksheet. The browser demo replaces its grid only after confirmation. This is literal text splitting, not a general CSV file parser.

| Category | Examples |
| --- | --- |
| Clean | Trim whitespace, remove spaces, normalize whitespace, remove invisible characters, full-width to half-width |
| Text | Uppercase, lowercase, title case, find and replace |
| Numbers | Text/number conversion, percentage text conversion, rounding, absolute value, sign reversal |
| Fill | Fill down, replace blanks with zero |
| Inspect | Duplicates, blanks, numbers stored as text, long identifiers, mixed column types |
| Extract | Descriptive statistics, unique rows, split a column, transpose, remove empty rows, add a sequence |

Transformations show a preview before applying changes and skip formula cells. In-place transformations have a conditional one-step undo: it will not overwrite target cells that have changed since the operation. New-sheet outputs are outside this undo scope. Extraction creates a new worksheet containing **values only**, not a copy of formulas, formatting, charts, or comments. Eight formula templates add examples that users can inspect and copy.

Numeric conversion preserves leading-zero identifiers and long values that cannot be represented safely. It cannot recover digits that Excel has already lost.

## Two modes

**Browser demo:** an editable synthetic grid with tab-separated paste support. It does not connect to a workbook or import `.xlsx` files.

**Excel add-in:** an Office.js task pane loaded through the XML manifest. Select a contiguous range in Excel, preview a tool, then apply it. Sideload availability depends on your Excel host and organization policy. This repository is not a Microsoft Marketplace listing, browser extension, or production installer. See the [setup guide](docs/getting-started.md).

## Privacy

No AI service, API key, telemetry, or application dataset-upload endpoint. Processing runs locally in the task pane or browser. Excel mode loads Office.js from Microsoft's CDN; the public application is served by GitHub Pages. Microsoft's own workbook storage and synchronization are separate from this project's behavior. [Privacy details](docs/privacy.md).

## Develop

Node.js 22 or newer:

```sh
git clone https://github.com/kkssss22/cellharbor.git
cd cellharbor
npm run dev
```

Open http://localhost:4173 for the browser demo. Native JavaScript modules require no third-party runtime or build dependencies.

```sh
npm test
npm run build
```

The build produces `dist/`. The local HTTP demo server does not install the add-in. The manifest points to the deployed HTTPS site. See [CONTRIBUTING.md](CONTRIBUTING.md) for development expectations.

Licensed under [MIT](LICENSE). CellHarbor is an independent project and is not affiliated with Microsoft.

## Validation scope

Core and Office bridge mock tests plus browser interaction checks passed. Real Excel sideload and host compatibility have not yet been verified. Test with sample data after sideloading. The grouping separator can start a new row or a new column. Browser outputs retain a source-table history; the current full table can be copied or downloaded as TSV.
