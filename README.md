# VBA-Linker

> **Lead paragraph (TODO — your input):** Write 2–4 sentences describing what VBA-Linker is and the pain it solves. Two framings to consider:
>
> 1. **Problem-first** — "Tired of manually re-pasting Excel charts into PowerPoint every time the model updates? VBA-Linker..."
> 2. **Capability-first** — "VBA-Linker is a paired Excel/PowerPoint addin that links charts, ranges, and cell values..."
>
> Once written, delete this blockquote.

VBA-Linker ships as two paired Office addins built from a shared VBA codebase:

- `xlVizer.xlam` — Excel addin (push side)
- `xlVizer.ppam` — PowerPoint addin (pull / refresh side)

The addin filename is `xlVizer` for historical reasons; the project, repo, and ribbon tab use the **VBA-Linker** name.

## Features

- **Send Chart** — push an Excel chart to the active PowerPoint slide as a linked PNG/EMF snapshot.
- **Send Range** — push a cell range as a linked image, or a single cell as a linked text box / inline span.
- **Batch Send** — send every linked item on the active sheet (or workbook) in one pass.
- **Refresh All** — re-pull every linked snapshot in the active deck from its source workbook(s).
- **Link Manager** — list every link in the deck with health status (`OK`, `STALE`, `BROKEN`), repair, remove, and go-to-slide actions.
- **Toggle Style** — flash all linked spans in the deck in a highlight color, or restore each span's captured original color.
- **Stale-edit detection** — if you hand-edit the value of a linked span in PowerPoint, refresh skips it and the Link Manager surfaces it as `STALE` so your edit is preserved.
- **Path resilience** — mapped drive letters resolve to UNC, OneDrive/SharePoint URLs resolve to local sync paths, and a fingerprint fallback finds source workbooks by filename + sheet when paths break.

## How it works

VBA-Linker stores link metadata directly on the PowerPoint side:

| Link type | Where the metadata lives |
|---|---|
| Chart, Range, Text-box | PowerPoint `Shape.Tags` collection, keys prefixed `xlVizer_*` |
| Inline span (linked text inside a paragraph) | A `LNK_NNN` shape tag, anchored to two zero-width-space (U+200B) markers wrapping the linked characters |

Each tag carries the source workbook path (normalized), sheet, address, and (for spans) the last value xlVizer wrote — so refresh can detect user hand-edits and back off.

**Flow ownership:**

- Excel = **push** only (Send Chart / Send Range / Batch Send).
- PowerPoint = **pull** owner (Refresh All, Link Manager, source-workbook reopening).

## Requirements

- Microsoft Office 2016+ (Excel **and** PowerPoint, Windows only)
- *Trust access to the VBA project object model* enabled in both Excel and PowerPoint Trust Center
- PowerShell 5.1+ (ships with Windows 10/11) — only needed if you build from source

## Status

This repository is the public face of an actively-developed internal project (currently at `v0.3.2`). Source code, build scripts, and pre-built `.xlam` / `.ppam` releases will be published here progressively.

## License

[MIT](LICENSE) — © 2026 Rory Sullivan
