# Nedev.FileConverters.DocxToDoc

`Nedev.FileConverters.DocxToDoc` is a .NET library for converting OpenXML `.docx` documents to legacy binary `.doc` files without requiring Microsoft Word or Office automation. The repository contains the reusable converter library, a CLI wrapper, and an xUnit regression suite.

The core package targets `net8.0` and `netstandard2.1`. The CLI targets `net8.0`.

> Dependency note: the package depends on [`Nedev.FileConverters.Core`](https://www.nuget.org/packages/Nedev.FileConverters.Core) for shared infrastructure.

## Current Status

The implementation is no longer limited to plain text and basic formatting. In the current repository state, the converter supports:

- paragraph and run text serialization, including inline tabs, positioned tabs, symbol characters, hyphen control characters, and manual break characters
- character formatting through CHPX
- paragraph formatting through PAPX, including alignment, spacing, line spacing, and indent
- sections with page size and page margins
- numbering and list plumbing
- bookmarks, comment parsing with main-document anchor CP recovery and reply metadata, plus top-level plain-text comment emission in the writer with document-end anchor fallback, author-to-initials descriptor fallback, and paragraph-break preservation inside annotation stories
- core and extended document property parsing from `docProps/core.xml` and `docProps/app.xml`
- simple and nested field markers, plus hyperlink fields
- footnote and endnote parsing from `footnotes.xml` / `endnotes.xml`, main-document reference CP recovery, bounded DOC note-story emission for plain-text notes with preserved internal paragraph breaks, bounded separator / continuation-separator / continuation-notice special-story support, and bounded single-mark custom note reference support
- native DOC picture blocks in the `Data` stream, including image placeholders inside hyperlink fields
- OfficeArt/Escher output for PNG, JPEG, EMF, and WMF images
- anchored/floating image geometry with page, margin, and paragraph-relative positioning heuristics
- table rows, cell markers, row markers, and TAPX output
- table width inference from `tcW` (`dxa` and `pct`), `tblGrid/gridCol`, `gridSpan`, and cell padding-aware width reduction for layout heuristics
- explicit zero cell margin (`tcMar`) overrides that suppress table-level default cell padding during layout heuristics
- table cell top/bottom padding-aware vertical layout heuristics for row height and paragraph-relative floating content
- table border thickness heuristics from `tblBorders` and `tcBorders` affecting effective cell width and row advance
- table inside-border heuristics from `insideH` / `insideV` affecting internal row and column boundaries without double-counting adjacent cells
- table border conflict heuristics that prefer explicit cell borders over table-level inside borders on shared boundaries, including explicit `none` / `nil` suppression on shared edges
- table preferred width heuristics from `tblW` with `dxa`, `pct`, and `auto` behavior for scaling grid-based and explicit cell widths
- row-level mixed width allocation heuristics that reconcile `tcW`, `tblW`, `tblGrid`, `gridSpan`, and unresolved auto-width cells before wrapped-line estimation
- auto-width cell overhead reservation for horizontal padding, resolved borders, and cell spacing during remaining-width allocation
- overcommitted mixed-width shrink heuristics that reserve a minimum width for auto cells and shrink resolved widths first when explicit widths already exceed `tblW`
- mixed explicit/grid/auto overflow heuristics that prefer shrinking explicit-width cells before narrowing grid-fallback cells when unresolved auto cells still need reserved width
- mixed explicit overflow heuristics that preserve percentage-based `tcW` cells ahead of absolute `dxa` cells when the row still needs to reserve width for unresolved auto cells
- table row height heuristics from `trHeight` / `hRule` affecting minimum/exact row advance and in-row vertical alignment offsets
- exact row-height overflow clipping heuristics so later cell-local content does not keep advancing beyond a fixed-height row
- table cell vertical alignment (`top`/`center`/`bottom`) heuristics inside tall rows
- table cell spacing (`tblCellSpacing`) heuristics affecting effective cell width and row advance
- run-aware paragraph width estimation that uses per-run font size and a character-width heuristic instead of only raw character count

Current local validation status: `214/214` tests passing, and the package builds and packs successfully from the current workspace.

## Feature Coverage

| Area | Status | Notes |
|------|--------|-------|
| Paragraphs and runs | Implemented | Text, basic run formatting, hyperlink run formatting, inline tabs including `w:ptab`, symbol characters, line/paragraph spacing, indent, and run-aware width heuristics are serialized/applied. |
| Styles and fonts | Implemented | Style sheet and font table emission are present. |
| Sections | Implemented | Page size, margins, and page-number start metadata are supported. |
| Document properties | Implemented | Core and extended metadata are parsed from `docProps/core.xml` and `docProps/app.xml`. |
| Numbering and lists | Implemented | Abstract numbering and LFO/LST structures are emitted. |
| Fields and hyperlinks | Implemented | Simple-field and nested-field boundaries are modeled, and hyperlink instructions are emitted. |
| Footnotes and endnotes | Implemented with bounded fidelity | Plain-text note text and main-document reference CPs are parsed from `footnotes.xml` and `endnotes.xml`; multi-paragraph note text is preserved in the current string model, and the writer emits note references plus footnote/endnote stories and PLCFs. `separator`, `continuationSeparator`, and `continuationNotice` special stories are preserved and emitted through bounded header-story support, and single visible custom note marks are recovered and emitted through the current plain-text model, but multi-fragment custom marks and fuller note fidelity are not complete. |
| Images | Implemented with heuristics | Inline and floating images are written via DOC picture blocks and OfficeArt records, including hyperlink-wrapped image placeholders. |
| Tables | Implemented with heuristics | Table width/layout logic uses `tcW` including `pct` cell widths, `tblW`, `tblGrid`, `gridSpan`, row-level mixed width allocation for auto cells including horizontal overhead reservation for padding/borders/cell spacing, overcommitted-width shrink rules including explicit-before-grid shrink preference in mixed overflow rows and pct-before-dxa preservation in mixed explicit overflow rows, cell padding including explicit zero `tcMar` overrides, outer and inside border thickness, border conflict rules including explicit `none` / `nil` suppression on shared edges, row height hints, exact-height overflow clipping, cell spacing, and cell vertical alignment on both horizontal and vertical geometry paths, but is not a full Word layout engine. |
| Comments and bookmarks | Partial | Bookmarks are parsed and emitted in DOC structures; comments are parsed from `comments.xml` and `commentsExtended.xml` with main-document anchor CP recovery for ranges and collapsed references, plus reply/done metadata. Multi-paragraph comment text is preserved in the current string model, and the writer emits top-level plain-text comments into simplified DOC annotation structures, clamps document-end or overflow anchors to the final visible CP, and derives missing descriptor initials from comment authors, but reply/threaded comments and full annotation fidelity are not complete. |
| Advanced Word features | Partial / unsupported | SmartArt, equations, VBA/macros, tracked changes fidelity, and full layout parity are not complete. |

## Known Limits

This converter now covers a broad set of common Word constructs, but it still relies on layout heuristics in several places. The most important current limits are:

- paragraph height estimation is width-aware and run-aware, but still heuristic rather than font-metric exact
- floating image placement is substantially improved, but not a full Word-compatible layout engine
- table layout uses inferred and preferred widths, `tcW` percentage cell widths, mixed row-level width allocation including auto-cell horizontal overhead reservation, overcommitted-width shrink heuristics including explicit-before-grid overflow handling and pct-before-dxa preservation, padding including explicit zero cell-margin overrides, outer and inside border thickness, border conflict rules, row height hints, exact-height overflow clipping, cell spacing, row-height heuristics, and coarse cell vertical alignment behavior, but does not yet model every table rule Word applies
- footnotes and endnotes now preserve paragraph breaks in plain-text note stories, main-document reference recovery, bounded `separator` / `continuationSeparator` / `continuationNotice` special stories, and bounded single-mark custom note references, but multi-fragment custom markers and fuller DOC note fidelity are still bounded
- comments now preserve anchor CPs, reply metadata, and paragraph breaks in plain-text annotation stories; reply/threaded comments and full DOC annotation fidelity remain incomplete, while unsupported malformed anchors are filtered or clamped rather than fully modeled
- advanced Office features such as SmartArt, equations, macros, and exact compatibility behavior are not implemented

## Next Phase

The next practical phase should stay narrow and sample-driven rather than opening another broad layout front:

- only deepen footnote/endnote fidelity if real documents require multi-fragment custom-mark behavior or richer note formatting than the current paragraph-preserving plain-text model with bounded special-story support
- keep comments/bookmarks in their current bounded state unless a specific need justifies fuller DOC annotation fidelity
- revisit deeper table and layout parity only after concrete document samples show a gap that the current heuristics cannot cover

## Installation

Install from NuGet:

```powershell
dotnet add package Nedev.FileConverters.DocxToDoc --version 0.1.0
```

or reference it directly:

```xml
<PackageReference Include="Nedev.FileConverters.DocxToDoc" Version="0.1.0" />
```

## Library Usage

```csharp
using Nedev.FileConverters.DocxToDoc;

var converter = new DocxToDocConverter();
converter.Convert("input.docx", "output.doc");

using var input = File.OpenRead("input.docx");
using var output = File.Create("output.doc");
converter.Convert(input, output);
```

The converter throws `ArgumentNullException` for invalid path arguments and wraps conversion failures in converter-specific exceptions.

## CLI Usage

Build the CLI:

```powershell
cd src\Nedev.FileConverters.DocxToDoc.Cli
dotnet build -c Release
```

Run it from the build output:

```powershell
dotnet bin\Release\net8.0\Nedev.FileConverters.DocxToDoc.Cli.dll <input.docx> <output.doc>
```

Use `-h` or `--help` for help text.

## Development

Repository layout:

- `src/Nedev.FileConverters.DocxToDoc` - converter library
- `src/Nedev.FileConverters.DocxToDoc.Cli` - CLI wrapper
- `src/Nedev.FileConverters.DocxToDoc.Tests` - xUnit regression tests

Typical workflow:

```powershell
dotnet restore
dotnet test .\Nedev.FileConverters.DocxToDoc.sln
```

## Packaging

Create a NuGet package with:

```powershell
dotnet pack src\Nedev.FileConverters.DocxToDoc\Nedev.FileConverters.DocxToDoc.csproj -c Release -o nupkg
```

The package includes the repository `README.md` and `LICENSE`.

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.
