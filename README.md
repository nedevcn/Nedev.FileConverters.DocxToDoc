# Nedev.FileConverters.DocxToDoc

`Nedev.FileConverters.DocxToDoc` is a .NET library for converting OpenXML `.docx` documents to legacy binary `.doc` files without requiring Microsoft Word or Office automation. The repository contains the reusable converter library, a CLI wrapper, and an xUnit regression suite.

The core package targets `net8.0` and `netstandard2.1`. The CLI targets `net8.0`.

> Dependency note: the package depends on [`Nedev.FileConverters.Core`](https://www.nuget.org/packages/Nedev.FileConverters.Core) for shared infrastructure.

## Current Status

The implementation is no longer limited to plain text and basic formatting. In the current repository state, the converter supports:

- paragraph and run text serialization
- character formatting through CHPX
- paragraph formatting through PAPX, including alignment, spacing, line spacing, and indent
- sections with page size and page margins
- numbering and list plumbing
- bookmarks and comments
- nested field markers and hyperlink fields
- native DOC picture blocks in the `Data` stream
- OfficeArt/Escher output for PNG, JPEG, EMF, and WMF images
- anchored/floating image geometry with page, margin, and paragraph-relative positioning heuristics
- table rows, cell markers, row markers, and TAPX output
- table width inference from `tcW`, `tblGrid/gridCol`, `gridSpan`, and cell padding-aware width reduction for layout heuristics
- table cell top/bottom padding-aware vertical layout heuristics for row height and paragraph-relative floating content
- table border thickness heuristics from `tblBorders` and `tcBorders` affecting effective cell width and row advance
- table row height heuristics from `trHeight` / `hRule` affecting minimum/exact row advance and in-row vertical alignment offsets
- table cell vertical alignment (`top`/`center`/`bottom`) heuristics inside tall rows
- table cell spacing (`tblCellSpacing`) heuristics affecting effective cell width and row advance
- run-aware paragraph width estimation that uses per-run font size and a character-width heuristic instead of only raw character count

Current local validation status: `117/117` tests passing.

## Feature Coverage

| Area | Status | Notes |
|------|--------|-------|
| Paragraphs and runs | ✅ Implemented | Text, basic run formatting, line/paragraph spacing, indent, and run-aware width heuristics are serialized/applied. |
| Styles and fonts | ✅ Implemented | Style sheet and font table emission are present. |
| Sections | ✅ Implemented | Page size, margins, and page-number start metadata are supported. |
| Numbering and lists | ✅ Implemented | Abstract numbering and LFO/LST structures are emitted. |
| Fields and hyperlinks | ✅ Implemented | Field boundaries, nested fields, and hyperlink instructions are emitted. |
| Images | ✅ Implemented with heuristics | Inline and floating images are written via DOC picture blocks and OfficeArt records. |
| Tables | ✅ Implemented with heuristics | Table width/layout logic uses `tcW`, `tblGrid`, `gridSpan`, cell padding, border thickness, row height hints, cell spacing, and cell vertical alignment on both horizontal and vertical geometry paths, but is not a full Word layout engine. |
| Comments and bookmarks | ✅ Implemented | Parsed and emitted in DOC structures. |
| Advanced Word features | ⚠️ Partial / unsupported | SmartArt, equations, VBA/macros, tracked changes fidelity, and full layout parity are not complete. |

## Known Limits

This converter now covers a broad set of common Word constructs, but it still relies on layout heuristics in several places. The most important current limits are:

- paragraph height estimation is width-aware and run-aware, but still heuristic rather than font-metric exact
- floating image placement is substantially improved, but not a full Word-compatible layout engine
- table layout uses inferred widths, padding, border thickness, row height hints, cell spacing, row-height heuristics, and coarse cell vertical alignment behavior, but does not yet model every table rule Word applies
- advanced Office features such as SmartArt, equations, macros, and exact compatibility behavior are not implemented

## Next Phase

The next fidelity phase is focused on deeper table and layout behavior:

- more exact table layout beyond width inference alone
- richer table border and border-conflict behavior beyond simple thickness-driven geometry
- deeper row rules such as richer interaction between explicit heights and complex cell content
- improved line measurement and paragraph height estimation beyond the current per-run heuristic
- additional parity work for complex floating objects and edge-case Word documents

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
