# Aspose.Slides FOSS for .NET

[![NuGet](https://img.shields.io/nuget/v/Aspose.Slides.FOSS.svg)](https://www.nuget.org/packages/Aspose.Slides.FOSS/)
[![Downloads](https://img.shields.io/nuget/dt/Aspose.Slides.FOSS.svg)](https://www.nuget.org/packages/Aspose.Slides.FOSS/)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](https://opensource.org/licenses/MIT)

A free, MIT-licensed library for creating, reading and editing PowerPoint `.pptx` presentations from
.NET. It builds the Office Open XML package itself, so it has **no package dependencies** and needs
neither PowerPoint nor COM interop, and it runs on Windows, Linux and macOS.

It is a focused library rather than a small one. Everything below is verified by opening the file it
produced and asserting on the XML inside — never by asking the library to read its own output back.

[![Aspose.Slides FOSS for .NET](https://raw.githubusercontent.com/aspose-slides-foss/Aspose.Slides-FOSS-for-.NET/main/docs/media/banner-readme.png)](https://products.aspose.org/slides/net/)

[Product page](https://products.aspose.org/slides/net/) | [Documentation](https://docs.aspose.org/) | [API reference](https://reference.aspose.org/) | [Source](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-.NET) | [Issues](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-.NET/issues) | [Free support](https://forum.aspose.com/c/slides/11)

## Install

```bash
dotnet add package Aspose.Slides.FOSS
```

Targets `net8.0` and `net10.0`.

## Quick start

```csharp
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Export;

using var presentation = new Presentation();

ISlide slide = presentation.Slides[0];
IAutoShape shape = slide.Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 400, 100);
ITextFrame text = shape.AddTextFrame("Hello from Aspose.Slides FOSS")!;
text.Paragraphs[0].Portions[0].PortionFormat!.FontBold = NullableBool.True;

presentation.Save("hello.pptx", SaveFormat.Pptx);
```

The namespace is `Aspose.Slides.Foss` — the package identifier is capitalised `FOSS`, the code is
not. The `!` operators are needed rather than decorative: `ISlide.Shapes` and
`IAutoShape.AddTextFrame` are declared nullable, and a project with nullable reference types enabled
warns without them.

## What it does

- **Presentations** — create, open a file or a stream, save to a file or a stream.
- **Slides** — add, remove, clone, hide. **Sections**, including reordering slides with them.
- **Masters and layouts** — enumerate them, and clone a master.
- **Shapes** — AutoShapes (187 of the 189 `ShapeType` values), connectors with connection sites,
  tables with cell merging, picture frames.
- **Text** — text frames, paragraphs and portions; character and paragraph formatting; alignment.
- **Appearance** — solid, gradient, pattern and picture fills; line formatting; all eight visual
  effects; 3-D formatting.
- **Pictures** — added from a `byte[]` or a `Stream`.
- **Notes and comments** — speaker notes, and classic comments with their authors.
- **Round-trip** — XML parts the library does not model are preserved verbatim. On the 46-part
  PowerPoint deck in the test data, 46 parts go in and 46 come out, none dropped and none added; 42
  are byte-identical and 4 are rewritten (`[Content_Types].xml`, `docProps/app.xml`,
  `docProps/core.xml`, `ppt/presentation.xml`).

## What it does not do

Stated plainly, because a library that fails silently is worse than one that says no. Every entry
below raises rather than producing a file that looks right and is not.

- **No rendering or conversion of any kind.** No PDF, HTML, SVG, image or text export. `SaveFormat`
  declares 21 values; three are written — `Pptx`, `Ppsx`, `Potx` — and the other 18 raise
  `NotSupportedException` naming the format and the three that work.
- **No macro-enabled formats.** `Pptm`, `Ppsm` and `Potm` need a VBA project part this library does
  not write, and they raise for that reason specifically.
- **No charts, no SmartArt, no OLE objects, no video or audio frames, no group shapes.**
- **No animations and no slide transitions.**
- **No hyperlinks.** `IHyperlinkContainer` is declared and several types implement it, but the
  interface has no members — implementing it gives you nothing.
- **No slide backgrounds, no themes, and no slide size control.** A new presentation is 4:3 and
  cannot be changed.
- **No threaded comments.** A reply is accepted by the object model and no `ppt/threadedComments/`
  part is written, so it does not survive a save. Classic comments do.
- **No encryption, protection, digital signatures or mathematical text.**
- **`IImageCollection.AddImage` has no overload taking a path** — read the bytes or open the stream
  yourself.

## When you want the commercial product

This library is not a drop-in replacement for
[Aspose.Slides for .NET](https://products.aspose.com/slides/net/), and it is not trying to be. If
you need PDF or image rendering, charts, animations, the other presentation formats, or a support
commitment, that is the product to use. If you need to build and read `.pptx` packages under a
permissive licence with nothing else in your dependency tree, this one is complete for that.

## Licence and support

MIT. The full text is packed alongside this file.

Questions and bug reports are welcome on the
[issue tracker](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-.NET/issues) or on the
[free support forum](https://forum.aspose.com/c/slides/11). A report that names the version and
attaches the file PowerPoint would not open is one that can be acted on.

[Product page](https://products.aspose.org/slides/net/) | [Documentation](https://docs.aspose.org/) | [API reference](https://reference.aspose.org/) | [Source](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-.NET) | [Issues](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-.NET/issues) | [Free support](https://forum.aspose.com/c/slides/11)
