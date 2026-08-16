# Aspose.Slides FOSS for .NET

[![CI](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-.NET/actions/workflows/ci.yml/badge.svg)](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-.NET/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-.NET/blob/main/LICENSE)
[![.NET 9.0](https://img.shields.io/badge/.NET-9.0-512BD4.svg)](https://dotnet.microsoft.com/)

An MIT-licensed .NET library that creates, reads and edits PowerPoint `.pptx` presentations by
building the Office Open XML package itself — no PowerPoint installation, no COM interop, no native
dependency and no NuGet dependencies at all.

It is for developers who need to generate or modify decks on a server or in a build, and who are
working *inside* the PowerPoint file format. It is not a renderer or a converter: it does not
produce PDF, HTML or images, and the things it cannot do are listed in full under
[What it cannot do](#what-it-cannot-do).

---

## Requirements

| | |
|---|---|
| Target framework | `net9.0` |
| Package dependencies | none — the library declares no `PackageReference` |
| Native dependencies | none — no `DllImport`, no `LibraryImport`, no `unsafe` code |
| Platforms | the CI workflow builds and runs every test on `ubuntu-latest`, `windows-latest` and `macos-latest` |

A `net9.0` assembly also loads on later .NET runtimes by roll-forward, but CI does not exercise
that, so this file does not promise it.

## Installation

**There is no NuGet package yet.** `Aspose.Slides.Foss` is the package id the project builds under,
but nothing has been published to nuget.org, so `dotnet add package Aspose.Slides.Foss` will not
resolve. The two routes that work today are source and CI artefacts.

### Build from source

```bash
git clone https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-.NET.git
cd Aspose.Slides-FOSS-for-.NET
dotnet build src/Aspose.Slides.Foss/Aspose.Slides.Foss.csproj --configuration Release
```

Then reference the project from your own:

```xml
<ItemGroup>
  <ProjectReference Include="../Aspose.Slides-FOSS-for-.NET/src/Aspose.Slides.Foss/Aspose.Slides.Foss.csproj" />
</ItemGroup>
```

### Install the package CI built

Every push, on any branch, whose tests pass produces a `.nupkg` and a symbol `.snupkg` as workflow
artefacts. Download them from the
[Actions](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-.NET/actions) tab, put them in
a folder, and point NuGet at that folder:

```xml
<!-- nuget.config next to your solution -->
<configuration>
  <packageSources>
    <add key="local" value="./local-packages" />
  </packageSources>
</configuration>
```

```bash
dotnet add package Aspose.Slides.Foss --version 26.8.0
```

`26.8.0` is the version in `Directory.Build.props` today; use whatever the artefact you downloaded
is called.

---

## Quick start

```csharp
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Export;

using (var presentation = new Presentation())
{
    ISlide slide = presentation.Slides[0];
    IAutoShape shape = slide.Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 400, 100);
    shape.AddTextFrame("Hello from Aspose.Slides FOSS");

    presentation.Save("hello.pptx", SaveFormat.Pptx);
}

using var reopened = new Presentation("hello.pptx");
var reread = (IAutoShape)reopened.Slides[0].Shapes![0];

Console.WriteLine($"slides: {reopened.Slides.Count}");
Console.WriteLine($"shapes: {reopened.Slides[0].Shapes!.Count}");
Console.WriteLine($"text:   {reread.TextFrame?.Text}");
```

Output:

```
slides: 1
shapes: 1
text:   Hello from Aspose.Slides FOSS
```

A new `Presentation` is a 13-part package with one slide, one master, one layout and a slide size of
`9144000 × 6858000` EMU — 4:3, `type="screen4x3"`. There is no API to change the slide size.

### Why the `!`

`IBaseSlide.Shapes` is declared `IShapeCollection?` and returns `null` for a slide that is not backed
by a package part; a few other members are annotated the same way. The samples here use the
null-forgiving operator so that they compile without warnings in a project with
`<Nullable>enable</Nullable>`, which is the default for new .NET projects. Every sample on this page
was compiled with nullable enabled and produced zero warnings.

---

## Examples

Every example below was compiled and executed against this revision, and the values quoted after each
one were read out of the `.pptx` it wrote by an independent ZIP/XML reader — not by asking the
library to read its own file back.

Three namespaces cover every sample on this page, and each sample below assumes all three:

```csharp
using Aspose.Slides.Foss;           // Presentation, the I* interfaces, ShapeType, FontData
using Aspose.Slides.Foss.Drawing;   // Color, PointF
using Aspose.Slides.Foss.Export;    // SaveFormat
```

`Color` and `PointF` are this library's own types in `Aspose.Slides.Foss.Drawing`; they are not the
`System.Drawing` types of the same name, and neither namespace is in scope by default.

### Text and formatting

```csharp
using var presentation = new Presentation();
IAutoShape shape = presentation.Slides[0].Shapes!.AddAutoShape(
    ShapeType.Rectangle, 50, 50, 400, 150);
ITextFrame textFrame = shape.AddTextFrame("Formatted text")!;

IParagraph paragraph = textFrame.Paragraphs[0];
paragraph.ParagraphFormat.Alignment = TextAlignment.Center;

IBasePortionFormat format = paragraph.Portions[0].PortionFormat!;
format.FontHeight = 24;
format.FontBold = NullableBool.True;
format.LatinFont = new FontData("Verdana");
format.FillFormat!.FillType = FillType.Solid;
format.FillFormat.SolidFillColor.Color = Color.FromArgb(255, 0, 70, 127);

presentation.Save("text.pptx", SaveFormat.Pptx);
```

In `ppt/slides/slide1.xml`: `<a:rPr b="1" sz="2400">`, `<a:latin typeface="Verdana"/>`,
`<a:pPr algn="ctr">`.

### Tables

```csharp
using var presentation = new Presentation();
ITable table = presentation.Slides[0].Shapes!.AddTable(
    50, 50, [150.0, 150.0], [40.0, 40.0]);

table.Rows[0][0].TextFrame!.Text = "Region";
table.Rows[0][1].TextFrame!.Text = "Revenue";
table.Rows[1][0].TextFrame!.Text = "EMEA";
table.Rows[1][1].TextFrame!.Text = "1,240";

table.MergeCells(table.Rows[0][0], table.Rows[0][1], allowSplitting: false);

presentation.Save("table.pptx", SaveFormat.Pptx);
```

Four `<a:tc>` cells, two of them carrying a merge attribute.

### Pictures

```csharp
using var presentation = new Presentation();

// from a byte array
IPPImage image = presentation.Images.AddImage(File.ReadAllBytes("photo.png"));
presentation.Slides[0].Shapes!.AddPictureFrame(
    ShapeType.Rectangle, 50, 50, 200, 200, image);

// or from a stream
using FileStream stream = File.OpenRead("photo.png");
IPPImage streamed = presentation.Images.AddImage(stream);

presentation.Save("picture.pptx", SaveFormat.Pptx);
```

One `<a:blip r:embed="…">` in the slide and one part under `ppt/media/`. There is **no** overload
that takes a file path — read the bytes yourself, as above.

### Effects

```csharp
using var presentation = new Presentation();
IAutoShape shape = presentation.Slides[0].Shapes!.AddAutoShape(
    ShapeType.Rectangle, 50, 50, 300, 120);

shape.FillFormat.FillType = FillType.Solid;
shape.FillFormat.SolidFillColor.Color = Color.FromArgb(255, 30, 120, 200);

IEffectFormat effects = shape.EffectFormat;
effects.EnableOuterShadowEffect();
effects.OuterShadowEffect!.BlurRadius = 10;
effects.OuterShadowEffect.Distance = 5;

presentation.Save("effects.pptx", SaveFormat.Pptx);
```

All eight effects were enabled one at a time, each into its own file, and each produced its element
inside `<a:effectLst>`: `a:outerShdw`, `a:innerShdw`, `a:glow`, `a:softEdge`, `a:reflection`,
`a:blur`, `a:prstShdw`, `a:fillOverlay`.

### Notes and comments

```csharp
using var presentation = new Presentation();
ISlide slide = presentation.Slides[0];

INotesSlide notes = slide.NotesSlideManager.AddNotesSlide();
notes.NotesTextFrame.Text = "Open with the revenue figures.";

ICommentAuthor author = presentation.CommentAuthors.AddAuthor("Jane Smith", "JS");
author.Comments.AddComment("Check these numbers.", slide, new PointF(2.0f, 2.0f), DateTime.UtcNow);

presentation.Save("notes-comments.pptx", SaveFormat.Pptx);
```

Writes `ppt/notesSlides/notesSlide1.xml` with its `ppt/notesMasters/notesMaster1.xml`, and the
classic comment pair `ppt/commentAuthors.xml` + `ppt/comments/comment1.xml`. It does **not** write
the modern `ppt/threadedComments/` part — see [What it cannot do](#what-it-cannot-do).

### Sections

```csharp
using var presentation = new Presentation();
presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);

presentation.Sections.AddSection("Introduction", presentation.Slides[0]);
presentation.Sections.AddSection("Detail", presentation.Slides[1]);

presentation.Save("sections.pptx", SaveFormat.Pptx);
```

Two `<p14:section>` elements in `ppt/presentation.xml`. A `Section` has no public constructor: it is
created through `Sections.AddSection` or `Sections.AppendEmptySection`, which bind it to the
presentation whose slides it divides.

### Streams

```csharp
using var presentation = new Presentation();
presentation.Slides[0].Shapes!.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 50);

using var buffer = new MemoryStream();
presentation.Save(buffer, SaveFormat.Pptx);

buffer.Position = 0;
using var reopened = new Presentation(buffer);
```

`Presentation` opens from a path or a `Stream`, with or without `ILoadOptions`, and `Save` has the
matching overloads including a slide-subset one.

---

## What it can do

Everything in this list was exercised through the public API and then confirmed by reading the XML
inside the file that came out.

- **Presentations** — create, open a path or a stream, save to a path or a stream, save a subset of
  slides, `IDisposable` throughout.
- **Slides** — add empty, insert, remove (the slide part is removed from the package, not just from
  the list), clone, hide (`<p:sld show="0">`), enumerate; masters and layouts are enumerable and a
  master can be cloned.
- **Sections** — add, append, remove, remove with slides, reorder with slides.
- **Shapes** — AutoShapes for 187 of the 189 `ShapeType` values: each of the 187 was written into its
  own file and each produced its own distinct `<a:prstGeom prst="…">`. The two exceptions are
  `ShapeType.NotDefined` and `ShapeType.Custom`, which have no preset to write; `AddAutoShape`
  accepts them without complaint and the shape comes out as `prst="rect"`. Also picture frames,
  tables, connectors bound to shapes by connection site, `Reorder` for z-order, and adjust values on
  geometry shapes.
- **Text** — text frames, paragraphs, portions; character formatting (bold, italic, underline,
  strikethrough, size, spacing, caps, latin/east-asian/complex-script/symbol fonts), paragraph
  formatting (alignment, indent, margins, spacing, symbol and numbered bullets), text-frame
  formatting (margins, wrap, anchor, autofit, columns, rotation, vertical text).
- **Fill** — solid, gradient (stops, direction, shape, angle, tile flip), pattern, picture (with
  crop, stretch, tile and fill mode), no-fill; the same fill model applies to shapes, lines, text
  portions and table cells.
- **Lines** — width, dash style and custom dash pattern, cap, join, miter limit, compound style,
  alignment, and arrowheads at both ends with style, width and length.
- **Effects** — outer shadow, inner shadow, glow, soft edge, reflection, blur, preset shadow, fill
  overlay.
- **3-D** — bevel top and bottom, extrusion height and colour, contour width and colour, material
  preset, a camera with a preset, field of view, zoom and rotation, and a light rig with a preset,
  direction and rotation.
- **Tables** — rows, columns, cells, cell merge (splitting optional), cell fill, the six cell
  borders, cell margins and anchoring, and a table style preset written as `<a:tableStyleId>`.
- **Pictures** — embed from a byte array or a stream, deduplicated by content.
- **Notes** — a notes slide per slide, written with the notes master it requires, plus footer text,
  footer visibility and slide-number visibility on the notes slide.
- **Comments** — authors, comments with position and timestamp, on the classic comment list.
- **Document properties** — core, extended and custom; `docProps/core.xml`, `docProps/app.xml` and
  `docProps/custom.xml` are all written.
- **Unknown parts** — parts the library does not model are carried through a load and a save
  unchanged.

---

## What it cannot do

This section is the point of this file. Nothing below is a "coming soon"; it is what the API does
not contain today.

### Not in the public API

| | |
|---|---|
| Charts | `IShapeCollection.AddChart` does not exist |
| SmartArt, OLE objects, video, audio | no `AddSmartArt` / `AddOleObjectFrame` / `AddVideoFrame` / `AddAudioFrame` |
| Group shapes | no `AddGroupShape`. `IGroupShape` is declared but nothing implements it, and the public `GroupShape` class adds nothing to `Shape` — so no shape can hold child shapes |
| Animations and slide transitions | no `ISlide.Timeline`, no `ISlide.SlideShowTransition` |
| Hyperlinks | no `HyperlinkClick` on a shape or on a text portion |
| Slide backgrounds | no `ISlide.Background` |
| Themes | no `IPresentation.MasterTheme` |
| Slide size | no `IPresentation.SlideSize` — a new deck is 4:3 and cannot be changed |
| Threaded comments | replies are held in memory; no `ppt/threadedComments/` part is written |
| Encryption and protection | no `IPresentation.Protect` |
| Rendering and conversion | no PDF, HTML, SVG, image or text export of any kind |
| VBA macros, digital signatures | not modelled |
| Reordering slides | `IShapeCollection.Reorder` exists for shapes; `ISlideCollection` has no equivalent |

### Save formats

`SaveFormat` declares **21** values. `Save` writes **three** of them and raises
`NotSupportedException` for the other **eighteen** — it never writes a package under a name that
claims to be a format it did not produce.

| Written | Refused |
|---|---|
| `Pptx`, `Ppsx`, `Potx` | `Ppt`, `Pdf`, `Xps`, `Tiff`, `Odp`, `Pptm`, `Ppsm`, `Potm`, `Html`, `Swf`, `Otp`, `Pps`, `Pot`, `Fodp`, `Gif`, `Html5`, `Md`, `Xml` |

The three that are written are three genuinely different packages, each with its own main-part
content type:

```
Pptx  -> application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml
Ppsx  -> application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml
Potx  -> application/vnd.openxmlformats-officedocument.presentationml.template.main+xml
```

The refusal says what to do instead:

```
Save format 'Pdf' is not supported: it is not an Office Open XML presentation package and this
library does not render or convert. Supported formats: Pptx, Ppsx, Potx.

Save format 'Pptm' is not supported: a macro-enabled package requires a VBA project part, which
this library does not write. Supported formats: Pptx, Ppsx, Potx.
```

Give the file the extension that matches the format you asked for — `.pptx`, `.ppsx`, `.potx`.
Office keys off the name as well as the content type, and a file whose two disagree may be refused.

### What "round-trip" does and does not mean

Opening a deck and saving it again preserves every part, and preserves the great majority of them
byte for byte. It does not rewrite nothing at all. Measured on this repository's own 46-part
PowerPoint-authored fixture (`tests/test_data/PowerPointDeck.pptx`, `<Application>Microsoft Office
PowerPoint</Application>`):

| | |
|---|---|
| Parts in / out | 46 / 46 |
| Parts dropped | 0 |
| Parts added | 0 |
| Parts byte-identical afterwards | 42 of 46 |
| Parts rewritten | `[Content_Types].xml`, `docProps/app.xml`, `docProps/core.xml`, `ppt/presentation.xml` |
| Slide text runs | identical, in order |

`docProps/core.xml` is rewritten because saving stamps a modification time; `docProps/app.xml` is
regenerated from the presentation; the other two are rebuilt from the package model. Parts this
library has no model for — `presProps`, `viewProps`, `tableStyles`, thumbnails, extra themes, unused
layouts — are among the 42 that come back unchanged. That is a specific, measured claim about one
real deck, and it is deliberately narrower than "full fidelity".

### Text language

A deck written by this library contains no `lang=` attribute anywhere and no `<a:rPr>` on a portion
that was never formatted. PowerPoint will apply the authoring machine's default language rather than
a language the file states.

---

## Choosing an edition

There are four editions of this library and they are **not** interchangeable. Three of them —
.NET, Java and C++ — are the same design in three languages; Python is a larger and different one.
The table below was measured on 2026-08-16 by running each edition and reading the XML in the file
it produced.

| | .NET | Java | C++ | Python |
|---|---|---|---|---|
| Create, open, round-trip, save | yes | yes | yes | yes |
| Text, tables, connectors, fills, all 8 effects, 3-D | yes | yes | yes | yes |
| Notes and classic comments | yes | yes | yes | yes |
| Threaded comment part | no | yes | no | yes |
| Save to a stream | yes | yes | **no** | yes |
| Add a picture from a stream | yes | yes | **no** | yes |
| Sections | **yes** | no | no | no |
| Charts | no | no | no | **yes** |
| Animations | no | no | no | **yes** |
| Slide transitions | no | no | no | **yes** |
| Themes | no | no | no | **yes** |
| Slide backgrounds | no | no | no | **yes** |
| Group shapes | no | no | no | **yes** |
| Hyperlinks | no | no | no | **yes** |
| Markdown export | no | no | no | **yes** |
| Formats `Save` writes | 3 | 3 | 6 | 7 |
| Default slide size of a new deck | 4:3 | 4:3 | 4:3 | **16:9** |
| Layouts in a new deck | 1 | 1 | 1 | **11** |

The short version: **if you need charts, animations, transitions, themes, backgrounds, group shapes
or hyperlinks, none of them exist in this edition — use the Python one.** If you need sections, this
is the only edition that has them. Code written against the Python examples does not port to the
other three, and a new deck is not even the same shape.

- [Aspose.Slides FOSS for Java](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-Java)
- [Aspose.Slides FOSS for C++](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-Cpp)
- [Aspose.Slides FOSS for Python](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-Python)

---

## Documentation and support

This library mirrors the naming of the commercial **Aspose.Slides for .NET** product, so that
product's documentation is usually the fastest way to understand what a shared type name means.
It describes a much larger API: check the tables above before relying on anything you read there.

- [Aspose.Slides for .NET — product page](https://products.aspose.com/slides/net/)
- [Documentation](https://docs.aspose.com/slides/net/)
- [API reference](https://reference.aspose.com/slides/net/)
- [Knowledge base](https://kb.aspose.com/slides/net/)
- [Free support forum](https://forum.aspose.com/c/slides/11)

**When you want the commercial product instead of this one:** if you need to render or convert —
PDF, images, HTML, thumbnails — or need charts, SmartArt, animations, OLE objects or macro-enabled
files, or need a supported product with a licence, none of that is here and none of it is planned in
this repository. This library is the right choice when you are reading and writing `.pptx` and want
an MIT-licensed dependency with no other dependencies at all.

Bugs and feature requests for **this** library belong in
[this repository's issue tracker](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-.NET/issues),
not in the commercial product's forum.

---

## Contributing

Pull requests are welcome. Read
[CONTRIBUTING.md](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-.NET/blob/main/CONTRIBUTING.md)
first — it explains the build, the three test suites, and the one rule that is specific to this
project: **a writer fix ships with a test that asserts on the produced `.pptx` package, not on what
the library reads back.**

By participating you agree to the
[Code of Conduct](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-.NET/blob/main/CODE_OF_CONDUCT.md).
Changes are recorded in
[CHANGELOG.md](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-.NET/blob/main/CHANGELOG.md).

## Security

Do not report a vulnerability in a public issue. See
[SECURITY.md](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-.NET/blob/main/SECURITY.md)
for the reporting route.

## License

MIT — see [LICENSE](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-.NET/blob/main/LICENSE).
Copyright (c) 2026 Aspose Pty Ltd.
