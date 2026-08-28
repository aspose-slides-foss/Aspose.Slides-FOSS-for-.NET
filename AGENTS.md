# Aspose.Slides FOSS — Guide for AI Agents

You are working with `Aspose.Slides.Foss`, the official open-source .NET library by Aspose.Slides for creating, reading, and editing PowerPoint (.pptx) presentations.

## Getting Started

Requires .NET 9.0 — the framework the library targets and the only one it is tested on. No package
dependencies and no native ones.

```bash
git clone https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-.NET.git
cd Aspose.Slides-FOSS-for-.NET
dotnet build
```

## Core Concepts

- **`Presentation`** is the root object. It owns slides, masters, layouts, images, document properties, and comments.
- Always wrap `Presentation` in a `using` statement to ensure proper cleanup.
- Save with `prs.Save("out.pptx", SaveFormat.Pptx)`. `SaveFormat.Ppsx` and `SaveFormat.Potx` are written too; every other value raises `NotSupportedException`.
- Give the file the extension of the format you ask for — `.pptx`, `.ppsx`, `.potx`. PowerPoint refuses a file whose name and content type disagree.
- Parts this library has no model for are carried through a load and a save unchanged. A save is not
  a byte-for-byte copy of the whole package, though: `[Content_Types].xml`, `docProps/app.xml`,
  `docProps/core.xml` and `ppt/presentation.xml` are rebuilt. See "What 'round-trip' does and does
  not mean" in `README.md` for the measured figures.

## Import Pattern

```csharp
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Export;
using Aspose.Slides.Foss.Drawing;
```

## Quick Reference

### Create a presentation

```csharp
using var prs = new Presentation();
var slide = prs.Slides[0];  // first slide exists by default
prs.Save("new.pptx", SaveFormat.Pptx);
```

### Open an existing file

```csharp
using var prs = new Presentation("input.pptx");
foreach (var slide in prs.Slides)
{
    foreach (var shape in slide.Shapes)
        Console.WriteLine(shape.Name);
}
prs.Save("output.pptx", SaveFormat.Pptx);
```

### Add shapes

```csharp
var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, x, y, width, height);
shape.AddTextFrame("Hello");
```

Coordinates and dimensions are in points (1 point = 1/72 inch).

### Text formatting

```csharp
var portion = shape.TextFrame.Paragraphs[0].Portions[0];
var fmt = portion.PortionFormat;
fmt.FontHeight = 24;
fmt.FontBold = NullableBool.True;
fmt.FontItalic = NullableBool.True;
fmt.FillFormat.FillType = FillType.Solid;
fmt.FillFormat.SolidFillColor.Color = Color.FromArgb(255, 0, 70, 127);
```

### Tables

```csharp
double[] colWidths = [120.0, 120.0, 120.0];
double[] rowHeights = [40.0, 40.0];
var table = slide.Shapes.AddTable(x, y, colWidths, rowHeights);
table.Rows[0][0].TextFrame.Text = "Header";
```

### Connectors

```csharp
var conn = slide.Shapes.AddConnector(ShapeType.BentConnector3, 0, 0, 10, 10);
conn.StartShapeConnectedTo = shapeA;
conn.StartShapeConnectionSiteIndex = 3;  // 0=top, 1=left, 2=bottom, 3=right
conn.EndShapeConnectedTo = shapeB;
conn.EndShapeConnectionSiteIndex = 1;
```

### Fills

```csharp
shape.FillFormat.FillType = FillType.Solid;
shape.FillFormat.SolidFillColor.Color = Color.FromArgb(255, 30, 120, 200);
```

Also supports: `FillType.Gradient`, `FillType.Pattern`, `FillType.Picture`, `FillType.NoFill`.

### Images

```csharp
var imageBytes = File.ReadAllBytes("photo.png");
var image = prs.Images.AddImage(imageBytes);
slide.Shapes.AddPictureFrame(ShapeType.Rectangle, x, y, w, h, image);
```

### Notes

```csharp
var notes = slide.NotesSlideManager.AddNotesSlide();
notes.NotesTextFrame.Text = "Speaker notes here.";
```

### Comments

```csharp
var author = prs.CommentAuthors.AddAuthor("Jane Smith", "JS");
author.Comments.AddComment("Review this", slide, new PointF(2.0f, 2.0f), DateTime.Now);
```

### Document properties

```csharp
prs.DocumentProperties.Title = "Quarterly Report";
prs.DocumentProperties.Author = "Finance Team";
prs.DocumentProperties.SetCustomPropertyValue("Version", 3);
```

### Slide operations

```csharp
prs.Slides.AddEmptySlide(prs.LayoutSlides[0]);  // add slide
prs.Slides.RemoveAt(1);                          // remove by index
var cloned = prs.Slides.AddClone(prs.Slides[0]); // clone slide
slide.Hidden = true;                              // hide slide
```

### Effects and 3D

```csharp
// Outer shadow
var ef = shape.EffectFormat;
ef.EnableOuterShadowEffect();
ef.OuterShadowEffect.BlurRadius = 10;
ef.OuterShadowEffect.Distance = 5;

// 3D bevel
var td = shape.ThreeDFormat;
td.BevelTop.BevelType = BevelPresetType.Circle;
td.BevelTop.Height = 6;
td.BevelTop.Width = 6;
```

### Line formatting

```csharp
var lf = shape.LineFormat;
lf.Width = 2.5;
lf.DashStyle = LineDashStyle.DashDot;
lf.FillFormat.FillType = FillType.Solid;
lf.FillFormat.SolidFillColor.Color = Color.Red;
```

## Package Structure

```
src/
  Aspose.Slides.Foss/           # Main library
    Presentation.cs              # Root object
    Slide.cs                     # Slide, LayoutSlide, MasterSlide
    ShapeCollection.cs           # Shape management
    AutoShape.cs                 # AutoShape with text frames
    Table.cs                     # Tables, rows, columns, cells
    Connector.cs                 # Shape-to-shape connectors
    TextFrame.cs                 # Text content model
    FillFormat.cs                # Fill styling
    LineFormat.cs                # Line styling
    EffectFormat.cs              # Visual effects
    ThreeDFormat.cs              # 3D formatting
    Comment.cs                   # Slide comments
    DocumentProperties.cs
    Drawing/                     # Color, PointF, SizeF, Size
    Export/                      # SaveFormat, SaveOptions
    Effects/                     # Effect-related interfaces
    Theme/                       # Theme-related types
    Internal/                    # Implementation internals (do not reference directly)
```

## Do

- Always wrap `Presentation` in a `using` statement
- Use `SaveFormat.Pptx` when saving, unless you want a slideshow (`Ppsx`) or a template (`Potx`) — those three are what `Save` writes
- Use `Color.FromArgb(a, r, g, b)` or named constants like `Color.Red`, `Color.Blue`
- Access slides via `prs.Slides[index]` — slides are 0-indexed
- Use `NullableBool` enum (`NullableBool.False`, `NullableBool.True`, `NullableBool.NotDefined`) for boolean formatting properties like `FontBold`
- Import drawing types from `Aspose.Slides.Foss.Drawing`

## Don't

- Don't reference `Aspose.Slides.Foss.Internal` — it is a private implementation detail
- Don't attempt PDF, HTML, SVG, or image export — this library does not render or convert, and
  `Save` writes only the three OOXML presentation packages `Pptx`, `Ppsx` and `Potx`
- Don't use charts, SmartArt, animations, or VBA — they are not in the API
- Don't modify the public API class signatures — they are fixed

## Limitations

These are absences in the API today, not a schedule. **The complete list, with the missing API
member named for each, is the "What it cannot do" section of `README.md`** — read it there rather
than from the summary below, which exists only so that this file is not silent about them:

- Charts, SmartArt, OLE objects, video and audio
- Group shapes, animations, slide transitions, hyperlinks
- Slide backgrounds, themes, and any API for the slide size (a new deck is 4:3 and stays 4:3)
- Adding or cloning a layout — a new deck has one layout and no way to add another
- Threaded comments; replies live in memory and the file carries the classic comment list
- Encryption and protection, VBA macros, digital signatures
- Rendering and conversion of any kind — no PDF, HTML, SVG, image or text export

`Save` writes `SaveFormat.Pptx`, `SaveFormat.Ppsx` and `SaveFormat.Potx`; every other
value of the enum raises `NotSupportedException`.

## Links

- [GitHub](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-.NET)
- [Issue Tracker](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-.NET/issues)
- [Aspose.Slides](https://products.aspose.org/slides/)
