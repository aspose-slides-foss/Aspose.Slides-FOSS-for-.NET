# Aspose.Slides FOSS

The official open-source .NET library by Aspose.Slides for creating, reading, and editing PowerPoint (`.pptx`) presentations.

---

## Quick Start

```csharp
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Export;

// Open an existing presentation
using var prs = new Presentation("input.pptx");
Console.WriteLine($"Slides: {prs.Slides.Count}");
prs.Save("output.pptx", SaveFormat.Pptx);

// Create a new presentation
using var newPrs = new Presentation();
var slide = newPrs.Slides[0];
newPrs.Save("new.pptx", SaveFormat.Pptx);
```

---

## Features

- **Presentation I/O** — Open, create, and save `.pptx` files with full round-trip fidelity
- **Slides** — Add, remove, clone, reorder, and iterate slides
- **Shapes** — AutoShapes, PictureFrames, Tables, Connectors
- **Text** — TextFrame, Paragraph, Portion with character, paragraph, and text frame formatting (including bullets)
- **Fill** — Solid, gradient, pattern, and picture fills
- **Lines** — Width, dash style, arrows, join and alignment
- **Effects** — Outer shadow, glow, soft edge, blur, reflection, inner shadow
- **3D** — Bevel, camera, light rig, material, extrusion depth
- **Document properties** — Core, app, and custom properties
- **Notes slides** — Per-slide notes with header/footer management
- **Comments** — Threaded comments with authors, timestamps, and positions
- **Images** — Embed from file, bytes, or stream

---

## Usage Examples

### Shapes

```csharp
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Export;

using var prs = new Presentation();
var slide = prs.Slides[0];
var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 300, 100);
shape.AddTextFrame("Hello, world!");
prs.Save("shapes.pptx", SaveFormat.Pptx);
```

### Text Formatting

```csharp
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using Aspose.Slides.Foss.Export;

using var prs = new Presentation();
var shape = prs.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 400, 150);
var tf = shape.AddTextFrame("Formatted text");
var fmt = tf.Paragraphs[0].Portions[0].PortionFormat;
fmt.FontHeight = 24;
fmt.FontBold = NullableBool.True;
fmt.FillFormat.FillType = FillType.Solid;
fmt.FillFormat.SolidFillColor.Color = Color.FromArgb(255, 0, 70, 127);
prs.Save("text.pptx", SaveFormat.Pptx);
```

### Table

```csharp
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Export;

using var prs = new Presentation();
var table = prs.Slides[0].Shapes.AddTable(50, 50, [120.0, 120.0, 120.0], [40.0, 40.0]);
table.Rows[0][0].TextFrame.Text = "Name";
table.Rows[0][1].TextFrame.Text = "Value";
prs.Save("table.pptx", SaveFormat.Pptx);
```

### Connector

```csharp
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Export;

using var prs = new Presentation();
var slide = prs.Slides[0];
var box1 = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 100, 150, 60);
var box2 = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 350, 100, 150, 60);
var conn = slide.Shapes.AddConnector(ShapeType.BentConnector3, 0, 0, 10, 10);
conn.StartShapeConnectedTo = box1;
conn.StartShapeConnectionSiteIndex = 3;  // right
conn.EndShapeConnectedTo = box2;
conn.EndShapeConnectionSiteIndex = 1;    // left
prs.Save("connector.pptx", SaveFormat.Pptx);
```

### Fill

```csharp
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using Aspose.Slides.Foss.Export;

using var prs = new Presentation();
var shape = prs.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 300, 150);
shape.FillFormat.FillType = FillType.Solid;
shape.FillFormat.SolidFillColor.Color = Color.FromArgb(255, 30, 120, 200);
prs.Save("fill.pptx", SaveFormat.Pptx);
```

### Notes

```csharp
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Export;

using var prs = new Presentation();
var notes = prs.Slides[0].NotesSlideManager.AddNotesSlide();
notes.NotesTextFrame.Text = "Speaker notes go here.";
prs.Save("notes.pptx", SaveFormat.Pptx);
```

### Comments

```csharp
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using Aspose.Slides.Foss.Export;

using var prs = new Presentation();
var author = prs.CommentAuthors.AddAuthor("Jane Smith", "JS");
var slide = prs.Slides[0];
author.Comments.AddComment("Review this slide", slide, new PointF(2.0f, 2.0f), DateTime.Now);
prs.Save("comments.pptx", SaveFormat.Pptx);
```

### Document Properties

```csharp
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Export;

using var prs = new Presentation();
prs.DocumentProperties.Title = "Q1 Results";
prs.DocumentProperties.Author = "Finance Team";
prs.DocumentProperties.SetCustomPropertyValue("Version", 3);
prs.Save("deck.pptx", SaveFormat.Pptx);
```

---

## Limitations

The following areas are not yet implemented:

- Charts, SmartArt, OLE objects, mathematical text
- Animations and slide transitions
- Export to non-PPTX formats (PDF, HTML, SVG, images)
- VBA macros, digital signatures
- Hyperlinks and action settings

Unknown XML parts encountered during load are preserved verbatim on save —
opening and re-saving a file will never strip content this library does not yet understand.

---

## Links

- [GitHub Repository](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-.NET)
- [Issue Tracker](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-.NET/issues)

---

## License

[MIT License](https://github.com/aspose-slides-foss/Aspose.Slides-FOSS-for-.NET/blob/main/LICENSE)
