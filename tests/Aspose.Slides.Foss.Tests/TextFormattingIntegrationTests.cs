using System.Text;
using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using Aspose.Slides.Foss.Internal;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Verifies that text formatting properties persist through XML round-trip cycles.
/// Uses a slide XML tree to exercise the full formatting pipeline (AddAutoShape,
/// TextFrame, Paragraphs, Portions, PortionFormat) with a serialize/deserialize cycle.
/// </summary>
public sealed class TextFormattingIntegrationTests
{
    private static readonly string SlideXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
               xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:cSld>
            <p:spTree>
              <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
              <p:grpSpPr/>
            </p:spTree>
          </p:cSld>
        </p:sld>
        """;

    /// <summary>
    /// Creates a SlidePart backed by a parsed slide XML tree and a ShapeCollection wired to it.
    /// </summary>
    private static (SlidePart slidePart, ShapeCollection shapes) CreateSlideWithShapes()
    {
        var slidePart = new SlidePart();
        slidePart.InitInternal("ppt/slides/slide1.xml");
        slidePart.Element = XDocument.Parse(SlideXml).Root;

        var shapes = new ShapeCollection();
        shapes.InitInternal(slidePart, null);
        return (slidePart, shapes);
    }

    /// <summary>
    /// Serializes the slide XML to string and re-parses it, returning a fresh ShapeCollection.
    /// This simulates a save/reload cycle at the XML level.
    /// </summary>
    private static ShapeCollection RoundTripShapes(SlidePart slidePart)
    {
        var xml = slidePart.Element!.ToString();
        var reloadedRoot = XDocument.Parse(xml).Root!;

        var newSlidePart = new SlidePart();
        newSlidePart.InitInternal("ppt/slides/slide1.xml");
        newSlidePart.Element = reloadedRoot;

        var shapes = new ShapeCollection();
        shapes.InitInternal(newSlidePart, null);
        return shapes;
    }

    private static IBasePortionFormat GetFirstPortionFormat(IAutoShape shape)
    {
        return shape.TextFrame!.Paragraphs[0].Portions[0].PortionFormat!;
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void BoldItalic_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 300, 100);
        shape.AddTextFrame("Bold Italic");
        var fmt = GetFirstPortionFormat(shape);

        fmt.FontBold = NullableBool.True;
        fmt.FontItalic = NullableBool.True;

        var reloadedShapes = RoundTripShapes(slidePart);
        var reloadedShape = (IAutoShape)reloadedShapes[0];
        var reloadedFmt = GetFirstPortionFormat(reloadedShape);

        reloadedFmt.FontBold.Should().Be(NullableBool.True);
        reloadedFmt.FontItalic.Should().Be(NullableBool.True);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void Underline_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 300, 100);
        shape.AddTextFrame("Underlined");
        var fmt = GetFirstPortionFormat(shape);

        fmt.FontUnderline = TextUnderlineType.Single;

        var reloadedShapes = RoundTripShapes(slidePart);
        var reloadedShape = (IAutoShape)reloadedShapes[0];
        var reloadedFmt = GetFirstPortionFormat(reloadedShape);

        reloadedFmt.FontUnderline.Should().Be(TextUnderlineType.Single);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void Strikethrough_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 300, 100);
        shape.AddTextFrame("Strikethrough");
        var fmt = GetFirstPortionFormat(shape);

        fmt.StrikethroughType = TextStrikethroughType.Single;

        var reloadedShapes = RoundTripShapes(slidePart);
        var reloadedShape = (IAutoShape)reloadedShapes[0];
        var reloadedFmt = GetFirstPortionFormat(reloadedShape);

        reloadedFmt.StrikethroughType.Should().Be(TextStrikethroughType.Single);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void FontSize_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 300, 100);
        shape.AddTextFrame("Big Text");
        var fmt = GetFirstPortionFormat(shape);

        fmt.FontHeight = 28;

        var reloadedShapes = RoundTripShapes(slidePart);
        var reloadedShape = (IAutoShape)reloadedShapes[0];
        var reloadedFmt = GetFirstPortionFormat(reloadedShape);

        reloadedFmt.FontHeight.Should().Be(28);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void FontColor_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 300, 100);
        shape.AddTextFrame("Red Text");
        var fmt = GetFirstPortionFormat(shape);

        fmt.FillFormat!.FillType = FillType.Solid;
        fmt.FillFormat.SolidFillColor.Color = Color.Red;

        var reloadedShapes = RoundTripShapes(slidePart);
        var reloadedShape = (IAutoShape)reloadedShapes[0];
        var reloadedFmt = GetFirstPortionFormat(reloadedShape);

        reloadedFmt.FillFormat!.FillType.Should().Be(FillType.Solid);
        reloadedFmt.FillFormat.SolidFillColor.Color.Should().Be(Color.Red);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void LatinFont_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 300, 100);
        shape.AddTextFrame("Courier");
        var fmt = GetFirstPortionFormat(shape);

        fmt.LatinFont = new FontData("Courier New");

        var reloadedShapes = RoundTripShapes(slidePart);
        var reloadedShape = (IAutoShape)reloadedShapes[0];
        var reloadedFmt = GetFirstPortionFormat(reloadedShape);

        reloadedFmt.LatinFont.Should().NotBeNull();
        reloadedFmt.LatinFont!.FontName.Should().Be("Courier New");
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void ParagraphAlignment_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 400, 200);
        shape.AddTextFrame("Centered");
        shape.TextFrame!.Paragraphs[0].ParagraphFormat.Alignment = TextAlignment.Center;

        var reloadedShapes = RoundTripShapes(slidePart);
        var reloadedShape = (IAutoShape)reloadedShapes[0];
        var pf = reloadedShape.TextFrame!.Paragraphs[0].ParagraphFormat;

        pf.Alignment.Should().Be(TextAlignment.Center);
    }
}
