using System.Text;
using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using Aspose.Slides.Foss.Internal;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Verifies that line formatting properties persist through XML round-trip cycles.
/// </summary>
public sealed class LineFormatIntegrationTests
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

    private static (SlidePart slidePart, ShapeCollection shapes) CreateSlideWithShapes()
    {
        var slidePart = new SlidePart();
        slidePart.InitInternal("ppt/slides/slide1.xml");
        slidePart.Element = XDocument.Parse(SlideXml).Root;

        var shapes = new ShapeCollection();
        shapes.InitInternal(slidePart, null);
        return (slidePart, shapes);
    }

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

    /// <summary>
    /// </summary>
    [Fact]
    public void LineColorAndWidth_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);

        shape.LineFormat.Width = 5;
        shape.LineFormat.FillFormat.FillType = FillType.Solid;
        shape.LineFormat.FillFormat.SolidFillColor.Color = Color.DarkRed;

        var reloadedShapes = RoundTripShapes(slidePart);
        var reloadedShape = reloadedShapes[0];

        reloadedShape.LineFormat.Width.Should().BeApproximately(5f, 0.1f);
        reloadedShape.LineFormat.FillFormat.FillType.Should().Be(FillType.Solid);
        reloadedShape.LineFormat.FillFormat.SolidFillColor.Color.Should().Be(Color.DarkRed);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void LineDashStyle_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);

        shape.LineFormat.DashStyle = LineDashStyle.Dash;

        var reloadedShapes = RoundTripShapes(slidePart);
        var reloadedShape = reloadedShapes[0];

        reloadedShape.LineFormat.DashStyle.Should().Be(LineDashStyle.Dash);
    }

    /// <summary>
    /// </summary>
    [Theory]
    [InlineData(LineDashStyle.Solid)]
    [InlineData(LineDashStyle.Dash)]
    [InlineData(LineDashStyle.Dot)]
    [InlineData(LineDashStyle.DashDot)]
    public void MultipleDashStyles_PersistAfterRoundTrip(LineDashStyle dashStyle)
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);

        shape.LineFormat.DashStyle = dashStyle;

        var reloadedShapes = RoundTripShapes(slidePart);
        var reloadedShape = reloadedShapes[0];

        reloadedShape.LineFormat.DashStyle.Should().Be(dashStyle);
    }

    /// <summary>
    /// CustomDashPattern persists after XML round-trip.
    /// </summary>
    [Fact]
    public void CustomDashPattern_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);

        shape.LineFormat.CustomDashPattern = [3.0f, 1.0f, 1.0f, 1.0f];

        var reloadedShapes = RoundTripShapes(slidePart);
        var reloadedShape = reloadedShapes[0];

        reloadedShape.LineFormat.CustomDashPattern.Should().HaveCount(4);
        reloadedShape.LineFormat.CustomDashPattern[0].Should().Be(3.0f);
        reloadedShape.LineFormat.DashStyle.Should().Be(LineDashStyle.Custom);
    }

    /// <summary>
    /// Style persists after XML round-trip.
    /// </summary>
    [Fact]
    public void Style_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);

        shape.LineFormat.Style = LineStyle.ThinThick;

        var reloadedShapes = RoundTripShapes(slidePart);
        var reloadedShape = reloadedShapes[0];

        reloadedShape.LineFormat.Style.Should().Be(LineStyle.ThinThick);
    }

    /// <summary>
    /// MiterLimit persists after XML round-trip.
    /// </summary>
    [Fact]
    public void MiterLimit_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);

        shape.LineFormat.MiterLimit = 8.0f;

        var reloadedShapes = RoundTripShapes(slidePart);
        var reloadedShape = reloadedShapes[0];

        reloadedShape.LineFormat.MiterLimit.Should().Be(8.0f);
        reloadedShape.LineFormat.JoinStyle.Should().Be(LineJoinStyle.Miter);
    }
}
