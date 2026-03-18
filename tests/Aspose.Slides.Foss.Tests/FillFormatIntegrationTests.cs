using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using Aspose.Slides.Foss.Internal;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Integration tests for FillFormat round-trip through XML serialization.
/// </summary>
public sealed class FillFormatIntegrationTests
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
    /// Solid fill colour persists after XML round-trip.
    /// </summary>
    [Fact]
    public void SolidFill_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);

        shape.FillFormat.FillType = FillType.Solid;
        shape.FillFormat.SolidFillColor.Color = Color.FromArgb(255, 0, 128, 255);

        var reloaded = RoundTripShapes(slidePart);
        var reloadedFill = reloaded[0].FillFormat;

        reloadedFill.FillType.Should().Be(FillType.Solid);
        var color = reloadedFill.SolidFillColor.Color;
        color.Should().NotBeNull();
        color!.R.Should().Be(0);
        color.G.Should().Be(128);
        color.B.Should().Be(255);
    }

    /// <summary>
    /// Gradient fill type and angle persist after XML round-trip.
    /// </summary>
    [Fact]
    public void GradientFill_TypeAndAnglePersistAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);

        shape.FillFormat.FillType = FillType.Gradient;
        var gradFormat = shape.FillFormat.GradientFormat;
        gradFormat.LinearGradientAngle = 45f;

        var reloaded = RoundTripShapes(slidePart);
        var reloadedFill = reloaded[0].FillFormat;

        reloadedFill.FillType.Should().Be(FillType.Gradient);
        var reloadedGrad = reloadedFill.GradientFormat;
        reloadedGrad.GradientShape.Should().Be(GradientShape.Linear);
        reloadedGrad.LinearGradientAngle.Should().BeApproximately(45f, 0.1f);
    }

    /// <summary>
    /// Gradient fill type alone persists after XML round-trip (no stops).
    /// </summary>
    [Fact]
    public void GradientFill_FillTypePersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);

        shape.FillFormat.FillType = FillType.Gradient;

        var reloaded = RoundTripShapes(slidePart);
        var reloadedFill = reloaded[0].FillFormat;

        reloadedFill.FillType.Should().Be(FillType.Gradient);
    }

    /// <summary>
    /// NoFill type persists after XML round-trip.
    /// </summary>
    [Fact]
    public void NoFill_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);

        shape.FillFormat.FillType = FillType.NoFill;

        var reloaded = RoundTripShapes(slidePart);
        var reloadedFill = reloaded[0].FillFormat;

        reloadedFill.FillType.Should().Be(FillType.NoFill);
    }

    /// <summary>
    /// Pattern fill type persists after XML round-trip.
    /// Note: PatternFormat class has minimal implementation, so we verify FillType only.
    /// </summary>
    [Fact]
    public void PatternFill_TypePersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);

        shape.FillFormat.FillType = FillType.Pattern;

        var reloaded = RoundTripShapes(slidePart);
        var reloadedFill = reloaded[0].FillFormat;

        reloadedFill.FillType.Should().Be(FillType.Pattern);
    }

    /// <summary>
    /// Picture fill type persists after XML round-trip.
    /// Note: PictureFillFormat has minimal implementation; verifies type round-trips.
    /// </summary>
    [Fact]
    public void PictureFill_TypePersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);

        shape.FillFormat.FillType = FillType.Picture;

        var reloaded = RoundTripShapes(slidePart);
        var reloadedFill = reloaded[0].FillFormat;

        reloadedFill.FillType.Should().Be(FillType.Picture);
    }

    /// <summary>
    /// Gradient stops can be added and count persists after round-trip.
    /// </summary>
    [Fact]
    public void GradientFill_StopsPersistAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 300, 150);

        shape.FillFormat.FillType = FillType.Gradient;
        var gf = shape.FillFormat.GradientFormat;
        gf.LinearGradientAngle = 45;
        gf.GradientStops.Add(0.0f, Color.Blue);
        gf.GradientStops.Add(1.0f, Color.Red);

        var reloaded = RoundTripShapes(slidePart);
        var reloadedFill = reloaded[0].FillFormat;

        reloadedFill.FillType.Should().Be(FillType.Gradient);
        reloadedFill.GradientFormat.GradientStops.Count.Should().BeGreaterThanOrEqualTo(2);
    }

    /// <summary>
    /// Pattern fill style persists after XML round-trip.
    /// </summary>
    [Fact]
    public void PatternFill_StylePersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);

        shape.FillFormat.FillType = FillType.Pattern;
        var pf = shape.FillFormat.PatternFormat;
        pf.PatternStyle = PatternStyle.Percent50;
        pf.ForeColor.Color = Color.DarkBlue;
        pf.BackColor.Color = Color.LightYellow;

        var reloaded = RoundTripShapes(slidePart);
        var reloadedFill = reloaded[0].FillFormat;

        reloadedFill.FillType.Should().Be(FillType.Pattern);
        reloadedFill.PatternFormat.PatternStyle.Should().Be(PatternStyle.Percent50);
    }

    /// <summary>
    /// Setting FillType preserves existing fill element when type matches.
    /// </summary>
    [Fact]
    public void FillType_PreservesExistingWhenTypeMatches()
    {
        var (_, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);

        shape.FillFormat.FillType = FillType.Solid;
        shape.FillFormat.SolidFillColor.Color = Color.Red;

        // Setting same type again should preserve color
        shape.FillFormat.FillType = FillType.Solid;
        var c = shape.FillFormat.SolidFillColor.Color;
        c.Should().NotBeNull();
        c!.R.Should().Be(255);
    }
}
