using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using Aspose.Slides.Foss.Effects;
using Aspose.Slides.Foss.Internal;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Unit tests for the Glow effect class.
/// </summary>
public sealed class GlowTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

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

    [Fact]
    public void Radius_DefaultsToZero_WhenAttributeAbsent()
    {
        var element = new XElement(ANs + "glow");
        var glow = new Glow(element);

        glow.Radius.Should().Be(0f);
    }

    [Fact]
    public void Radius_ReadsFromEmuAttribute()
    {
        // 15 points * 12700 EMU/point = 190500 EMU
        var element = new XElement(ANs + "glow", new XAttribute("rad", "190500"));
        var glow = new Glow(element);

        glow.Radius.Should().Be(15f);
    }

    [Fact]
    public void Radius_WritesEmuAttribute()
    {
        var element = new XElement(ANs + "glow");
        var glow = new Glow(element);

        glow.Radius = 15f;

        element.Attribute("rad")!.Value.Should().Be("190500");
    }

    [Fact]
    public void Color_ReturnsColorFormat()
    {
        var element = new XElement(ANs + "glow");
        var glow = new Glow(element);

        glow.Color.Should().NotBeNull();
        glow.Color.Should().BeAssignableTo<IColorFormat>();
    }

    [Fact]
    public void Color_CanSetAndReadSrgbColor()
    {
        var element = new XElement(ANs + "glow");
        var glow = new Glow(element);

        glow.Color.Color = new Color(a: 255, r: 255, g: 215, b: 0); // Gold

        var readColor = glow.Color.Color;
        readColor!.R.Should().Be(255);
        readColor.G.Should().Be(215);
        readColor.B.Should().Be(0);
    }

    [Fact]
    public void AsIImageTransformOperation_ReturnsSelf()
    {
        var glow = new Glow();

        glow.AsIImageTransformOperation.Should().BeSameAs(glow);
    }

    [Fact]
    public void Glow_ImplementsIGlow()
    {
        var glow = new Glow();

        glow.Should().BeAssignableTo<IGlow>();
    }

    [Fact]
    public void Glow_ImplementsIImageTransformOperation()
    {
        var glow = new Glow();

        glow.Should().BeAssignableTo<IImageTransformOperation>();
    }

    [Fact]
    public void GlowEffect_RadiusPersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 200);

        shape.EffectFormat.EnableGlowEffect();
        shape.EffectFormat.GlowEffect!.Radius = 15f;

        var reloaded = RoundTripShapes(slidePart);
        var glow = reloaded[0].EffectFormat.GlowEffect;

        glow.Should().NotBeNull();
        glow!.Radius.Should().Be(15f);
    }

    [Fact]
    public void GlowEffect_ColorPersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 200);

        shape.EffectFormat.EnableGlowEffect();
        shape.EffectFormat.GlowEffect!.Radius = 15f;
        shape.EffectFormat.GlowEffect!.Color.Color = new Color(a: 255, r: 255, g: 215, b: 0);

        var reloaded = RoundTripShapes(slidePart);
        var glow = reloaded[0].EffectFormat.GlowEffect;

        glow.Should().NotBeNull();
        var color = glow!.Color.Color;
        color!.R.Should().Be(255);
        color.G.Should().Be(215);
        color.B.Should().Be(0);
    }

    /// <summary>
    /// A glow of radius zero is a glow nobody can see, so enabling one must not leave the radius
    /// unset. <c>rad</c> is optional in CT_GlowEffect and defaults to 0, which is why the enabler
    /// writes an explicit value.
    /// </summary>
    [Fact]
    public void GlowEffect_AfterEnable_HasAVisibleDefaultRadius()
    {
        var (_, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);

        shape.EffectFormat.EnableGlowEffect();

        var glow = shape.EffectFormat.GlowEffect;
        glow.Should().NotBeNull();
        glow!.Radius.Should().Be(5f);
    }
}
