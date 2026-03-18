using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Internal;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Integration tests for EffectFormat round-trip through XML serialization.
/// </summary>
public sealed class EffectFormatIntegrationTests
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
    /// Outer shadow effect persists after XML round-trip.
    /// </summary>
    [Fact]
    public void OuterShadow_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);

        shape.EffectFormat.EnableOuterShadowEffect();

        var reloaded = RoundTripShapes(slidePart);
        var reloadedEf = reloaded[0].EffectFormat;

        reloadedEf.OuterShadowEffect.Should().NotBeNull();
        reloadedEf.IsNoEffects.Should().BeFalse();
    }

    /// <summary>
    /// Glow effect persists after XML round-trip.
    /// </summary>
    [Fact]
    public void GlowEffect_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);

        shape.EffectFormat.EnableGlowEffect();

        var reloaded = RoundTripShapes(slidePart);
        var reloadedEf = reloaded[0].EffectFormat;

        reloadedEf.GlowEffect.Should().NotBeNull();
        reloadedEf.IsNoEffects.Should().BeFalse();
    }

    /// <summary>
    /// Soft edge effect persists after XML round-trip.
    /// </summary>
    [Fact]
    public void SoftEdgeEffect_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);

        shape.EffectFormat.EnableSoftEdgeEffect();

        var reloaded = RoundTripShapes(slidePart);
        var reloadedEf = reloaded[0].EffectFormat;

        reloadedEf.SoftEdgeEffect.Should().NotBeNull();
        reloadedEf.IsNoEffects.Should().BeFalse();
    }

    /// <summary>
    /// Blur effect with radius persists after XML round-trip.
    /// </summary>
    [Fact]
    public void BlurEffect_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);

        shape.EffectFormat.SetBlurEffect(8.0f, true);

        var reloaded = RoundTripShapes(slidePart);
        var reloadedEf = reloaded[0].EffectFormat;

        reloadedEf.BlurEffect.Should().NotBeNull();
        reloadedEf.BlurEffect!.Radius.Should().Be(8.0f);
        reloadedEf.BlurEffect!.Grow.Should().BeTrue();
        reloadedEf.IsNoEffects.Should().BeFalse();
    }

    /// <summary>
    /// Blur effect with grow=false persists after XML round-trip.
    /// </summary>
    [Fact]
    public void BlurEffect_GrowFalse_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);

        shape.EffectFormat.SetBlurEffect(5.0f, false);

        var reloaded = RoundTripShapes(slidePart);
        var blur = reloaded[0].EffectFormat.BlurEffect;

        blur.Should().NotBeNull();
        blur!.Radius.Should().Be(5.0f);
        blur!.Grow.Should().BeFalse();
    }

    /// <summary>
    /// Blur radius defaults to 0 when rad attribute is absent.
    /// </summary>
    [Fact]
    public void BlurEffect_DefaultRadius_IsZero()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);

        shape.EffectFormat.EnableBlurEffect();

        var blur = shape.EffectFormat.BlurEffect;
        blur.Should().NotBeNull();
        blur!.Radius.Should().Be(0f);
        blur!.Grow.Should().BeTrue(); // default when attribute absent
    }

    /// <summary>
    /// Effects can be enabled then disabled, and the disabled state persists.
    /// </summary>
    [Fact]
    public void EnableDisableEffects_DisabledStatePersists()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);

        var ef = shape.EffectFormat;
        ef.EnableOuterShadowEffect();
        ef.EnableGlowEffect();
        ef.EnableSoftEdgeEffect();
        ef.IsNoEffects.Should().BeFalse();

        ef.DisableOuterShadowEffect();
        ef.DisableGlowEffect();
        ef.DisableSoftEdgeEffect();
        ef.IsNoEffects.Should().BeTrue();

        var reloaded = RoundTripShapes(slidePart);
        var reloadedEf = reloaded[0].EffectFormat;

        reloadedEf.IsNoEffects.Should().BeTrue();
        reloadedEf.OuterShadowEffect.Should().BeNull();
        reloadedEf.GlowEffect.Should().BeNull();
        reloadedEf.SoftEdgeEffect.Should().BeNull();
    }
}
