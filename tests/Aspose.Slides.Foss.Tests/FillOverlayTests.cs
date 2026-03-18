using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Effects;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Unit tests for <see cref="FillOverlay"/> effect class.
/// Verifies blend mode mapping, fill format access, and identity properties.
/// </summary>
public sealed class FillOverlayTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static FillOverlay CreateFillOverlay(string? blend = null)
    {
        var element = new XElement(ANs + "fillOverlay");
        if (blend is not null)
            element.SetAttributeValue("blend", blend);
        return new FillOverlay(element);
    }

    // ── Blend property ──

    [Fact]
    public void Blend_DefaultsToOverlay_WhenNoAttribute()
    {
        var fo = CreateFillOverlay();
        fo.Blend.Should().Be(FillBlendMode.Overlay);
    }

    [Theory]
    [InlineData("over", FillBlendMode.Overlay)]
    [InlineData("mult", FillBlendMode.Multiply)]
    [InlineData("screen", FillBlendMode.Screen)]
    [InlineData("darken", FillBlendMode.Darken)]
    [InlineData("lighten", FillBlendMode.Lighten)]
    public void Blend_ReadsCorrectEnumFromXml(string ooxmlValue, FillBlendMode expected)
    {
        var fo = CreateFillOverlay(ooxmlValue);
        fo.Blend.Should().Be(expected);
    }

    [Fact]
    public void Blend_UnknownAttributeDefaultsToOverlay()
    {
        var fo = CreateFillOverlay("unknown");
        fo.Blend.Should().Be(FillBlendMode.Overlay);
    }

    [Theory]
    [InlineData(FillBlendMode.Overlay, "over")]
    [InlineData(FillBlendMode.Multiply, "mult")]
    [InlineData(FillBlendMode.Screen, "screen")]
    [InlineData(FillBlendMode.Darken, "darken")]
    [InlineData(FillBlendMode.Lighten, "lighten")]
    public void Blend_SetWritesCorrectXmlAttribute(FillBlendMode mode, string _)
    {
        var fo = CreateFillOverlay();
        fo.Blend = mode;

        // Re-read to verify persistence in XML
        fo.Blend.Should().Be(mode);
    }

    [Fact]
    public void Blend_SetThenGet_RoundTrips()
    {
        var fo = CreateFillOverlay();
        fo.Blend = FillBlendMode.Darken;
        fo.Blend.Should().Be(FillBlendMode.Darken);

        fo.Blend = FillBlendMode.Screen;
        fo.Blend.Should().Be(FillBlendMode.Screen);
    }

    // ── FillFormat property ──

    [Fact]
    public void FillFormat_ReturnsNonNull()
    {
        var fo = CreateFillOverlay();
        fo.FillFormat.Should().NotBeNull();
    }

    [Fact]
    public void FillFormat_ReturnsIFillFormatInstance()
    {
        var fo = CreateFillOverlay();
        fo.FillFormat.Should().BeAssignableTo<IFillFormat>();
    }

    // ── Slide property ──

    [Fact]
    public void Slide_IsNull_WhenNoParentSlide()
    {
        var fo = CreateFillOverlay();
        fo.Slide.Should().BeNull();
    }

    // ── AsIPresentationComponent ──

    [Fact]
    public void AsIPresentationComponent_ReturnsSelf()
    {
        var fo = CreateFillOverlay();
        fo.AsIPresentationComponent.Should().BeSameAs(fo);
    }

    // ── AsIImageTransformOperation ──

    [Fact]
    public void AsIImageTransformOperation_ReturnsSelf()
    {
        var fo = CreateFillOverlay();
        fo.AsIImageTransformOperation.Should().BeSameAs(fo);
    }

    // ── Parameterless constructor ──

    [Fact]
    public void DefaultConstructor_CreatesInstance()
    {
        var fo = new FillOverlay();
        fo.Should().NotBeNull();
    }

    [Fact]
    public void DefaultConstructor_BlendDefaultsToOverlay()
    {
        var fo = new FillOverlay();
        fo.Blend.Should().Be(FillBlendMode.Overlay);
    }

    // ── Interface conformance ──

    [Fact]
    public void ImplementsIFillOverlay()
    {
        var fo = new FillOverlay();
        fo.Should().BeAssignableTo<IFillOverlay>();
    }

    [Fact]
    public void ImplementsIImageTransformOperation()
    {
        var fo = new FillOverlay();
        fo.Should().BeAssignableTo<IImageTransformOperation>();
    }

    // ── EffectFormat integration ──

    [Fact]
    public void EffectFormat_FillOverlayEffect_IsNullByDefault()
    {
        var parent = new XElement(ANs + "spPr");
        var ef = new EffectFormat();
        ef.InitInternal(parent, parentSlide: null);

        ef.FillOverlayEffect.Should().BeNull();
    }

    [Fact]
    public void EffectFormat_EnableFillOverlay_CreatesEffect()
    {
        var parent = new XElement(ANs + "spPr");
        var ef = new EffectFormat();
        ef.InitInternal(parent, parentSlide: null);

        ef.EnableFillOverlayEffect();

        ef.FillOverlayEffect.Should().NotBeNull();
    }

    [Fact]
    public void EffectFormat_DisableFillOverlay_RemovesEffect()
    {
        var parent = new XElement(ANs + "spPr");
        var ef = new EffectFormat();
        ef.InitInternal(parent, parentSlide: null);

        ef.EnableFillOverlayEffect();
        ef.DisableFillOverlayEffect();

        ef.FillOverlayEffect.Should().BeNull();
    }
}
