using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Effects;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests for the PresetShadow class backed by XML elements.
/// </summary>
public sealed class PresetShadowTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static PresetShadow CreateShadow(string attributes = "")
    {
        var xml = $"<a:prstShdw xmlns:a=\"{ANs}\" {attributes}/>";
        var element = XElement.Parse(xml);
        return new PresetShadow(element);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void Direction_SetAndReadBack()
    {
        var shadow = CreateShadow();
        shadow.Direction = 315;
        shadow.Direction.Should().Be(315);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void Distance_SetAndReadBack()
    {
        var shadow = CreateShadow();
        shadow.Distance = 8;
        shadow.Distance.Should().Be(8);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void MultipleProperties_SetTogether()
    {
        var shadow = CreateShadow();
        shadow.Direction = 315;
        shadow.Distance = 8;
        shadow.Preset = PresetShadowType.TopLeftDropShadow;

        shadow.Direction.Should().Be(315);
        shadow.Distance.Should().Be(8);
        shadow.Preset.Should().Be(PresetShadowType.TopLeftDropShadow);
    }

    [Fact]
    public void Direction_DefaultIsZero()
    {
        var shadow = CreateShadow();
        shadow.Direction.Should().Be(0);
    }

    [Fact]
    public void Distance_DefaultIsZero()
    {
        var shadow = CreateShadow();
        shadow.Distance.Should().Be(0);
    }

    [Fact]
    public void Preset_DefaultIsTopLeftDropShadow()
    {
        var shadow = CreateShadow();
        shadow.Preset.Should().Be(PresetShadowType.TopLeftDropShadow);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void Preset_SetAndReadBack()
    {
        var shadow = CreateShadow();
        shadow.Preset = PresetShadowType.BottomRightDropShadow;
        shadow.Preset.Should().Be(PresetShadowType.BottomRightDropShadow);
    }

    /// <summary>
    /// Verifies OOXML preset values map correctly from XML attributes.
    /// </summary>
    [Theory]
    [InlineData("shdw1", PresetShadowType.TopLeftDropShadow)]
    [InlineData("shdw6", PresetShadowType.BottomRightSmallDropShadow)]
    [InlineData("shdw12", PresetShadowType.TopRightDropShadow)]
    [InlineData("shdw16", PresetShadowType.BottomLeftDropShadow)]
    [InlineData("shdw17", PresetShadowType.BottomRightDropShadow)]
    [InlineData("shdw20", PresetShadowType.TopLeftSmallDropShadow)]
    public void Preset_ReadsOoxmlValues(string ooxmlVal, PresetShadowType expected)
    {
        var shadow = CreateShadow($"prst=\"{ooxmlVal}\"");
        shadow.Preset.Should().Be(expected);
    }

    /// <summary>
    /// Verifies preset round-trips through set/get.
    /// </summary>
    [Theory]
    [InlineData(PresetShadowType.TopLeftDropShadow)]
    [InlineData(PresetShadowType.BottomRightSmallDropShadow)]
    [InlineData(PresetShadowType.FrontLeftLongPerspectiveShadow)]
    [InlineData(PresetShadowType.OuterBoxShadow3D)]
    [InlineData(PresetShadowType.TopLeftSmallDropShadow)]
    public void Preset_RoundTrips(PresetShadowType preset)
    {
        var shadow = CreateShadow();
        shadow.Preset = preset;
        shadow.Preset.Should().Be(preset);
    }

    [Fact]
    public void PresetShadow_IsImageTransformOperation()
    {
        var shadow = CreateShadow();
        shadow.Should().BeAssignableTo<IImageTransformOperation>();
    }

    [Fact]
    public void AsIImageTransformOperation_ReturnsSelf()
    {
        var shadow = CreateShadow();
        shadow.AsIImageTransformOperation.Should().BeSameAs(shadow);
    }

    [Fact]
    public void ShadowColor_IsNotNull()
    {
        var shadow = CreateShadow();
        shadow.ShadowColor.Should().NotBeNull();
    }

    /// <summary>
    /// Direction reads from the dir attribute in 60000ths of a degree.
    /// </summary>
    [Fact]
    public void Direction_ReadsFromXmlAttribute()
    {
        var shadow = CreateShadow("dir=\"18900000\"");
        shadow.Direction.Should().Be(315);
    }

    /// <summary>
    /// Distance reads from the dist attribute in EMUs (12700 per point).
    /// </summary>
    [Fact]
    public void Distance_ReadsFromXmlAttribute()
    {
        var shadow = CreateShadow("dist=\"101600\"");
        shadow.Distance.Should().Be(8);
    }
}
