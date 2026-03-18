using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests for ColorFormat: Color, ColorType, R/G/B, FloatR/FloatG/FloatB,
/// PresetColor, SchemeColor properties.
/// </summary>
public sealed class ColorFormatTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static ColorFormat CreateColorFormat(XElement? parentElement = null)
    {
        var element = parentElement ?? new XElement(ANs + "solidFill");
        var cf = new ColorFormat();
        cf.InitInternal(element, parentSlide: null);
        return cf;
    }

    // --- ColorType property ---

    [Fact]
    public void ColorType_DefaultIsNotDefinedWhenNoColorElement()
    {
        var cf = CreateColorFormat();
        cf.ColorType.Should().Be(ColorType.NotDefined);
    }

    [Fact]
    public void ColorType_NotDefinedSentinelIsZero()
    {
        ((int)ColorType.NotDefined).Should().Be(0);
    }

    [Theory]
    [InlineData(ColorType.RGB)]
    [InlineData(ColorType.RGBPercentage)]
    [InlineData(ColorType.HSL)]
    [InlineData(ColorType.Scheme)]
    [InlineData(ColorType.System)]
    [InlineData(ColorType.Preset)]
    public void ColorType_AllValuesAreDefined(ColorType colorType)
    {
        Enum.IsDefined(colorType).Should().BeTrue();
    }

    [Fact]
    public void ColorType_CanBeSetToRgb()
    {
        var cf = CreateColorFormat();
        cf.ColorType = ColorType.RGB;
        cf.ColorType.Should().Be(ColorType.RGB);
    }

    [Fact]
    public void ColorType_CanBeSetToScheme()
    {
        var cf = CreateColorFormat();
        cf.ColorType = ColorType.Scheme;
        cf.ColorType.Should().Be(ColorType.Scheme);
    }

    [Fact]
    public void ColorType_CanBeSetToPreset()
    {
        var cf = CreateColorFormat();
        cf.ColorType = ColorType.Preset;
        cf.ColorType.Should().Be(ColorType.Preset);
    }

    [Fact]
    public void ColorType_IsRgbWhenSrgbClrElementPresent()
    {
        var parent = new XElement(ANs + "solidFill",
            new XElement(ANs + "srgbClr", new XAttribute("val", "FF0000")));
        var cf = CreateColorFormat(parent);
        cf.ColorType.Should().Be(ColorType.RGB);
    }

    [Fact]
    public void ColorType_IsSchemeWhenSchemeClrElementPresent()
    {
        var parent = new XElement(ANs + "solidFill",
            new XElement(ANs + "schemeClr", new XAttribute("val", "accent1")));
        var cf = CreateColorFormat(parent);
        cf.ColorType.Should().Be(ColorType.Scheme);
    }

    [Fact]
    public void ColorType_IsPresetWhenPrstClrElementPresent()
    {
        var parent = new XElement(ANs + "solidFill",
            new XElement(ANs + "prstClr", new XAttribute("val", "red")));
        var cf = CreateColorFormat(parent);
        cf.ColorType.Should().Be(ColorType.Preset);
    }

    // --- Color property (get/set with srgbClr XML) ---

    [Fact]
    public void Color_CanBeSetAndReadBack()
    {
        var cf = CreateColorFormat();
        cf.Color = Color.FromArgb(255, 0, 128, 255);

        var c = cf.Color;
        c.Should().NotBeNull();
        c!.R.Should().Be(0);
        c.G.Should().Be(128);
        c.B.Should().Be(255);
    }

    [Fact]
    public void Color_RedColorPersistsInXml()
    {
        var cf = CreateColorFormat();
        cf.Color = Color.Red;

        var c = cf.Color;
        c.Should().NotBeNull();
        c!.R.Should().Be(255);
        c.G.Should().Be(0);
        c.B.Should().Be(0);
    }

    [Fact]
    public void Color_SemiTransparentBlackPersistsAlpha()
    {
        var cf = CreateColorFormat();
        cf.Color = Color.FromArgb(128, 0, 0, 0);

        var c = cf.Color;
        c.Should().NotBeNull();
        c!.R.Should().Be(0);
        c.G.Should().Be(0);
        c.B.Should().Be(0);
        c.A.Should().BeCloseTo(128, 1);
    }

    [Fact]
    public void Color_GoldColorPersistsInXml()
    {
        var cf = CreateColorFormat();
        cf.Color = Color.Gold;

        var c = cf.Color;
        c.Should().NotBeNull();
        c!.R.Should().Be(255);
        c.G.Should().Be(215);
        c.B.Should().Be(0);
    }

    [Fact]
    public void Color_ReadsFromExistingSrgbClrElement()
    {
        var parent = new XElement(ANs + "solidFill",
            new XElement(ANs + "srgbClr", new XAttribute("val", "0080FF")));
        var cf = CreateColorFormat(parent);

        var c = cf.Color;
        c.Should().NotBeNull();
        c!.R.Should().Be(0);
        c.G.Should().Be(128);
        c.B.Should().Be(255);
    }

    [Fact]
    public void Color_SettingColorReplacesExistingElement()
    {
        var parent = new XElement(ANs + "solidFill",
            new XElement(ANs + "srgbClr", new XAttribute("val", "0000FF")));
        var cf = CreateColorFormat(parent);

        cf.Color = Color.Red;

        var c = cf.Color;
        c.Should().NotBeNull();
        c!.R.Should().Be(255);
        c.G.Should().Be(0);
        c.B.Should().Be(0);
    }

    // --- R, G, B properties ---

    [Fact]
    public void R_ReturnsRedComponentFromSrgbClr()
    {
        var parent = new XElement(ANs + "solidFill",
            new XElement(ANs + "srgbClr", new XAttribute("val", "8B0000")));
        var cf = CreateColorFormat(parent);
        cf.R.Should().Be(139);
    }

    [Fact]
    public void G_ReturnsGreenComponentFromSrgbClr()
    {
        var parent = new XElement(ANs + "solidFill",
            new XElement(ANs + "srgbClr", new XAttribute("val", "008000")));
        var cf = CreateColorFormat(parent);
        cf.G.Should().Be(128);
    }

    [Fact]
    public void B_ReturnsBlueComponentFromSrgbClr()
    {
        var parent = new XElement(ANs + "solidFill",
            new XElement(ANs + "srgbClr", new XAttribute("val", "0000FF")));
        var cf = CreateColorFormat(parent);
        cf.B.Should().Be(255);
    }

    [Fact]
    public void RGB_AllZeroWhenNoColorElement()
    {
        var cf = CreateColorFormat();
        cf.R.Should().Be(0);
        cf.G.Should().Be(0);
        cf.B.Should().Be(0);
    }

    [Fact]
    public void RGB_MatchColorPropertyAfterSet()
    {
        var cf = CreateColorFormat();
        cf.Color = Color.FromArgb(255, 0, 128, 255);

        cf.R.Should().Be(0);
        cf.G.Should().Be(128);
        cf.B.Should().Be(255);
    }

    // --- FloatR, FloatG, FloatB properties ---

    [Fact]
    public void FloatR_ReturnsNormalizedRedComponent()
    {
        var parent = new XElement(ANs + "solidFill",
            new XElement(ANs + "srgbClr", new XAttribute("val", "FF0000")));
        var cf = CreateColorFormat(parent);
        cf.FloatR.Should().BeApproximately(1.0f, 0.01f);
    }

    [Fact]
    public void FloatG_ReturnsNormalizedGreenComponent()
    {
        var parent = new XElement(ANs + "solidFill",
            new XElement(ANs + "srgbClr", new XAttribute("val", "00FF00")));
        var cf = CreateColorFormat(parent);
        cf.FloatG.Should().BeApproximately(1.0f, 0.01f);
    }

    [Fact]
    public void FloatB_ReturnsNormalizedBlueComponent()
    {
        var parent = new XElement(ANs + "solidFill",
            new XElement(ANs + "srgbClr", new XAttribute("val", "0000FF")));
        var cf = CreateColorFormat(parent);
        cf.FloatB.Should().BeApproximately(1.0f, 0.01f);
    }

    [Fact]
    public void FloatRGB_AllZeroWhenNoColorElement()
    {
        var cf = CreateColorFormat();
        cf.FloatR.Should().Be(0f);
        cf.FloatG.Should().Be(0f);
        cf.FloatB.Should().Be(0f);
    }

    [Fact]
    public void FloatRGB_HalfValueReturnsApproximatelyHalf()
    {
        var parent = new XElement(ANs + "solidFill",
            new XElement(ANs + "srgbClr", new XAttribute("val", "808080")));
        var cf = CreateColorFormat(parent);
        cf.FloatR.Should().BeApproximately(128 / 255f, 0.01f);
        cf.FloatG.Should().BeApproximately(128 / 255f, 0.01f);
        cf.FloatB.Should().BeApproximately(128 / 255f, 0.01f);
    }

    // --- PresetColor property ---

    [Fact]
    public void PresetColor_NotDefinedSentinelIsZero()
    {
        ((int)PresetColor.NotDefined).Should().Be(0);
    }

    [Fact]
    public void PresetColor_ReturnsNotDefinedWhenNoPresetElement()
    {
        var cf = CreateColorFormat();
        cf.PresetColor.Should().Be(PresetColor.NotDefined);
    }

    [Fact]
    public void PresetColor_ReadsRedFromPrstClrElement()
    {
        var parent = new XElement(ANs + "solidFill",
            new XElement(ANs + "prstClr", new XAttribute("val", "red")));
        var cf = CreateColorFormat(parent);
        cf.PresetColor.Should().Be(PresetColor.Red);
    }

    [Fact]
    public void PresetColor_CanBeSetAndReadBack()
    {
        var cf = CreateColorFormat();
        cf.PresetColor = PresetColor.Gold;
        cf.PresetColor.Should().Be(PresetColor.Gold);
    }

    [Theory]
    [InlineData(PresetColor.Red)]
    [InlineData(PresetColor.Blue)]
    [InlineData(PresetColor.Gold)]
    [InlineData(PresetColor.DarkBlue)]
    [InlineData(PresetColor.LightYellow)]
    [InlineData(PresetColor.Black)]
    public void PresetColor_CommonValuesAreDefined(PresetColor preset)
    {
        Enum.IsDefined(preset).Should().BeTrue();
        preset.Should().NotBe(PresetColor.NotDefined);
    }

    // --- SchemeColor property ---

    [Fact]
    public void SchemeColor_NotDefinedSentinelIsZero()
    {
        ((int)SchemeColor.NotDefined).Should().Be(0);
    }

    [Fact]
    public void SchemeColor_ReturnsNotDefinedWhenNoSchemeElement()
    {
        var cf = CreateColorFormat();
        cf.SchemeColor.Should().Be(SchemeColor.NotDefined);
    }

    [Fact]
    public void SchemeColor_ReadsAccent1FromSchemeClrElement()
    {
        var parent = new XElement(ANs + "solidFill",
            new XElement(ANs + "schemeClr", new XAttribute("val", "accent1")));
        var cf = CreateColorFormat(parent);
        cf.SchemeColor.Should().Be(SchemeColor.Accent1);
    }

    [Fact]
    public void SchemeColor_CanBeSetAndReadBack()
    {
        var cf = CreateColorFormat();
        cf.SchemeColor = SchemeColor.Accent1;
        cf.SchemeColor.Should().Be(SchemeColor.Accent1);
    }

    [Theory]
    [InlineData(SchemeColor.Background1)]
    [InlineData(SchemeColor.Text1)]
    [InlineData(SchemeColor.Accent1)]
    [InlineData(SchemeColor.Accent2)]
    [InlineData(SchemeColor.Hyperlink)]
    [InlineData(SchemeColor.Dark1)]
    [InlineData(SchemeColor.Light1)]
    public void SchemeColor_CommonValuesAreDefined(SchemeColor scheme)
    {
        Enum.IsDefined(scheme).Should().BeTrue();
        scheme.Should().NotBe(SchemeColor.NotDefined);
    }

    // --- ColorFormat implements IColorFormat ---

    [Fact]
    public void ColorFormat_ImplementsIColorFormat()
    {
        var cf = new ColorFormat();
        cf.Should().BeAssignableTo<IColorFormat>();
    }

    // --- Verify Color enum values used across tests ---

    [Fact]
    public void Color_DarkRedHasRedComponent()
    {
        Color.DarkRed.Should().NotBe(Color.Empty);
        Color.DarkRed.R.Should().BeGreaterThan(0);
    }

    [Fact]
    public void Color_DarkBlueHasBlueComponent()
    {
        Color.DarkBlue.Should().NotBe(Color.Empty);
        Color.DarkBlue.B.Should().BeGreaterThan(0);
    }

    [Fact]
    public void Color_LightYellowExists()
    {
        Color.LightYellow.Should().NotBe(Color.Empty);
    }

    [Fact]
    public void Color_LightBlueHasBlueComponent()
    {
        Color.LightBlue.Should().NotBe(Color.Empty);
        Color.LightBlue.B.Should().BeGreaterThan(0);
    }

    [Fact]
    public void Color_BlackHasZeroRgbComponents()
    {
        Color.Black.R.Should().Be(0);
        Color.Black.G.Should().Be(0);
        Color.Black.B.Should().Be(0);
    }
}
