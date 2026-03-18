using Aspose.Slides.Foss.Drawing;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests for Color: construction, FromArgb, component access, named constants, equality.
/// test_line_format (line colors), test_effect_format (shadow/glow colors),
/// and test_text_formatting (font colors).
/// </summary>
public sealed class ColorTests
{
    [Fact]
    public void Constructor_DefaultAlpha_Is255()
    {
        var color = new Color(r: 0, g: 128, b: 255);

        color.A.Should().Be(255);
        color.R.Should().Be(0);
        color.G.Should().Be(128);
        color.B.Should().Be(255);
    }

    [Fact]
    public void Constructor_AllDefaults_IsBlackOpaque()
    {
        var color = new Color();

        color.A.Should().Be(255);
        color.R.Should().Be(0);
        color.G.Should().Be(0);
        color.B.Should().Be(0);
    }

    [Fact]
    public void Constructor_ExplicitAlpha_IsPreserved()
    {
        // From test_outer_shadow: Color.from_argb(128, 0, 0, 0) — semi-transparent black
        var color = new Color(a: 128, r: 0, g: 0, b: 0);

        color.A.Should().Be(128);
        color.R.Should().Be(0);
        color.G.Should().Be(0);
        color.B.Should().Be(0);
    }

    [Fact]
    public void FromArgb_ReturnsCorrectComponents()
    {
        // From test_solid_fill: Color.from_argb(255, 0, 128, 255)
        var color = Color.FromArgb(255, 0, 128, 255);

        color.A.Should().Be(255);
        color.R.Should().Be(0);
        color.G.Should().Be(128);
        color.B.Should().Be(255);
    }

    [Fact]
    public void FromArgb_SemiTransparent_PreservesAlpha()
    {
        // From test_outer_shadow: shadow color with alpha 128
        var color = Color.FromArgb(128, 0, 0, 0);

        color.A.Should().Be(128);
    }

    [Fact]
    public void Red_ComponentsMatchExpected()
    {
        // From test_font_color: asserts c.r == 255 and c.g == 0 and c.b == 0
        Color.Red.R.Should().Be(255);
        Color.Red.G.Should().Be(0);
        Color.Red.B.Should().Be(0);
        Color.Red.A.Should().Be(255);
    }

    [Fact]
    public void DarkRed_ComponentsMatchExpected()
    {
        // From test_line_color_and_width: asserts c.r == Color.dark_red.r
        Color.DarkRed.R.Should().Be(139);
        Color.DarkRed.G.Should().Be(0);
        Color.DarkRed.B.Should().Be(0);
    }

    [Fact]
    public void Blue_ComponentsMatchExpected()
    {
        // From test_gradient_fill: uses Color.blue as gradient stop
        Color.Blue.R.Should().Be(0);
        Color.Blue.G.Should().Be(0);
        Color.Blue.B.Should().Be(255);
    }

    [Fact]
    public void Gold_ComponentsMatchExpected()
    {
        // From test_glow: glow color set to Color.gold
        Color.Gold.R.Should().Be(255);
        Color.Gold.G.Should().Be(215);
        Color.Gold.B.Should().Be(0);
    }

    [Fact]
    public void DarkBlue_ComponentsMatchExpected()
    {
        // From test_pattern_fill: fore_color set to Color.dark_blue
        Color.DarkBlue.R.Should().Be(0);
        Color.DarkBlue.G.Should().Be(0);
        Color.DarkBlue.B.Should().Be(139);
    }

    [Fact]
    public void LightYellow_ComponentsMatchExpected()
    {
        // From test_pattern_fill: back_color set to Color.light_yellow
        Color.LightYellow.R.Should().Be(255);
        Color.LightYellow.G.Should().Be(255);
        Color.LightYellow.B.Should().Be(224);
    }

    [Fact]
    public void Black_ComponentsMatchExpected()
    {
        // From test_line_dash_style: line color set to Color.black
        Color.Black.R.Should().Be(0);
        Color.Black.G.Should().Be(0);
        Color.Black.B.Should().Be(0);
        Color.Black.A.Should().Be(255);
    }

    [Fact]
    public void LightBlue_ComponentsMatchExpected()
    {
        // From test_cell_fill: cell fill set to Color.light_blue
        Color.LightBlue.R.Should().Be(173);
        Color.LightBlue.G.Should().Be(216);
        Color.LightBlue.B.Should().Be(230);
    }

    [Fact]
    public void Equals_SameComponents_ReturnsTrue()
    {
        var a = new Color(r: 100, g: 200, b: 50);
        var b = new Color(r: 100, g: 200, b: 50);

        a.Should().Be(b);
    }

    [Fact]
    public void Equals_DifferentComponents_ReturnsFalse()
    {
        var a = new Color(r: 100, g: 200, b: 50);
        var b = new Color(r: 100, g: 200, b: 51);

        a.Should().NotBe(b);
    }

    [Fact]
    public void Equals_DifferentAlpha_ReturnsFalse()
    {
        var a = Color.FromArgb(255, 0, 0, 0);
        var b = Color.FromArgb(128, 0, 0, 0);

        a.Should().NotBe(b);
    }

    [Fact]
    public void GetHashCode_EqualColors_SameHash()
    {
        var a = Color.FromArgb(128, 10, 20, 30);
        var b = Color.FromArgb(128, 10, 20, 30);

        a.GetHashCode().Should().Be(b.GetHashCode());
    }

    [Fact]
    public void FromArgb_EqualsConstructor()
    {
        var fromFactory = Color.FromArgb(128, 0, 0, 0);
        var fromCtor = new Color(a: 128, r: 0, g: 0, b: 0);

        fromFactory.Should().Be(fromCtor);
    }

    [Fact]
    public void Transparent_HasZeroAlpha()
    {
        Color.Transparent.A.Should().Be(0);
    }

    [Fact]
    public void Empty_HasAllZeros()
    {
        Color.Empty.A.Should().Be(0);
        Color.Empty.R.Should().Be(0);
        Color.Empty.G.Should().Be(0);
        Color.Empty.B.Should().Be(0);
    }

    [Theory]
    [InlineData(255, 0, 128, 255)]
    [InlineData(128, 0, 0, 0)]
    [InlineData(0, 255, 255, 255)]
    public void FromArgb_RoundTrips_AllComponents(int a, int r, int g, int b)
    {
        var color = Color.FromArgb(a, r, g, b);

        color.A.Should().Be(a);
        color.R.Should().Be(r);
        color.G.Should().Be(g);
        color.B.Should().Be(b);
    }
}
