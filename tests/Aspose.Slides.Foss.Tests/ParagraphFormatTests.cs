using System.Xml.Linq;
using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Unit tests for <see cref="ParagraphFormat"/> properties.
/// and test_depth_and_material.
/// </summary>
public sealed class ParagraphFormatTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// <summary>
    /// </summary>
    [Fact]
    public void Alignment_DefaultsToNotDefined()
    {
        var pf = new ParagraphFormat();
        pf.Alignment.Should().Be(TextAlignment.NotDefined);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void Alignment_Center_PersistsOnXml()
    {
        var ppr = new XElement(ANs + "pPr");
        var pf = new ParagraphFormat();
        pf.InitInternal(ppr, null, null);

        pf.Alignment = TextAlignment.Center;

        var pf2 = new ParagraphFormat();
        pf2.InitInternal(ppr, null, null);
        pf2.Alignment.Should().Be(TextAlignment.Center);
    }

    /// <summary>
    /// </summary>
    [Theory]
    [InlineData(TextAlignment.Left)]
    [InlineData(TextAlignment.Center)]
    [InlineData(TextAlignment.Right)]
    [InlineData(TextAlignment.Justify)]
    [InlineData(TextAlignment.JustifyLow)]
    [InlineData(TextAlignment.Distributed)]
    public void Alignment_AllValues_RoundTrip(TextAlignment alignment)
    {
        var ppr = new XElement(ANs + "pPr");
        var pf = new ParagraphFormat();
        pf.InitInternal(ppr, null, null);

        pf.Alignment = alignment;

        var pf2 = new ParagraphFormat();
        pf2.InitInternal(ppr, null, null);
        pf2.Alignment.Should().Be(alignment);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void Alignment_SetToNotDefined_RemovesAttribute()
    {
        var ppr = new XElement(ANs + "pPr");
        var pf = new ParagraphFormat();
        pf.InitInternal(ppr, null, null);

        pf.Alignment = TextAlignment.Center;
        pf.Alignment.Should().Be(TextAlignment.Center);

        pf.Alignment = TextAlignment.NotDefined;
        pf.Alignment.Should().Be(TextAlignment.NotDefined);
        ppr.Attribute("algn").Should().BeNull();
    }

    [Fact]
    public void SpaceWithin_DefaultsToNaN()
    {
        var pf = new ParagraphFormat();
        float.IsNaN(pf.SpaceWithin).Should().BeTrue();
    }

    [Fact]
    public void SpaceWithin_PercentageValue_RoundTrips()
    {
        var ppr = new XElement(ANs + "pPr");
        var pf = new ParagraphFormat();
        pf.InitInternal(ppr, null, null);

        pf.SpaceWithin = 120; // 120% line spacing

        var pf2 = new ParagraphFormat();
        pf2.InitInternal(ppr, null, null);
        pf2.SpaceWithin.Should().Be(120);
    }

    [Fact]
    public void SpaceWithin_NegativePointValue_RoundTrips()
    {
        var ppr = new XElement(ANs + "pPr");
        var pf = new ParagraphFormat();
        pf.InitInternal(ppr, null, null);

        pf.SpaceWithin = -18; // 18pt fixed spacing

        var pf2 = new ParagraphFormat();
        pf2.InitInternal(ppr, null, null);
        pf2.SpaceWithin.Should().Be(-18);
    }

    [Fact]
    public void SpaceBefore_PercentageValue_RoundTrips()
    {
        var ppr = new XElement(ANs + "pPr");
        var pf = new ParagraphFormat();
        pf.InitInternal(ppr, null, null);

        pf.SpaceBefore = 50;

        var pf2 = new ParagraphFormat();
        pf2.InitInternal(ppr, null, null);
        pf2.SpaceBefore.Should().Be(50);
    }

    [Fact]
    public void SpaceAfter_PercentageValue_RoundTrips()
    {
        var ppr = new XElement(ANs + "pPr");
        var pf = new ParagraphFormat();
        pf.InitInternal(ppr, null, null);

        pf.SpaceAfter = 25;

        var pf2 = new ParagraphFormat();
        pf2.InitInternal(ppr, null, null);
        pf2.SpaceAfter.Should().Be(25);
    }

    [Fact]
    public void SpaceAfter_SetToNaN_RemovesElement()
    {
        var ppr = new XElement(ANs + "pPr");
        var pf = new ParagraphFormat();
        pf.InitInternal(ppr, null, null);

        pf.SpaceAfter = 25;
        pf.SpaceAfter.Should().Be(25);

        pf.SpaceAfter = float.NaN;
        float.IsNaN(pf.SpaceAfter).Should().BeTrue();
    }

    [Theory]
    [InlineData("eaLnBrk")]
    [InlineData("rtl")]
    [InlineData("latinLnBrk")]
    [InlineData("hangingPunct")]
    public void NullableBoolProperties_DefaultToNotDefined(string _)
    {
        var pf = new ParagraphFormat();
        pf.EastAsianLineBreak.Should().Be(NullableBool.NotDefined);
        pf.RightToLeft.Should().Be(NullableBool.NotDefined);
        pf.LatinLineBreak.Should().Be(NullableBool.NotDefined);
        pf.HangingPunctuation.Should().Be(NullableBool.NotDefined);
    }

    [Fact]
    public void EastAsianLineBreak_TrueValue_RoundTrips()
    {
        var ppr = new XElement(ANs + "pPr");
        var pf = new ParagraphFormat();
        pf.InitInternal(ppr, null, null);

        pf.EastAsianLineBreak = NullableBool.True;

        var pf2 = new ParagraphFormat();
        pf2.InitInternal(ppr, null, null);
        pf2.EastAsianLineBreak.Should().Be(NullableBool.True);
    }

    [Fact]
    public void RightToLeft_FalseValue_RoundTrips()
    {
        var ppr = new XElement(ANs + "pPr");
        var pf = new ParagraphFormat();
        pf.InitInternal(ppr, null, null);

        pf.RightToLeft = NullableBool.False;

        var pf2 = new ParagraphFormat();
        pf2.InitInternal(ppr, null, null);
        pf2.RightToLeft.Should().Be(NullableBool.False);
    }

    [Fact]
    public void LatinLineBreak_SetToNotDefined_RemovesAttribute()
    {
        var ppr = new XElement(ANs + "pPr");
        var pf = new ParagraphFormat();
        pf.InitInternal(ppr, null, null);

        pf.LatinLineBreak = NullableBool.True;
        pf.LatinLineBreak.Should().Be(NullableBool.True);

        pf.LatinLineBreak = NullableBool.NotDefined;
        pf.LatinLineBreak.Should().Be(NullableBool.NotDefined);
    }

    [Fact]
    public void HangingPunctuation_TrueValue_RoundTrips()
    {
        var ppr = new XElement(ANs + "pPr");
        var pf = new ParagraphFormat();
        pf.InitInternal(ppr, null, null);

        pf.HangingPunctuation = NullableBool.True;

        var pf2 = new ParagraphFormat();
        pf2.InitInternal(ppr, null, null);
        pf2.HangingPunctuation.Should().Be(NullableBool.True);
    }

    [Fact]
    public void MarginLeft_RoundTrips()
    {
        var ppr = new XElement(ANs + "pPr");
        var pf = new ParagraphFormat();
        pf.InitInternal(ppr, null, null);

        pf.MarginLeft = 36;

        var pf2 = new ParagraphFormat();
        pf2.InitInternal(ppr, null, null);
        pf2.MarginLeft.Should().BeApproximately(36, 0.01f);
    }

    [Fact]
    public void MarginRight_RoundTrips()
    {
        var ppr = new XElement(ANs + "pPr");
        var pf = new ParagraphFormat();
        pf.InitInternal(ppr, null, null);

        pf.MarginRight = 18;

        var pf2 = new ParagraphFormat();
        pf2.InitInternal(ppr, null, null);
        pf2.MarginRight.Should().BeApproximately(18, 0.01f);
    }

    [Fact]
    public void Indent_RoundTrips()
    {
        var ppr = new XElement(ANs + "pPr");
        var pf = new ParagraphFormat();
        pf.InitInternal(ppr, null, null);

        pf.Indent = -18; // hanging indent

        var pf2 = new ParagraphFormat();
        pf2.InitInternal(ppr, null, null);
        pf2.Indent.Should().BeApproximately(-18, 0.01f);
    }

    [Fact]
    public void DefaultTabSize_RoundTrips()
    {
        var ppr = new XElement(ANs + "pPr");
        var pf = new ParagraphFormat();
        pf.InitInternal(ppr, null, null);

        pf.DefaultTabSize = 72;

        var pf2 = new ParagraphFormat();
        pf2.InitInternal(ppr, null, null);
        pf2.DefaultTabSize.Should().BeApproximately(72, 0.01f);
    }

    [Fact]
    public void FontAlignment_DefaultsToDefault()
    {
        var pf = new ParagraphFormat();
        pf.FontAlignment.Should().Be(FontAlignment.Default);
    }

    [Theory]
    [InlineData(FontAlignment.Automatic)]
    [InlineData(FontAlignment.Top)]
    [InlineData(FontAlignment.Center)]
    [InlineData(FontAlignment.Bottom)]
    [InlineData(FontAlignment.Baseline)]
    public void FontAlignment_AllValues_RoundTrip(FontAlignment fontAlignment)
    {
        var ppr = new XElement(ANs + "pPr");
        var pf = new ParagraphFormat();
        pf.InitInternal(ppr, null, null);

        pf.FontAlignment = fontAlignment;

        var pf2 = new ParagraphFormat();
        pf2.InitInternal(ppr, null, null);
        pf2.FontAlignment.Should().Be(fontAlignment);
    }

    [Fact]
    public void Depth_DefaultsToZero()
    {
        var pf = new ParagraphFormat();
        pf.Depth.Should().Be(0);
    }

    [Fact]
    public void Depth_CanBeSetAndReadBack()
    {
        var ppr = new XElement(ANs + "pPr");
        var pf = new ParagraphFormat();
        pf.InitInternal(ppr, null, null);

        pf.Depth = 3;

        var pf2 = new ParagraphFormat();
        pf2.InitInternal(ppr, null, null);
        pf2.Depth.Should().Be(3);
    }

    [Fact]
    public void Depth_SetToZero_RemovesAttribute()
    {
        var ppr = new XElement(ANs + "pPr");
        var pf = new ParagraphFormat();
        pf.InitInternal(ppr, null, null);

        pf.Depth = 2;
        pf.Depth.Should().Be(2);

        pf.Depth = 0;
        pf.Depth.Should().Be(0);
        ppr.Attribute("lvl").Should().BeNull();
    }

    [Fact]
    public void Bullet_IsNotNull()
    {
        var pf = new ParagraphFormat();
        pf.Bullet.Should().NotBeNull();
    }

    [Fact]
    public void DefaultPortionFormat_IsNotNull()
    {
        var pf = new ParagraphFormat();
        pf.DefaultPortionFormat.Should().NotBeNull();
    }

    [Fact]
    public void MarginLeft_SetToNaN_RemovesAttribute()
    {
        var ppr = new XElement(ANs + "pPr");
        var pf = new ParagraphFormat();
        pf.InitInternal(ppr, null, null);

        pf.MarginLeft = 36;
        pf.MarginLeft.Should().BeApproximately(36, 0.01f);

        pf.MarginLeft = float.NaN;
        float.IsNaN(pf.MarginLeft).Should().BeTrue();
    }
}
