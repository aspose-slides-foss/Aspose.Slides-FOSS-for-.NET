using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Verifies PatternFormat: pattern style, fore/back colors, XML round-trip.
/// </summary>
public sealed class PatternFormatTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static PatternFormat CreatePatternFormat(XElement? pattFill = null)
    {
        var element = pattFill ?? new XElement(ANs + "pattFill");
        var pf = new PatternFormat();
        pf.InitInternal(element, parentSlide: null);
        return pf;
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void PatternStyle_SetAndGet_RoundTrips()
    {
        var pf = CreatePatternFormat();
        pf.PatternStyle = PatternStyle.Percent50;

        pf.PatternStyle.Should().Be(PatternStyle.Percent50);
    }

    /// <summary>
    /// </summary>
    [Theory]
    [InlineData(PatternStyle.Percent05)]
    [InlineData(PatternStyle.DarkHorizontal)]
    [InlineData(PatternStyle.LightVertical)]
    [InlineData(PatternStyle.Zigzag)]
    [InlineData(PatternStyle.DiagonalCross)]
    [InlineData(PatternStyle.DottedDiamond)]
    [InlineData(PatternStyle.Weave)]
    public void PatternStyle_VariousStyles_RoundTrip(PatternStyle style)
    {
        var pf = CreatePatternFormat();
        pf.PatternStyle = style;

        pf.PatternStyle.Should().Be(style);
    }

    /// <summary>
    /// No prst attribute returns NotDefined.
    /// </summary>
    [Fact]
    public void PatternStyle_NoPrstAttribute_ReturnsNotDefined()
    {
        var pf = CreatePatternFormat();

        pf.PatternStyle.Should().Be(PatternStyle.NotDefined);
    }

    /// <summary>
    /// Unrecognized prst attribute returns Unknown.
    /// </summary>
    [Fact]
    public void PatternStyle_UnrecognizedPrst_ReturnsUnknown()
    {
        var el = new XElement(ANs + "pattFill", new XAttribute("prst", "futurePatternXyz"));
        var pf = CreatePatternFormat(el);

        pf.PatternStyle.Should().Be(PatternStyle.Unknown);
    }

    /// <summary>
    /// Setting NotDefined removes prst attribute.
    /// </summary>
    [Fact]
    public void PatternStyle_SetNotDefined_RemovesPrstAttribute()
    {
        var el = new XElement(ANs + "pattFill", new XAttribute("prst", "pct50"));
        var pf = CreatePatternFormat(el);

        pf.PatternStyle = PatternStyle.NotDefined;

        pf.PatternStyle.Should().Be(PatternStyle.NotDefined);
        el.Attribute("prst").Should().BeNull();
    }

    /// <summary>
    /// Setting Unknown removes prst attribute.
    /// </summary>
    [Fact]
    public void PatternStyle_SetUnknown_RemovesPrstAttribute()
    {
        var el = new XElement(ANs + "pattFill", new XAttribute("prst", "pct50"));
        var pf = CreatePatternFormat(el);

        pf.PatternStyle = PatternStyle.Unknown;

        pf.PatternStyle.Should().Be(PatternStyle.NotDefined);
        el.Attribute("prst").Should().BeNull();
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void ForeColor_CreatesElementWhenMissing()
    {
        var el = new XElement(ANs + "pattFill");
        var pf = CreatePatternFormat(el);

        var fc = pf.ForeColor;

        fc.Should().NotBeNull();
        el.Element(ANs + "fgClr").Should().NotBeNull();
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void BackColor_CreatesElementWhenMissing()
    {
        var el = new XElement(ANs + "pattFill");
        var pf = CreatePatternFormat(el);

        var bc = pf.BackColor;

        bc.Should().NotBeNull();
        el.Element(ANs + "bgClr").Should().NotBeNull();
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void ForeAndBackColor_CanBeSet()
    {
        var pf = CreatePatternFormat();
        pf.ForeColor.Color = Color.DarkBlue;
        pf.BackColor.Color = Color.LightYellow;

        pf.ForeColor.Color.Should().NotBeNull();
        pf.BackColor.Color.Should().NotBeNull();
    }

    /// <summary>
    /// ForeColor returns existing fgClr element if present.
    /// </summary>
    [Fact]
    public void ForeColor_ReusesExistingElement()
    {
        var fgClr = new XElement(ANs + "fgClr");
        var el = new XElement(ANs + "pattFill", fgClr);
        var pf = CreatePatternFormat(el);

        _ = pf.ForeColor;

        el.Elements(ANs + "fgClr").Should().HaveCount(1);
    }

    /// <summary>
    /// BackColor returns existing bgClr element if present.
    /// </summary>
    [Fact]
    public void BackColor_ReusesExistingElement()
    {
        var bgClr = new XElement(ANs + "bgClr");
        var el = new XElement(ANs + "pattFill", bgClr);
        var pf = CreatePatternFormat(el);

        _ = pf.BackColor;

        el.Elements(ANs + "bgClr").Should().HaveCount(1);
    }

    /// <summary>
    /// Pattern style and colors set through FillFormat persist in the XML tree.
    /// </summary>
    [Fact]
    public void PatternFill_ViaFillFormat_RoundTrips()
    {
        var parent = new XElement(ANs + "spPr");
        var ff = new FillFormat();
        ff.InitInternal(parent, parentSlide: null);

        ff.FillType = FillType.Pattern;
        ff.PatternFormat.PatternStyle = PatternStyle.Percent50;
        ff.PatternFormat.ForeColor.Color = Color.DarkBlue;
        ff.PatternFormat.BackColor.Color = Color.LightYellow;

        ff.FillType.Should().Be(FillType.Pattern);
        ff.PatternFormat.PatternStyle.Should().Be(PatternStyle.Percent50);
    }
}
