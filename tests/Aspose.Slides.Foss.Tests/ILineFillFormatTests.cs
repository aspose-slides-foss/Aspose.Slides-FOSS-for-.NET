using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Verifies ILineFillFormat contract through LineFillFormat backed by XML.
/// </summary>
public sealed class ILineFillFormatTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static LineFillFormat CreateLineFillFormat(XElement? lnElement = null)
    {
        var el = lnElement ?? new XElement(ANs + "ln");
        var lff = new LineFillFormat();
        lff.InitInternal(el, null, null);
        return lff;
    }

    /// <summary>
    /// Setting FillType to Solid creates solidFill child element and reads back correctly.
    /// </summary>
    [Fact]
    public void FillType_SolidCanBeSetAndRead()
    {
        var lff = CreateLineFillFormat();
        lff.FillType = FillType.Solid;

        lff.FillType.Should().Be(FillType.Solid);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void SolidFillColor_ColorComponentsPersist()
    {
        var lff = CreateLineFillFormat();
        lff.FillType = FillType.Solid;
        lff.SolidFillColor.Color = Color.FromArgb(255, 0, 128, 255);

        var c = lff.SolidFillColor.Color;
        c.Should().NotBeNull();
        c!.R.Should().Be(0);
        c.G.Should().Be(128);
        c.B.Should().Be(255);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void FillType_GradientCanBeSetAndRead()
    {
        var lff = CreateLineFillFormat();
        lff.FillType = FillType.Gradient;

        lff.FillType.Should().Be(FillType.Gradient);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void GradientFormat_StopsCanBeAddedAndCounted()
    {
        var ln = new XElement(ANs + "ln",
            new XElement(ANs + "gradFill",
                new XElement(ANs + "gsLst")));
        var lff = CreateLineFillFormat(ln);

        lff.FillType.Should().Be(FillType.Gradient);
        lff.GradientFormat.GradientStops.Add(0.0f, Color.Blue);
        lff.GradientFormat.GradientStops.Add(1.0f, Color.Red);

        lff.GradientFormat.GradientStops.Count.Should().BeGreaterThanOrEqualTo(2);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void FillType_PatternCanBeSetAndRead()
    {
        var lff = CreateLineFillFormat();
        lff.FillType = FillType.Pattern;

        lff.FillType.Should().Be(FillType.Pattern);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void PatternFormat_PatternStyleCanBeSetAndRead()
    {
        var lff = CreateLineFillFormat();
        lff.FillType = FillType.Pattern;
        lff.PatternFormat.PatternStyle = PatternStyle.Percent50;

        lff.PatternFormat.PatternStyle.Should().Be(PatternStyle.Percent50);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void PatternFormat_ForeAndBackColorsCanBeSet()
    {
        var lff = CreateLineFillFormat();
        lff.FillType = FillType.Pattern;
        lff.PatternFormat.ForeColor.Color = Color.DarkBlue;
        lff.PatternFormat.BackColor.Color = Color.LightYellow;

        lff.PatternFormat.ForeColor.Color.Should().NotBeNull();
        lff.PatternFormat.BackColor.Color.Should().NotBeNull();
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void FillType_NoFillCanBeSetAndRead()
    {
        var lff = CreateLineFillFormat();
        lff.FillType = FillType.NoFill;

        lff.FillType.Should().Be(FillType.NoFill);
    }

    /// <summary>
    /// LineFillFormat supports solid fill with a specific color (DarkRed).
    /// </summary>
    [Fact]
    public void SolidFillColor_DarkRedColorPersists()
    {
        var lff = CreateLineFillFormat();
        lff.FillType = FillType.Solid;
        lff.SolidFillColor.Color = Color.DarkRed;

        var c = lff.SolidFillColor.Color;
        c.Should().NotBeNull();
        c!.R.Should().Be(Color.DarkRed.R);
    }

    /// <summary>
    /// LineFillFormat solid fill with black color persists.
    /// </summary>
    [Fact]
    public void SolidFillColor_BlackColorPersists()
    {
        var lff = CreateLineFillFormat();
        lff.FillType = FillType.Solid;
        lff.SolidFillColor.Color = Color.Black;

        var c = lff.SolidFillColor.Color;
        c.Should().NotBeNull();
        c!.R.Should().Be(0);
        c.G.Should().Be(0);
        c.B.Should().Be(0);
    }

    /// <summary>
    /// LineFillFormat solid fill with red color on border persists.
    /// </summary>
    [Fact]
    public void SolidFillColor_RedColorPersists()
    {
        var lff = CreateLineFillFormat();
        lff.FillType = FillType.Solid;
        lff.SolidFillColor.Color = Color.Red;

        lff.FillType.Should().Be(FillType.Solid);
        var c = lff.SolidFillColor.Color;
        c.Should().NotBeNull();
        c!.R.Should().Be(255);
        c.G.Should().Be(0);
        c.B.Should().Be(0);
    }

    /// <summary>
    /// Solid fill colour on portion text persists — verifies FillType.Solid round-trip.
    /// </summary>
    [Fact]
    public void FillType_ChangingTypeReplacesPreviousFill()
    {
        var lff = CreateLineFillFormat();
        lff.FillType = FillType.Solid;
        lff.FillType.Should().Be(FillType.Solid);

        lff.FillType = FillType.Gradient;
        lff.FillType.Should().Be(FillType.Gradient);
    }

    /// <summary>
    /// Verifies that RotateWithShape can be set and read on a fill element.
    /// </summary>
    [Fact]
    public void RotateWithShape_CanBeSetToTrue()
    {
        var lff = CreateLineFillFormat();
        lff.FillType = FillType.Solid;
        lff.RotateWithShape = NullableBool.True;

        lff.RotateWithShape.Should().Be(NullableBool.True);
    }

    /// <summary>
    /// Verifies that RotateWithShape can be set to False.
    /// </summary>
    [Fact]
    public void RotateWithShape_CanBeSetToFalse()
    {
        var lff = CreateLineFillFormat();
        lff.FillType = FillType.Solid;
        lff.RotateWithShape = NullableBool.False;

        lff.RotateWithShape.Should().Be(NullableBool.False);
    }

    /// <summary>
    /// Verifies that RotateWithShape returns NotDefined when no fill element exists.
    /// </summary>
    [Fact]
    public void RotateWithShape_ReturnsNotDefinedWhenNoFill()
    {
        var lff = CreateLineFillFormat();

        lff.RotateWithShape.Should().Be(NullableBool.NotDefined);
    }

    /// <summary>
    /// Verifies that FillType returns NotDefined when no fill child exists.
    /// </summary>
    [Fact]
    public void FillType_ReturnsNotDefinedWhenEmpty()
    {
        var lff = CreateLineFillFormat();

        lff.FillType.Should().Be(FillType.NotDefined);
    }

    /// <summary>
    /// SolidFillColor accessor auto-creates solidFill element.
    /// </summary>
    [Fact]
    public void SolidFillColor_AutoCreatesSolidFillElement()
    {
        var lff = CreateLineFillFormat();

        lff.SolidFillColor.Color = Color.Red;

        lff.FillType.Should().Be(FillType.Solid);
    }

    /// <summary>
    /// GradientFormat accessor auto-creates gradFill element.
    /// </summary>
    [Fact]
    public void GradientFormat_AutoCreatesGradFillElement()
    {
        var lff = CreateLineFillFormat();

        var gf = lff.GradientFormat;

        lff.FillType.Should().Be(FillType.Gradient);
    }

    /// <summary>
    /// PatternFormat accessor auto-creates pattFill element.
    /// </summary>
    [Fact]
    public void PatternFormat_AutoCreatesPattFillElement()
    {
        var lff = CreateLineFillFormat();

        var pf = lff.PatternFormat;

        lff.FillType.Should().Be(FillType.Pattern);
    }

    /// <summary>
    /// Fill element is inserted before dash/end elements in ln.
    /// </summary>
    [Fact]
    public void FillType_InsertedBeforeDashElements()
    {
        var ln = new XElement(ANs + "ln",
            new XElement(ANs + "prstDash", new XAttribute("val", "dash")));
        var lff = CreateLineFillFormat(ln);

        lff.FillType = FillType.Solid;

        lff.FillType.Should().Be(FillType.Solid);
        var firstChild = ln.Elements().First();
        firstChild.Name.LocalName.Should().Be("solidFill");
    }

    /// <summary>
    /// Setting FillType resets the lnRef idx to 0 in the parent shape's style
    /// so the explicit line fill takes priority over the theme reference.
    /// </summary>
    [Fact]
    public void FillType_ResetsStyleLnRefIdx()
    {
        var pNs = XNamespace.Get("http://schemas.openxmlformats.org/presentationml/2006/main");
        var ln = new XElement(ANs + "ln");
        var sp = new XElement(pNs + "sp",
            new XElement(pNs + "spPr", ln),
            new XElement(pNs + "style",
                new XElement(ANs + "lnRef", new XAttribute("idx", "2"))));
        var lff = CreateLineFillFormat(ln);

        lff.FillType = FillType.Solid;

        var lnRef = sp.Element(pNs + "style")!.Element(ANs + "lnRef")!;
        lnRef.Attribute("idx")!.Value.Should().Be("0");
    }

    /// <summary>
    /// Setting FillType to NotDefined does not reset lnRef (no fill element inserted).
    /// </summary>
    [Fact]
    public void FillType_DoesNotResetLnRefWhenNotDefined()
    {
        var pNs = XNamespace.Get("http://schemas.openxmlformats.org/presentationml/2006/main");
        var ln = new XElement(ANs + "ln",
            new XElement(ANs + "solidFill"));
        var sp = new XElement(pNs + "sp",
            new XElement(pNs + "spPr", ln),
            new XElement(pNs + "style",
                new XElement(ANs + "lnRef", new XAttribute("idx", "2"))));
        var lff = CreateLineFillFormat(ln);

        lff.FillType = FillType.NotDefined;

        var lnRef = sp.Element(pNs + "style")!.Element(ANs + "lnRef")!;
        lnRef.Attribute("idx")!.Value.Should().Be("2");
    }
}
