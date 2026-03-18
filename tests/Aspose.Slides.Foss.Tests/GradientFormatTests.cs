using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests GradientFormat properties: GradientShape, GradientDirection,
/// LinearGradientAngle, LinearGradientScaled, TileFlip, and GradientStops.
/// </summary>
public sealed class GradientFormatTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// <summary>
    /// Creates a GradientFormat backed by a fresh &lt;a:gradFill&gt; element.
    /// </summary>
    private static GradientFormat CreateGradientFormat(XElement? gradFill = null)
    {
        gradFill ??= new XElement(ANs + "gradFill");
        var gf = new GradientFormat();
        gf.InitInternal(gradFill, slidePart: null, parentSlide: null);
        return gf;
    }

    // --- GradientShape ---

    [Fact]
    public void GradientShape_NoLinOrPath_ReturnsNotDefined()
    {
        var gf = CreateGradientFormat();
        gf.GradientShape.Should().Be(GradientShape.NotDefined);
    }

    [Fact]
    public void GradientShape_WithLinElement_ReturnsLinear()
    {
        var gradFill = new XElement(ANs + "gradFill",
            new XElement(ANs + "lin", new XAttribute("ang", "0")));
        var gf = CreateGradientFormat(gradFill);
        gf.GradientShape.Should().Be(GradientShape.Linear);
    }

    [Theory]
    [InlineData("rect", GradientShape.Rectangle)]
    [InlineData("circle", GradientShape.Radial)]
    [InlineData("shape", GradientShape.Path)]
    public void GradientShape_WithPathElement_ReturnsExpected(string pathVal, GradientShape expected)
    {
        var gradFill = new XElement(ANs + "gradFill",
            new XElement(ANs + "path", new XAttribute("path", pathVal)));
        var gf = CreateGradientFormat(gradFill);
        gf.GradientShape.Should().Be(expected);
    }

    [Fact]
    public void GradientShape_SetLinear_CreatesLinElement()
    {
        var gf = CreateGradientFormat();
        gf.GradientShape = GradientShape.Linear;

        gf.GradientShape.Should().Be(GradientShape.Linear);
    }

    [Theory]
    [InlineData(GradientShape.Rectangle)]
    [InlineData(GradientShape.Radial)]
    [InlineData(GradientShape.Path)]
    public void GradientShape_SetNonLinear_CreatesPathElement(GradientShape shape)
    {
        var gf = CreateGradientFormat();
        gf.GradientShape = shape;

        gf.GradientShape.Should().Be(shape);
    }

    [Fact]
    public void GradientShape_SetLinear_RemovesExistingPath()
    {
        var gradFill = new XElement(ANs + "gradFill",
            new XElement(ANs + "path", new XAttribute("path", "circle")));
        var gf = CreateGradientFormat(gradFill);

        gf.GradientShape = GradientShape.Linear;

        gf.GradientShape.Should().Be(GradientShape.Linear);
        gradFill.Element(ANs + "path").Should().BeNull();
    }

    [Fact]
    public void GradientShape_SetRadial_RemovesExistingLin()
    {
        var gradFill = new XElement(ANs + "gradFill",
            new XElement(ANs + "lin", new XAttribute("ang", "0")));
        var gf = CreateGradientFormat(gradFill);

        gf.GradientShape = GradientShape.Radial;

        gf.GradientShape.Should().Be(GradientShape.Radial);
        gradFill.Element(ANs + "lin").Should().BeNull();
    }

    [Fact]
    public void GradientShape_SetNotDefined_DoesNothing()
    {
        var gradFill = new XElement(ANs + "gradFill",
            new XElement(ANs + "lin", new XAttribute("ang", "0")));
        var gf = CreateGradientFormat(gradFill);

        gf.GradientShape = GradientShape.NotDefined;

        gf.GradientShape.Should().Be(GradientShape.Linear);
    }

    // --- LinearGradientAngle ---

    [Fact]
    public void LinearGradientAngle_NoLin_ReturnsZero()
    {
        var gf = CreateGradientFormat();
        gf.LinearGradientAngle.Should().Be(0f);
    }

    [Fact]
    public void LinearGradientAngle_SetAndRead_RoundTrips()
    {
        var gf = CreateGradientFormat();
        gf.LinearGradientAngle = 45;

        gf.LinearGradientAngle.Should().Be(45f);
    }

    [Fact]
    public void LinearGradientAngle_SetOnExistingLin_UpdatesValue()
    {
        var gradFill = new XElement(ANs + "gradFill",
            new XElement(ANs + "lin", new XAttribute("ang", "0")));
        var gf = CreateGradientFormat(gradFill);

        gf.LinearGradientAngle = 90;

        gf.LinearGradientAngle.Should().Be(90f);
    }

    [Fact]
    public void LinearGradientAngle_ReadFromXml_ConvertsCorrectly()
    {
        var gradFill = new XElement(ANs + "gradFill",
            new XElement(ANs + "lin", new XAttribute("ang", "2700000")));
        var gf = CreateGradientFormat(gradFill);

        gf.LinearGradientAngle.Should().Be(45f);
    }

    // --- LinearGradientScaled ---

    [Fact]
    public void LinearGradientScaled_NoLin_ReturnsNotDefined()
    {
        var gf = CreateGradientFormat();
        gf.LinearGradientScaled.Should().Be(NullableBool.NotDefined);
    }

    [Fact]
    public void LinearGradientScaled_SetTrue_ReadsBack()
    {
        var gf = CreateGradientFormat();
        gf.LinearGradientScaled = NullableBool.True;

        gf.LinearGradientScaled.Should().Be(NullableBool.True);
    }

    [Fact]
    public void LinearGradientScaled_SetFalse_ReadsBack()
    {
        var gradFill = new XElement(ANs + "gradFill",
            new XElement(ANs + "lin", new XAttribute("ang", "0")));
        var gf = CreateGradientFormat(gradFill);

        gf.LinearGradientScaled = NullableBool.False;

        gf.LinearGradientScaled.Should().Be(NullableBool.False);
    }

    [Fact]
    public void LinearGradientScaled_SetNotDefined_RemovesAttribute()
    {
        var gradFill = new XElement(ANs + "gradFill",
            new XElement(ANs + "lin",
                new XAttribute("ang", "0"),
                new XAttribute("scaled", "1")));
        var gf = CreateGradientFormat(gradFill);

        gf.LinearGradientScaled = NullableBool.NotDefined;

        gf.LinearGradientScaled.Should().Be(NullableBool.NotDefined);
    }

    // --- TileFlip ---

    [Fact]
    public void TileFlip_NoAttribute_ReturnsNotDefined()
    {
        var gf = CreateGradientFormat();
        gf.TileFlip.Should().Be(TileFlip.NotDefined);
    }

    [Theory]
    [InlineData("x", TileFlip.FlipX)]
    [InlineData("y", TileFlip.FlipY)]
    [InlineData("xy", TileFlip.FlipBoth)]
    [InlineData("none", TileFlip.NoFlip)]
    public void TileFlip_XmlValue_MapsCorrectly(string xmlVal, TileFlip expected)
    {
        var gradFill = new XElement(ANs + "gradFill", new XAttribute("flip", xmlVal));
        var gf = CreateGradientFormat(gradFill);
        gf.TileFlip.Should().Be(expected);
    }

    [Fact]
    public void TileFlip_SetAndRead_RoundTrips()
    {
        var gf = CreateGradientFormat();
        gf.TileFlip = TileFlip.FlipBoth;

        gf.TileFlip.Should().Be(TileFlip.FlipBoth);
    }

    [Fact]
    public void TileFlip_SetNotDefined_RemovesAttribute()
    {
        var gradFill = new XElement(ANs + "gradFill", new XAttribute("flip", "x"));
        var gf = CreateGradientFormat(gradFill);

        gf.TileFlip = TileFlip.NotDefined;

        gf.TileFlip.Should().Be(TileFlip.NotDefined);
    }

    // --- GradientDirection ---

    [Fact]
    public void GradientDirection_NoLinOrPath_ReturnsNotDefined()
    {
        var gf = CreateGradientFormat();
        gf.GradientDirection.Should().Be(GradientDirection.NotDefined);
    }

    [Fact]
    public void GradientDirection_WithPath_ReturnsFromCenter()
    {
        var gradFill = new XElement(ANs + "gradFill",
            new XElement(ANs + "path", new XAttribute("path", "circle")));
        var gf = CreateGradientFormat(gradFill);
        gf.GradientDirection.Should().Be(GradientDirection.FromCenter);
    }

    [Theory]
    [InlineData("0", GradientDirection.FromCorner1)]
    [InlineData("5400000", GradientDirection.FromCorner2)]
    [InlineData("10800000", GradientDirection.FromCorner4)]
    [InlineData("16200000", GradientDirection.FromCorner3)]
    public void GradientDirection_LinAngle_MapsToDirection(string ang, GradientDirection expected)
    {
        var gradFill = new XElement(ANs + "gradFill",
            new XElement(ANs + "lin", new XAttribute("ang", ang)));
        var gf = CreateGradientFormat(gradFill);
        gf.GradientDirection.Should().Be(expected);
    }

    [Fact]
    public void GradientDirection_SetFromCenter_CreatesPathElement()
    {
        var gf = CreateGradientFormat();
        gf.GradientDirection = GradientDirection.FromCenter;

        gf.GradientDirection.Should().Be(GradientDirection.FromCenter);
    }

    [Fact]
    public void GradientDirection_SetCorner_CreatesLinElement()
    {
        var gf = CreateGradientFormat();
        gf.GradientDirection = GradientDirection.FromCorner1;

        gf.GradientDirection.Should().Be(GradientDirection.FromCorner1);
    }

    [Fact]
    public void GradientDirection_SetNotDefined_DoesNothing()
    {
        var gradFill = new XElement(ANs + "gradFill",
            new XElement(ANs + "lin", new XAttribute("ang", "0")));
        var gf = CreateGradientFormat(gradFill);

        gf.GradientDirection = GradientDirection.NotDefined;

        gf.GradientDirection.Should().Be(GradientDirection.FromCorner1);
    }

    [Fact]
    public void GradientDirection_SetFromCenter_RemovesExistingLin()
    {
        var gradFill = new XElement(ANs + "gradFill",
            new XElement(ANs + "lin", new XAttribute("ang", "0")));
        var gf = CreateGradientFormat(gradFill);

        gf.GradientDirection = GradientDirection.FromCenter;

        gradFill.Element(ANs + "lin").Should().BeNull();
        gf.GradientDirection.Should().Be(GradientDirection.FromCenter);
    }

    [Fact]
    public void GradientDirection_SetCorner_RemovesExistingPath()
    {
        var gradFill = new XElement(ANs + "gradFill",
            new XElement(ANs + "path", new XAttribute("path", "circle")));
        var gf = CreateGradientFormat(gradFill);

        gf.GradientDirection = GradientDirection.FromCorner2;

        gradFill.Element(ANs + "path").Should().BeNull();
        gf.GradientDirection.Should().Be(GradientDirection.FromCorner2);
    }

    // --- GradientStops ---

    [Fact]
    public void GradientStops_ReturnsCollection()
    {
        var gf = CreateGradientFormat();
        gf.GradientStops.Should().NotBeNull();
        gf.GradientStops.Should().BeAssignableTo<IGradientStopCollection>();
    }

    [Fact]
    public void GradientStops_CreatesGsLstIfMissing()
    {
        var gradFill = new XElement(ANs + "gradFill");
        var gf = CreateGradientFormat(gradFill);

        _ = gf.GradientStops;

        gradFill.Element(ANs + "gsLst").Should().NotBeNull();
    }

    [Fact]
    public void GradientStops_ExistingGsLst_IsPreserved()
    {
        var gradFill = new XElement(ANs + "gradFill",
            new XElement(ANs + "gsLst",
                new XElement(ANs + "gs", new XAttribute("pos", "0"),
                    new XElement(ANs + "srgbClr", new XAttribute("val", "0000FF")))));
        var gf = CreateGradientFormat(gradFill);

        gf.GradientStops.Count.Should().Be(1);
    }


    [Fact]
    public void GradientFill_ShapeAngleAndStops_Persist()
    {
        var parent = new XElement(ANs + "spPr");
        var ff = new FillFormat();
        ff.InitInternal(parent, parentSlide: null);

        ff.FillType = FillType.Gradient;
        var gf = ff.GradientFormat;
        gf.GradientShape = GradientShape.Linear;
        gf.LinearGradientAngle = 45;

        ff.FillType.Should().Be(FillType.Gradient);
        ff.GradientFormat.GradientShape.Should().Be(GradientShape.Linear);
        ff.GradientFormat.LinearGradientAngle.Should().Be(45f);
    }

    [Fact]
    public void GradientFill_StopsCanBeAddedViaFillFormat()
    {
        var parent = new XElement(ANs + "spPr");
        var ff = new FillFormat();
        ff.InitInternal(parent, parentSlide: null);

        ff.FillType = FillType.Gradient;
        ff.GradientFormat.GradientStops.Add(0.0f, Color.Blue);
        ff.GradientFormat.GradientStops.Add(1.0f, Color.Red);

        ff.GradientFormat.GradientStops.Count.Should().BeGreaterThanOrEqualTo(2);
    }
}
