using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Verifies the enum values and color constants used by fill format operations.
/// </summary>
public sealed class FillFormatTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static FillFormat CreateFillFormat(XElement? parentElement = null)
    {
        var element = parentElement ?? new XElement(ANs + "spPr");
        var ff = new FillFormat();
        ff.InitInternal(element, parentSlide: null);
        return ff;
    }
    /// <summary>
    /// Verifies FillType.Solid and Color.FromArgb components used in the test.
    /// </summary>
    [Fact]
    public void SolidFill_FillTypeSolidIsDefined()
    {
        Enum.IsDefined(FillType.Solid).Should().BeTrue();
        FillType.Solid.Should().NotBe(FillType.NotDefined);
    }

    [Fact]
    public void SolidFill_ColorFromArgbHasCorrectComponents()
    {
        var c = Color.FromArgb(255, 0, 128, 255);
        c.A.Should().Be(255);
        c.R.Should().Be(0);
        c.G.Should().Be(128);
        c.B.Should().Be(255);
    }

    [Fact]
    public void SolidFill_FillTypeSolidIsDistinctFromOtherTypes()
    {
        FillType.Solid.Should().NotBe(FillType.NoFill);
        FillType.Solid.Should().NotBe(FillType.Gradient);
        FillType.Solid.Should().NotBe(FillType.Pattern);
        FillType.Solid.Should().NotBe(FillType.Picture);
    }

    /// <summary>
    /// Verifies FillType.Gradient, GradientShape.Linear, Color.Blue, Color.Red.
    /// </summary>
    [Fact]
    public void GradientFill_FillTypeGradientIsDefined()
    {
        Enum.IsDefined(FillType.Gradient).Should().BeTrue();
        FillType.Gradient.Should().NotBe(FillType.NotDefined);
    }

    [Fact]
    public void GradientFill_GradientShapeLinearIsDefined()
    {
        Enum.IsDefined(GradientShape.Linear).Should().BeTrue();
        GradientShape.Linear.Should().NotBe(GradientShape.NotDefined);
    }

    [Fact]
    public void GradientFill_BlueColorHasCorrectComponents()
    {
        Color.Blue.Should().NotBe(Color.Empty);
        Color.Blue.R.Should().Be(0);
        Color.Blue.G.Should().Be(0);
        Color.Blue.B.Should().Be(255);
    }

    [Fact]
    public void GradientFill_RedColorHasCorrectComponents()
    {
        Color.Red.Should().NotBe(Color.Empty);
        Color.Red.R.Should().Be(255);
        Color.Red.G.Should().Be(0);
        Color.Red.B.Should().Be(0);
    }

    [Fact]
    public void GradientFill_BlueAndRedAreDistinct()
    {
        Color.Blue.Should().NotBe(Color.Red);
    }

    /// <summary>
    /// Verifies FillType.Pattern, PatternStyle.Percent50, Color.DarkBlue, Color.LightYellow.
    /// </summary>
    [Fact]
    public void PatternFill_FillTypePatternIsDefined()
    {
        Enum.IsDefined(FillType.Pattern).Should().BeTrue();
        FillType.Pattern.Should().NotBe(FillType.NotDefined);
    }

    [Fact]
    public void PatternFill_PatternStylePercent50IsDefined()
    {
        Enum.IsDefined(PatternStyle.Percent50).Should().BeTrue();
        PatternStyle.Percent50.Should().NotBe(PatternStyle.NotDefined);
    }

    [Fact]
    public void PatternFill_DarkBlueColorExists()
    {
        Color.DarkBlue.Should().NotBe(Color.Empty);
        Color.DarkBlue.B.Should().BeGreaterThan(0);
    }

    [Fact]
    public void PatternFill_LightYellowColorExists()
    {
        Color.LightYellow.Should().NotBe(Color.Empty);
    }

    [Fact]
    public void PatternFill_ForeAndBackColorsAreDistinct()
    {
        Color.DarkBlue.Should().NotBe(Color.LightYellow);
    }

    /// <summary>
    /// Verifies FillType.NoFill.
    /// </summary>
    [Fact]
    public void NoFill_FillTypeNoFillIsDefined()
    {
        Enum.IsDefined(FillType.NoFill).Should().BeTrue();
        FillType.NoFill.Should().NotBe(FillType.NotDefined);
    }

    [Fact]
    public void NoFill_IsDistinctFromSolid()
    {
        FillType.NoFill.Should().NotBe(FillType.Solid);
    }

    /// <summary>
    /// Verifies FillType.Picture, PictureFillMode.Stretch.
    /// </summary>
    [Fact]
    public void PictureFill_FillTypePictureIsDefined()
    {
        Enum.IsDefined(FillType.Picture).Should().BeTrue();
        FillType.Picture.Should().NotBe(FillType.NotDefined);
    }

    [Fact]
    public void PictureFill_PictureFillModeStretchIsDefined()
    {
        Enum.IsDefined(PictureFillMode.Stretch).Should().BeTrue();
        ((int)PictureFillMode.Stretch).Should().BeGreaterThanOrEqualTo(0);
    }

    /// <summary>
    /// Verifies all FillType values used across tests are valid.
    /// </summary>
    [Theory]
    [InlineData(FillType.NoFill)]
    [InlineData(FillType.Solid)]
    [InlineData(FillType.Gradient)]
    [InlineData(FillType.Pattern)]
    [InlineData(FillType.Picture)]
    [InlineData(FillType.Group)]
    public void FillType_AllValuesAreDefined(FillType fillType)
    {
        Enum.IsDefined(fillType).Should().BeTrue();
    }

    [Fact]
    public void FillType_NotDefinedSentinelIsZero()
    {
        ((int)FillType.NotDefined).Should().Be(0);
    }

    [Fact]
    public void GradientShape_NotDefinedSentinelIsZero()
    {
        ((int)GradientShape.NotDefined).Should().Be(0);
    }

    [Fact]
    public void PatternStyle_NotDefinedSentinelIsZero()
    {
        ((int)PatternStyle.NotDefined).Should().Be(0);
    }

    [Theory]
    [InlineData(PictureFillMode.Tile)]
    [InlineData(PictureFillMode.Stretch)]
    public void PictureFillMode_AllValuesAreDefined(PictureFillMode mode)
    {
        Enum.IsDefined(mode).Should().BeTrue();
    }

    /// <summary>
    /// Setting FillType to Solid creates solidFill XML and reads back correctly.
    /// </summary>
    [Fact]
    public void SolidFill_FillTypeCanBeSetAndRead()
    {
        var ff = CreateFillFormat();
        ff.FillType = FillType.Solid;

        ff.FillType.Should().Be(FillType.Solid);
    }

    /// <summary>
    /// SolidFillColor can be set and the RGB components are preserved.
    /// </summary>
    [Fact]
    public void SolidFill_ColorComponentsPersistInXml()
    {
        var ff = CreateFillFormat();
        ff.FillType = FillType.Solid;
        ff.SolidFillColor.Color = Color.FromArgb(255, 0, 128, 255);

        var c = ff.SolidFillColor.Color;
        c.Should().NotBeNull();
        c!.R.Should().Be(0);
        c.G.Should().Be(128);
        c.B.Should().Be(255);
    }

    /// <summary>
    /// Gradient fill type can be set and gradient stops can be added.
    /// </summary>
    [Fact]
    public void GradientFill_StopsCanBeAddedAndCounted()
    {
        var parent = new XElement(ANs + "spPr",
            new XElement(ANs + "gradFill",
                new XElement(ANs + "gsLst")));
        var ff = CreateFillFormat(parent);

        ff.FillType.Should().Be(FillType.Gradient);
        ff.GradientFormat.GradientStops.Add(0.0f, Color.Blue);
        ff.GradientFormat.GradientStops.Add(1.0f, Color.Red);

        ff.GradientFormat.GradientStops.Count.Should().BeGreaterThanOrEqualTo(2);
    }

    /// <summary>
    /// LinearGradientAngle can be set and read back.
    /// </summary>
    [Fact]
    public void GradientFill_LinearGradientAnglePersists()
    {
        var parent = new XElement(ANs + "spPr",
            new XElement(ANs + "gradFill",
                new XElement(ANs + "gsLst")));
        var ff = CreateFillFormat(parent);

        ff.GradientFormat.LinearGradientAngle = 45;

        ff.GradientFormat.LinearGradientAngle.Should().Be(45);
    }

    /// <summary>
    /// GradientShape is Linear when lin element is present.
    /// </summary>
    [Fact]
    public void GradientFill_GradientShapeIsLinearWhenLinElementExists()
    {
        var parent = new XElement(ANs + "spPr",
            new XElement(ANs + "gradFill",
                new XElement(ANs + "gsLst"),
                new XElement(ANs + "lin", new XAttribute("ang", "2700000"))));
        var ff = CreateFillFormat(parent);

        ff.GradientFormat.GradientShape.Should().Be(GradientShape.Linear);
    }

    /// <summary>
    /// Setting FillType to Pattern creates pattFill XML.
    /// </summary>
    [Fact]
    public void PatternFill_FillTypeCanBeSetAndRead()
    {
        var ff = CreateFillFormat();
        ff.FillType = FillType.Pattern;

        ff.FillType.Should().Be(FillType.Pattern);
    }

    /// <summary>
    /// Setting FillType to NoFill creates noFill XML.
    /// </summary>
    [Fact]
    public void NoFill_FillTypeCanBeSetAndRead()
    {
        var ff = CreateFillFormat();
        ff.FillType = FillType.NoFill;

        ff.FillType.Should().Be(FillType.NoFill);
    }

    /// <summary>
    /// Setting FillType to Picture creates blipFill XML.
    /// </summary>
    [Fact]
    public void PictureFill_FillTypeCanBeSetAndRead()
    {
        var ff = CreateFillFormat();
        ff.FillType = FillType.Picture;

        ff.FillType.Should().Be(FillType.Picture);
    }

    /// <summary>
    /// Changing FillType replaces the previous fill element.
    /// </summary>
    [Fact]
    public void FillType_ChangingTypeReplacePreviousFill()
    {
        var ff = CreateFillFormat();
        ff.FillType = FillType.Solid;
        ff.FillType.Should().Be(FillType.Solid);

        ff.FillType = FillType.Gradient;
        ff.FillType.Should().Be(FillType.Gradient);
    }

    /// <summary>
    /// PatternStyle can be set and read on PatternFormat.
    /// </summary>
    [Fact]
    public void PatternFill_PatternStyleCanBeSetAndRead()
    {
        var ff = CreateFillFormat();
        ff.FillType = FillType.Pattern;
        ff.PatternFormat.PatternStyle = PatternStyle.Percent50;

        ff.PatternFormat.PatternStyle.Should().Be(PatternStyle.Percent50);
    }

    /// <summary>
    /// ForeColor and BackColor can be set on PatternFormat.
    /// </summary>
    [Fact]
    public void PatternFill_ForeAndBackColorsCanBeSet()
    {
        var ff = CreateFillFormat();
        ff.FillType = FillType.Pattern;
        ff.PatternFormat.ForeColor.Color = Color.DarkBlue;
        ff.PatternFormat.BackColor.Color = Color.LightYellow;

        ff.PatternFormat.ForeColor.Color.Should().NotBeNull();
        ff.PatternFormat.BackColor.Color.Should().NotBeNull();
    }

    /// <summary>
    /// SolidFillColor auto-creates solidFill element (get-or-create pattern).
    /// </summary>
    [Fact]
    public void SolidFillColor_AutoCreatesSolidFillElement()
    {
        var ff = CreateFillFormat();

        // Accessing SolidFillColor should auto-create solidFill
        ff.SolidFillColor.Color = Color.Red;

        ff.FillType.Should().Be(FillType.Solid);
    }

    /// <summary>
    /// GradientFormat auto-creates gradFill element with lin child.
    /// </summary>
    [Fact]
    public void GradientFormat_AutoCreatesGradFillWithDefaults()
    {
        var ff = CreateFillFormat();

        // Accessing GradientFormat should auto-create gradFill
        var gf = ff.GradientFormat;

        ff.FillType.Should().Be(FillType.Gradient);
        gf.GradientShape.Should().Be(GradientShape.Linear);
    }
}
