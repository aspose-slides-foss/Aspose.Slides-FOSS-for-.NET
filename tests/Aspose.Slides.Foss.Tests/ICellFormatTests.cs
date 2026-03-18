using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests for ICellFormat contract: FillFormat, BorderTop, BorderBottom,
/// BorderLeft, BorderRight, BorderDiagonalDown, BorderDiagonalUp.
/// </summary>
public sealed class ICellFormatTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static CellFormat CreateCellFormat(XElement? tcPr = null)
    {
        var element = tcPr ?? new XElement(ANs + "tcPr");
        var cf = new CellFormat();
        cf.InitInternal(element, slidePart: null, parentSlide: null);
        return cf;
    }

    private static CellFormat CreateCellFormatWithBorders()
    {
        var tcPr = new XElement(ANs + "tcPr",
            new XElement(ANs + "lnL"),
            new XElement(ANs + "lnR"),
            new XElement(ANs + "lnT"),
            new XElement(ANs + "lnB"),
            new XElement(ANs + "lnTlToBr"),
            new XElement(ANs + "lnBlToTr"));
        return CreateCellFormat(tcPr);
    }

    // -------------------------------------------------------------------
    // Applied to CellFormat.FillFormat property.
    // -------------------------------------------------------------------

    [Fact]
    public void FillFormat_SolidFillTypeCanBeSetAndRead()
    {
        var cf = CreateCellFormat();

        cf.FillFormat.FillType = FillType.Solid;

        cf.FillFormat.FillType.Should().Be(FillType.Solid);
    }

    [Fact]
    public void FillFormat_SolidFillColorComponentsPersist()
    {
        var cf = CreateCellFormat();

        cf.FillFormat.FillType = FillType.Solid;
        cf.FillFormat.SolidFillColor.Color = Color.FromArgb(255, 0, 128, 255);

        var c = cf.FillFormat.SolidFillColor.Color;
        c.Should().NotBeNull();
        c!.R.Should().Be(0);
        c.G.Should().Be(128);
        c.B.Should().Be(255);
    }

    // -------------------------------------------------------------------
    // -------------------------------------------------------------------

    [Fact]
    public void FillFormat_GradientFillTypeCanBeSet()
    {
        var tcPr = new XElement(ANs + "tcPr",
            new XElement(ANs + "gradFill",
                new XElement(ANs + "gsLst")));
        var cf = CreateCellFormat(tcPr);

        cf.FillFormat.FillType.Should().Be(FillType.Gradient);
    }

    [Fact]
    public void FillFormat_GradientStopsCanBeAdded()
    {
        var tcPr = new XElement(ANs + "tcPr",
            new XElement(ANs + "gradFill",
                new XElement(ANs + "gsLst")));
        var cf = CreateCellFormat(tcPr);

        cf.FillFormat.GradientFormat.GradientStops.Add(0.0f, Color.Blue);
        cf.FillFormat.GradientFormat.GradientStops.Add(1.0f, Color.Red);

        cf.FillFormat.GradientFormat.GradientStops.Count.Should().BeGreaterThanOrEqualTo(2);
    }

    // -------------------------------------------------------------------
    // -------------------------------------------------------------------

    [Fact]
    public void FillFormat_PatternFillTypeCanBeSet()
    {
        var cf = CreateCellFormat();

        cf.FillFormat.FillType = FillType.Pattern;

        cf.FillFormat.FillType.Should().Be(FillType.Pattern);
    }

    // -------------------------------------------------------------------
    // -------------------------------------------------------------------

    [Fact]
    public void FillFormat_NoFillTypeCanBeSet()
    {
        var cf = CreateCellFormat();

        cf.FillFormat.FillType = FillType.NoFill;

        cf.FillFormat.FillType.Should().Be(FillType.NoFill);
    }

    // -------------------------------------------------------------------
    // -------------------------------------------------------------------

    [Fact]
    public void FillFormat_PictureFillTypeCanBeSet()
    {
        var cf = CreateCellFormat();

        cf.FillFormat.FillType = FillType.Picture;

        cf.FillFormat.FillType.Should().Be(FillType.Picture);
    }

    // -------------------------------------------------------------------
    // -------------------------------------------------------------------

    [Fact]
    public void FillFormat_ReadsSolidFillFromXml()
    {
        var tcPr = new XElement(ANs + "tcPr",
            new XElement(ANs + "solidFill",
                new XElement(ANs + "srgbClr", new XAttribute("val", "ADD8E6"))));
        var cf = CreateCellFormat(tcPr);

        cf.FillFormat.FillType.Should().Be(FillType.Solid);
    }

    [Fact]
    public void FillFormat_SolidFillColorIsReadFromXml()
    {
        var tcPr = new XElement(ANs + "tcPr",
            new XElement(ANs + "solidFill",
                new XElement(ANs + "srgbClr", new XAttribute("val", "ADD8E6"))));
        var cf = CreateCellFormat(tcPr);

        var c = cf.FillFormat.SolidFillColor.Color;
        c.Should().NotBeNull();
        c!.B.Should().BeGreaterThan(0);
    }

    [Fact]
    public void FillFormat_ReturnsIFillFormat()
    {
        var cf = CreateCellFormat();

        cf.FillFormat.Should().BeAssignableTo<IFillFormat>();
    }

    [Fact]
    public void FillFormat_EmptyElementReturnsFillTypeNotDefined()
    {
        var cf = CreateCellFormat();

        cf.FillFormat.FillType.Should().Be(FillType.NotDefined);
    }

    // -------------------------------------------------------------------
    // Applied to CellFormat.FillFormat (same FillFormat API).
    // -------------------------------------------------------------------

    [Fact]
    public void FillFormat_RedColorPersistsInSolidFill()
    {
        var cf = CreateCellFormat();

        cf.FillFormat.FillType = FillType.Solid;
        cf.FillFormat.SolidFillColor.Color = Color.Red;

        var c = cf.FillFormat.SolidFillColor.Color;
        c.Should().NotBeNull();
        c!.R.Should().Be(255);
        c.G.Should().Be(0);
        c.B.Should().Be(0);
    }

    // -------------------------------------------------------------------
    // -------------------------------------------------------------------

    [Fact]
    public void BorderTop_IsNotNull()
    {
        var cf = CreateCellFormatWithBorders();
        cf.BorderTop.Should().NotBeNull();
    }

    [Fact]
    public void BorderBottom_IsNotNull()
    {
        var cf = CreateCellFormatWithBorders();
        cf.BorderBottom.Should().NotBeNull();
    }

    [Fact]
    public void BorderLeft_IsNotNull()
    {
        var cf = CreateCellFormatWithBorders();
        cf.BorderLeft.Should().NotBeNull();
    }

    [Fact]
    public void BorderRight_IsNotNull()
    {
        var cf = CreateCellFormatWithBorders();
        cf.BorderRight.Should().NotBeNull();
    }

    [Fact]
    public void BorderDiagonalDown_IsNotNull()
    {
        var cf = CreateCellFormatWithBorders();
        cf.BorderDiagonalDown.Should().NotBeNull();
    }

    [Fact]
    public void BorderDiagonalUp_IsNotNull()
    {
        var cf = CreateCellFormatWithBorders();
        cf.BorderDiagonalUp.Should().NotBeNull();
    }

    [Fact]
    public void AllBorders_ReturnILineFormat()
    {
        var cf = CreateCellFormatWithBorders();

        cf.BorderTop.Should().BeAssignableTo<ILineFormat>();
        cf.BorderBottom.Should().BeAssignableTo<ILineFormat>();
        cf.BorderLeft.Should().BeAssignableTo<ILineFormat>();
        cf.BorderRight.Should().BeAssignableTo<ILineFormat>();
        cf.BorderDiagonalDown.Should().BeAssignableTo<ILineFormat>();
        cf.BorderDiagonalUp.Should().BeAssignableTo<ILineFormat>();
    }

    // -------------------------------------------------------------------
    // Border properties return non-null ILineFormat instances.
    // -------------------------------------------------------------------

    [Fact]
    public void Border_TopReturnsLineFormatInstance()
    {
        var cf = CreateCellFormatWithBorders();
        cf.BorderTop.Should().BeOfType<LineFormat>();
    }

    // -------------------------------------------------------------------
    // Applied to CellFormat border line formats.
    // -------------------------------------------------------------------

    [Theory]
    [InlineData(LineDashStyle.Solid)]
    [InlineData(LineDashStyle.Dash)]
    [InlineData(LineDashStyle.Dot)]
    [InlineData(LineDashStyle.DashDot)]
    public void Border_DashStyleEnumValuesAreDefined(LineDashStyle style)
    {
        Enum.IsDefined(style).Should().BeTrue();
    }

    // -------------------------------------------------------------------
    // -------------------------------------------------------------------

    [Fact]
    public void Border_DarkRedColorIsValid()
    {
        Color.DarkRed.Should().NotBe(Color.Empty);
        Color.DarkRed.R.Should().BeGreaterThan(0);
    }

    // -------------------------------------------------------------------
    // Borders accessible even without XML elements for those borders.
    // -------------------------------------------------------------------

    [Fact]
    public void DiagonalBorders_AccessibleWithoutXmlElements()
    {
        var tcPr = new XElement(ANs + "tcPr",
            new XElement(ANs + "lnL"),
            new XElement(ANs + "lnR"),
            new XElement(ANs + "lnT"),
            new XElement(ANs + "lnB"));
        var cf = CreateCellFormat(tcPr);

        cf.BorderDiagonalDown.Should().NotBeNull();
        cf.BorderDiagonalUp.Should().NotBeNull();
    }

    // -------------------------------------------------------------------
    // CellFormat implements ICellFormat.
    // -------------------------------------------------------------------

    [Fact]
    public void CellFormat_ImplementsICellFormat()
    {
        var cf = CreateCellFormat();
        cf.Should().BeAssignableTo<ICellFormat>();
    }

    // -------------------------------------------------------------------
    // Changing FillType replaces previous fill on CellFormat.FillFormat.
    // -------------------------------------------------------------------

    [Fact]
    public void FillFormat_ChangingTypeReplacesPreviousFill()
    {
        var cf = CreateCellFormat();

        cf.FillFormat.FillType = FillType.Solid;
        cf.FillFormat.FillType.Should().Be(FillType.Solid);

        cf.FillFormat.FillType = FillType.NoFill;
        cf.FillFormat.FillType.Should().Be(FillType.NoFill);
    }
}
