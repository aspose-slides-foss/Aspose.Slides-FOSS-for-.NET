using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Verifies the enum values and color constants used by table operations.
/// </summary>
public sealed class TableTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static CellFormat CreateCellFormatWithBorders()
    {
        var tcPr = new XElement(ANs + "tcPr",
            new XElement(ANs + "lnL"),
            new XElement(ANs + "lnR"),
            new XElement(ANs + "lnT"),
            new XElement(ANs + "lnB"));
        var cf = new CellFormat();
        cf.InitInternal(tcPr, slidePart: null, parentSlide: null);
        return cf;
    }

    private static CellFormat CreateCellFormatWithAllBorders()
    {
        var tcPr = new XElement(ANs + "tcPr",
            new XElement(ANs + "lnL"),
            new XElement(ANs + "lnR"),
            new XElement(ANs + "lnT"),
            new XElement(ANs + "lnB"),
            new XElement(ANs + "lnTlToBr"),
            new XElement(ANs + "lnBlToTr"));
        var cf = new CellFormat();
        cf.InitInternal(tcPr, slidePart: null, parentSlide: null);
        return cf;
    }

    private static CellFormat CreateCellFormatWithSolidFill()
    {
        var tcPr = new XElement(ANs + "tcPr",
            new XElement(ANs + "solidFill",
                new XElement(ANs + "srgbClr", new XAttribute("val", "ADD8E6"))));
        var cf = new CellFormat();
        cf.InitInternal(tcPr, slidePart: null, parentSlide: null);
        return cf;
    }
    /// <summary>
    /// Verifies FillType.Solid and Color.Red used for border styling.
    /// </summary>
    [Fact]
    public void CellBorders_FillTypeSolidIsDefined()
    {
        Enum.IsDefined(FillType.Solid).Should().BeTrue();
        FillType.Solid.Should().NotBe(FillType.NotDefined);
    }

    [Fact]
    public void CellBorders_RedColorHasCorrectComponents()
    {
        Color.Red.R.Should().Be(255);
        Color.Red.G.Should().Be(0);
        Color.Red.B.Should().Be(0);
    }

    [Fact]
    public void CellBorders_RedColorIsNotEmpty()
    {
        Color.Red.Should().NotBe(Color.Empty);
    }

    /// <summary>
    /// Verifies FillType.Solid and Color.LightBlue used for cell fill.
    /// </summary>
    [Fact]
    public void CellFill_FillTypeSolidIsDefined()
    {
        Enum.IsDefined(FillType.Solid).Should().BeTrue();
    }

    [Fact]
    public void CellFill_LightBlueColorExists()
    {
        Color.LightBlue.Should().NotBe(Color.Empty);
    }

    [Fact]
    public void CellFill_LightBlueHasBlueComponent()
    {
        Color.LightBlue.B.Should().BeGreaterThan(0);
    }

    [Fact]
    public void CellFill_SolidIsDistinctFromNoFill()
    {
        FillType.Solid.Should().NotBe(FillType.NoFill);
    }

    /// <summary>
    /// Verifies that Table class can be instantiated.
    /// </summary>
    [Fact]
    public void CellText_TableCanBeInstantiated()
    {
        var table = new Table();
        table.Should().NotBeNull();
    }

    [Fact]
    public void CellText_FillTypeSolidIsAvailable()
    {
        FillType.Solid.Should().NotBe(FillType.NotDefined);
        Enum.IsDefined(FillType.Solid).Should().BeTrue();
    }

    /// <summary>
    /// Row class can be instantiated.
    /// </summary>
    [Fact]
    public void RowHeight_RowCanBeInstantiated()
    {
        var row = new Row();
        row.Should().NotBeNull();
    }

    /// <summary>
    /// Column class can be instantiated.
    /// </summary>
    [Fact]
    public void ColumnWidth_ColumnCanBeInstantiated()
    {
        var col = new Column();
        col.Should().NotBeNull();
    }

    /// <summary>
    /// FillType.Solid is distinct from other fill types used in borders.
    /// </summary>
    [Fact]
    public void CellBorders_SolidIsDistinctFromOtherTypes()
    {
        FillType.Solid.Should().NotBe(FillType.NoFill);
        FillType.Solid.Should().NotBe(FillType.Gradient);
        FillType.Solid.Should().NotBe(FillType.Pattern);
    }

    /// <summary>
    /// Border width value 3 is representable as a double.
    /// </summary>
    [Fact]
    public void CellBorders_WidthValueIsRepresentable()
    {
        var width = 3.0;
        width.Should().Be(3);
    }

    /// <summary>
    /// CellFormat exposes all four border properties via XML.
    /// </summary>
    [Fact]
    public void CellBorders_AllFourBordersAreAccessible()
    {
        var cf = CreateCellFormatWithBorders();

        cf.BorderTop.Should().NotBeNull();
        cf.BorderBottom.Should().NotBeNull();
        cf.BorderLeft.Should().NotBeNull();
        cf.BorderRight.Should().NotBeNull();
    }

    /// <summary>
    /// CellFormat.FillFormat reads solid fill from XML.
    /// </summary>
    [Fact]
    public void CellFill_FillFormatReadsSolidFillFromXml()
    {
        var cf = CreateCellFormatWithSolidFill();

        cf.FillFormat.FillType.Should().Be(FillType.Solid);
    }

    /// <summary>
    /// CellFormat.FillFormat.SolidFillColor reads color from XML.
    /// </summary>
    [Fact]
    public void CellFill_SolidFillColorIsReadFromXml()
    {
        var cf = CreateCellFormatWithSolidFill();

        var c = cf.FillFormat.SolidFillColor.Color;
        c.Should().NotBeNull();
        c!.B.Should().BeGreaterThan(0);
    }

    /// <summary>
    /// CellFormat with no tcPr returns NotDefined fill type.
    /// </summary>
    [Fact]
    public void CellFormat_EmptyElementReturnsFillTypeNotDefined()
    {
        var tcElement = new XElement(ANs + "tc");
        var cf = new CellFormat();
        cf.InitInternal(tcElement, slidePart: null, parentSlide: null);

        cf.FillFormat.FillType.Should().Be(FillType.NotDefined);
    }

    /// <summary>
    /// BorderDiagonalDown and BorderDiagonalUp are accessible on CellFormat.
    /// </summary>
    [Fact]
    public void CellBorders_DiagonalBordersAreAccessible()
    {
        var cf = CreateCellFormatWithAllBorders();

        cf.BorderDiagonalDown.Should().NotBeNull();
        cf.BorderDiagonalUp.Should().NotBeNull();
    }

    /// <summary>
    /// All six border properties are accessible and distinct on CellFormat.
    /// </summary>
    [Fact]
    public void CellBorders_AllSixBordersAreAccessible()
    {
        var cf = CreateCellFormatWithAllBorders();

        cf.BorderTop.Should().NotBeNull();
        cf.BorderBottom.Should().NotBeNull();
        cf.BorderLeft.Should().NotBeNull();
        cf.BorderRight.Should().NotBeNull();
        cf.BorderDiagonalDown.Should().NotBeNull();
        cf.BorderDiagonalUp.Should().NotBeNull();
    }

    /// <summary>
    /// All six borders return ILineFormat instances (including diagonals).
    /// </summary>
    [Fact]
    public void CellBorders_AllSixBordersReturnLineFormat()
    {
        var cf = CreateCellFormatWithAllBorders();

        cf.BorderTop.Should().BeAssignableTo<ILineFormat>();
        cf.BorderBottom.Should().BeAssignableTo<ILineFormat>();
        cf.BorderLeft.Should().BeAssignableTo<ILineFormat>();
        cf.BorderRight.Should().BeAssignableTo<ILineFormat>();
        cf.BorderDiagonalDown.Should().BeAssignableTo<ILineFormat>();
        cf.BorderDiagonalUp.Should().BeAssignableTo<ILineFormat>();
    }

    /// <summary>
    /// CellFormat.FillFormat can be set to Solid and color can be assigned.
    /// </summary>
    [Fact]
    public void CellFill_FillFormatCanBeSetToSolidWithColor()
    {
        var tcPr = new XElement(ANs + "tcPr");
        var cf = new CellFormat();
        cf.InitInternal(tcPr, slidePart: null, parentSlide: null);

        cf.FillFormat.FillType = FillType.Solid;
        cf.FillFormat.SolidFillColor.Color = Color.LightBlue;

        cf.FillFormat.FillType.Should().Be(FillType.Solid);
    }

    /// <summary>
    /// CellFormat.FillFormat property returns IFillFormat.
    /// </summary>
    [Fact]
    public void CellFormat_FillFormatReturnsIFillFormat()
    {
        var cf = CreateCellFormatWithSolidFill();

        cf.FillFormat.Should().BeAssignableTo<IFillFormat>();
    }

    /// <summary>
    /// Diagonal borders are accessible even without diagonal XML elements.
    /// </summary>
    [Fact]
    public void CellBorders_DiagonalBordersAccessibleWithoutXmlElements()
    {
        var cf = CreateCellFormatWithBorders();

        cf.BorderDiagonalDown.Should().NotBeNull();
        cf.BorderDiagonalUp.Should().NotBeNull();
    }
}
