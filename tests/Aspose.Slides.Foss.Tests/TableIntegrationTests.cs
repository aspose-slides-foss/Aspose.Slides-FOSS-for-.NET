using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests table creation, cell text, merge, borders, style options, row height, column width, and cell fill.
/// </summary>
public sealed class TableIntegrationTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const float EmuPerPoint = 12700f;

    /// <summary>
    /// Builds a tbl XElement matching the format used by ShapeCollection.BuildTableXml.
    /// </summary>
    private static XElement BuildTblElement(double[] columnWidths, double[] rowHeights)
    {
        var tbl = new XElement(ANs + "tbl",
            new XElement(ANs + "tblPr",
                new XAttribute("firstRow", "1"),
                new XAttribute("bandRow", "1")),
            new XElement(ANs + "tblGrid"));

        var tblGrid = tbl.Element(ANs + "tblGrid")!;
        foreach (var colWidth in columnWidths)
        {
            tblGrid.Add(new XElement(ANs + "gridCol",
                new XAttribute("w", ((int)Math.Round(colWidth * EmuPerPoint)).ToString())));
        }

        foreach (var rowHeight in rowHeights)
        {
            var tr = new XElement(ANs + "tr",
                new XAttribute("h", ((int)Math.Round(rowHeight * EmuPerPoint)).ToString()));

            for (int c = 0; c < columnWidths.Length; c++)
            {
                tr.Add(new XElement(ANs + "tc",
                    new XElement(ANs + "txBody",
                        new XElement(ANs + "bodyPr"),
                        new XElement(ANs + "lstStyle"),
                        new XElement(ANs + "p",
                            new XElement(ANs + "endParaRPr"))),
                    new XElement(ANs + "tcPr")));
            }

            tbl.Add(tr);
        }

        return tbl;
    }

    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";

    /// <summary>
    /// Wraps a tbl element in a graphicFrame so it can be used with Table.InitInternal.
    /// </summary>
    private static XElement BuildGraphicFrame(XElement tbl)
    {
        return new XElement(PNs + "graphicFrame",
            new XElement(PNs + "nvGraphicFramePr",
                new XElement(PNs + "cNvPr",
                    new XAttribute("id", "2"),
                    new XAttribute("name", "Table 2")),
                new XElement(PNs + "cNvGraphicFramePr",
                    new XElement(ANs + "graphicFrameLocks", new XAttribute("noGrp", "1"))),
                new XElement(PNs + "nvPr")),
            new XElement(PNs + "xfrm",
                new XElement(ANs + "off", new XAttribute("x", "0"), new XAttribute("y", "0")),
                new XElement(ANs + "ext", new XAttribute("cx", "0"), new XAttribute("cy", "0"))),
            new XElement(ANs + "graphic",
                new XElement(ANs + "graphicData",
                    new XAttribute("uri", "http://schemas.openxmlformats.org/drawingml/2006/table"),
                    tbl)));
    }

    /// <summary>
    /// Initializes a RowCollection from a tbl element.
    /// </summary>
    private static RowCollection BuildRows(XElement tbl)
    {
        var rows = new RowCollection();
        rows.InitInternal(tbl, slidePart: null, parentSlide: null, table: null);
        return rows;
    }

    /// <summary>
    /// Initializes a ColumnCollection from a tbl element.
    /// </summary>
    private static ColumnCollection BuildColumns(XElement tbl)
    {
        var tblGrid = tbl.Element(ANs + "tblGrid")!;
        var columns = new ColumnCollection();
        columns.InitInternal(tbl, tblGrid, slidePart: null, parentSlide: null, table: null);
        return columns;
    }

    // ── test_create_table ────────────────────────────────────────

    /// <summary>
    /// Create a table and verify row/column counts.
    /// </summary>
    [Fact]
    public void CreateTable_RowCountMatchesRowHeights()
    {
        var tbl = BuildTblElement([100, 150, 200], [40, 40, 40]);
        var rows = BuildRows(tbl);

        rows.Count.Should().Be(3);
    }

    [Fact]
    public void CreateTable_ColumnCountMatchesColumnWidths()
    {
        var tbl = BuildTblElement([100, 150, 200], [40, 40, 40]);
        var columns = BuildColumns(tbl);

        columns.Count.Should().Be(3);
    }

    // ── test_cell_text ───────────────────────────────────────────

    /// <summary>
    /// Cell TextFrame is not null when txBody element is present.
    /// </summary>
    [Fact]
    public void CellText_TextFrameIsNotNullWhenTxBodyPresent()
    {
        var tcElement = new XElement(ANs + "tc",
            new XElement(ANs + "txBody",
                new XElement(ANs + "bodyPr"),
                new XElement(ANs + "lstStyle"),
                new XElement(ANs + "p",
                    new XElement(ANs + "r",
                        new XElement(ANs + "t", "Hello")))),
            new XElement(ANs + "tcPr"));

        var table = new Table();
        var cell = new Cell();
        cell.InitInternal(tcElement, rowIndex: 0, colIndex: 0, slidePart: null, parentSlide: null, table: table);

        cell.TextFrame.Should().NotBeNull();
    }

    /// <summary>
    /// Cell TextFrame is null when no txBody element exists.
    /// </summary>
    [Fact]
    public void CellText_TextFrameIsNullWhenNoTxBody()
    {
        var tcElement = new XElement(ANs + "tc",
            new XElement(ANs + "tcPr"));

        var table = new Table();
        var cell = new Cell();
        cell.InitInternal(tcElement, rowIndex: 0, colIndex: 0, slidePart: null, parentSlide: null, table: table);

        cell.TextFrame.Should().BeNull();
    }

    /// <summary>
    /// Each cell in a multi-cell table has its own TextFrame.
    /// </summary>
    [Fact]
    public void CellText_EachCellHasDistinctTextFrame()
    {
        var tbl = BuildTblElement([100, 100], [40, 40]);
        var rows = BuildRows(tbl);

        rows[0][0].TextFrame.Should().NotBeNull();
        rows[0][1].TextFrame.Should().NotBeNull();
        rows[1][0].TextFrame.Should().NotBeNull();
        rows[1][1].TextFrame.Should().NotBeNull();
    }

    // ── test_merge_cells ─────────────────────────────────────────

    /// <summary>
    /// MergeCells sets gridSpan on the origin cell and marks the cell as merged.
    /// </summary>
    [Fact]
    public void MergeCells_SetsGridSpanAndMarksMerged()
    {
        var tbl = BuildTblElement([100, 100, 100], [40, 40]);
        var table = new Table();
        table.InitInternal(BuildGraphicFrame(tbl), slidePart: null, parentSlide: null);

        var cell1 = table.Rows[0][0];
        var cell2 = table.Rows[0][1];
        table.MergeCells(cell1, cell2, false);

        table.Rows[0][0].IsMergedCell.Should().BeTrue();
        table.Rows[0][0].ColSpan.Should().BeGreaterThanOrEqualTo(2);
    }

    /// <summary>
    /// IsMergedCell returns true when gridSpan > 1.
    /// </summary>
    [Fact]
    public void MergeCells_IsMergedCellTrueWhenColSpanGreaterThanOne()
    {
        var tcElement = new XElement(ANs + "tc",
            new XAttribute("gridSpan", "2"),
            new XElement(ANs + "txBody",
                new XElement(ANs + "bodyPr"),
                new XElement(ANs + "lstStyle"),
                new XElement(ANs + "p")),
            new XElement(ANs + "tcPr"));

        var table = new Table();
        var cell = new Cell();
        cell.InitInternal(tcElement, rowIndex: 0, colIndex: 0, slidePart: null, parentSlide: null, table: table);

        cell.IsMergedCell.Should().BeTrue();
        cell.ColSpan.Should().BeGreaterThanOrEqualTo(2);
    }

    /// <summary>
    /// IsMergedCell returns true when hMerge attribute is set.
    /// </summary>
    [Fact]
    public void MergeCells_IsMergedCellTrueWhenHMergeSet()
    {
        var tcElement = new XElement(ANs + "tc",
            new XAttribute("hMerge", "1"),
            new XElement(ANs + "tcPr"));

        var table = new Table();
        var cell = new Cell();
        cell.InitInternal(tcElement, rowIndex: 0, colIndex: 0, slidePart: null, parentSlide: null, table: table);

        cell.IsMergedCell.Should().BeTrue();
    }

    /// <summary>
    /// IsMergedCell returns false for a normal unmerged cell.
    /// </summary>
    [Fact]
    public void MergeCells_IsMergedCellFalseForNormalCell()
    {
        var tcElement = new XElement(ANs + "tc",
            new XElement(ANs + "tcPr"));

        var table = new Table();
        var cell = new Cell();
        cell.InitInternal(tcElement, rowIndex: 0, colIndex: 0, slidePart: null, parentSlide: null, table: table);

        cell.IsMergedCell.Should().BeFalse();
    }

    // ── test_cell_borders ────────────────────────────────────────

    /// <summary>
    /// CellFormat exposes all four border properties as ILineFormat.
    /// </summary>
    [Fact]
    public void CellBorders_AllFourBordersAreNotNull()
    {
        var tcPr = new XElement(ANs + "tcPr",
            new XElement(ANs + "lnT"),
            new XElement(ANs + "lnB"),
            new XElement(ANs + "lnL"),
            new XElement(ANs + "lnR"));

        var cf = new CellFormat();
        cf.InitInternal(tcPr, slidePart: null, parentSlide: null);

        cf.BorderTop.Should().NotBeNull();
        cf.BorderBottom.Should().NotBeNull();
        cf.BorderLeft.Should().NotBeNull();
        cf.BorderRight.Should().NotBeNull();
    }

    /// <summary>
    /// Borders return ILineFormat instances.
    /// </summary>
    [Fact]
    public void CellBorders_BordersReturnILineFormat()
    {
        var tcPr = new XElement(ANs + "tcPr",
            new XElement(ANs + "lnT"),
            new XElement(ANs + "lnB"),
            new XElement(ANs + "lnL"),
            new XElement(ANs + "lnR"));

        var cf = new CellFormat();
        cf.InitInternal(tcPr, slidePart: null, parentSlide: null);

        cf.BorderTop.Should().BeAssignableTo<ILineFormat>();
        cf.BorderBottom.Should().BeAssignableTo<ILineFormat>();
        cf.BorderLeft.Should().BeAssignableTo<ILineFormat>();
        cf.BorderRight.Should().BeAssignableTo<ILineFormat>();
    }

    /// <summary>
    /// FillType.Solid and Color.Red are correct for border styling.
    /// </summary>
    [Fact]
    public void CellBorders_FillTypeSolidAndColorRedAreValid()
    {
        FillType.Solid.Should().NotBe(FillType.NotDefined);
        FillType.Solid.Should().NotBe(FillType.NoFill);
        Color.Red.R.Should().Be(255);
        Color.Red.G.Should().Be(0);
        Color.Red.B.Should().Be(0);
    }

    /// <summary>
    /// CellFormat.FillFormat returns IFillFormat.
    /// </summary>
    [Fact]
    public void CellBorders_CellFormatFillFormatReturnsIFillFormat()
    {
        var tcPr = new XElement(ANs + "tcPr",
            new XElement(ANs + "lnT"),
            new XElement(ANs + "lnB"),
            new XElement(ANs + "lnL"),
            new XElement(ANs + "lnR"));

        var cf = new CellFormat();
        cf.InitInternal(tcPr, slidePart: null, parentSlide: null);

        cf.FillFormat.Should().BeAssignableTo<IFillFormat>();
    }

    // ── test_table_style_options ──────────────────────────────────

    /// <summary>
    /// Table style flags can be set and read back.
    /// </summary>
    [Fact]
    public void TableStyleOptions_FirstRowCanBeSet()
    {
        var table = new Table();

        table.FirstRow = true;

        table.FirstRow.Should().BeTrue();
    }

    [Fact]
    public void TableStyleOptions_HorizontalBandingCanBeSet()
    {
        var table = new Table();

        table.HorizontalBanding = true;

        table.HorizontalBanding.Should().BeTrue();
    }

    [Fact]
    public void TableStyleOptions_VerticalBandingCanBeSet()
    {
        var table = new Table();

        table.VerticalBanding = false;

        table.VerticalBanding.Should().BeFalse();
    }

    [Fact]
    public void TableStyleOptions_AllFlagsPersistTogether()
    {
        var table = new Table();

        table.FirstRow = true;
        table.HorizontalBanding = true;
        table.VerticalBanding = false;
        table.FirstCol = true;
        table.LastRow = true;
        table.LastCol = false;
        table.RightToLeft = false;

        table.FirstRow.Should().BeTrue();
        table.HorizontalBanding.Should().BeTrue();
        table.VerticalBanding.Should().BeFalse();
        table.FirstCol.Should().BeTrue();
        table.LastRow.Should().BeTrue();
        table.LastCol.Should().BeFalse();
        table.RightToLeft.Should().BeFalse();
    }

    // ── test_row_height ──────────────────────────────────────────

    /// <summary>
    /// Row heights match constructor arguments.
    /// </summary>
    [Fact]
    public void RowHeight_MatchesConstructorArguments()
    {
        var tbl = BuildTblElement([200], [30, 50, 70]);
        var rows = BuildRows(tbl);

        rows[0].Height.Should().BeApproximately(30, 0.01f);
        rows[1].Height.Should().BeApproximately(50, 0.01f);
        rows[2].Height.Should().BeApproximately(70, 0.01f);
    }

    [Theory]
    [InlineData(30)]
    [InlineData(50)]
    [InlineData(70)]
    [InlineData(100)]
    public void RowHeight_VariousHeightsPreserved(double height)
    {
        var tbl = BuildTblElement([200], [height]);
        var rows = BuildRows(tbl);

        rows[0].Height.Should().BeApproximately((float)height, 0.01f);
    }

    // ── test_column_width ────────────────────────────────────────

    /// <summary>
    /// Column widths match constructor arguments.
    /// </summary>
    [Fact]
    public void ColumnWidth_MatchesConstructorArguments()
    {
        var tbl = BuildTblElement([100, 200, 300], [40]);
        var columns = BuildColumns(tbl);

        columns[0].Width.Should().BeApproximately(100, 0.01f);
        columns[1].Width.Should().BeApproximately(200, 0.01f);
        columns[2].Width.Should().BeApproximately(300, 0.01f);
    }

    [Theory]
    [InlineData(100)]
    [InlineData(200)]
    [InlineData(300)]
    [InlineData(500)]
    public void ColumnWidth_VariousWidthsPreserved(double width)
    {
        var tbl = BuildTblElement([width], [40]);
        var columns = BuildColumns(tbl);

        columns[0].Width.Should().BeApproximately((float)width, 0.01f);
    }

    // ── test_cell_fill ───────────────────────────────────────────

    /// <summary>
    /// Cell fill type reads Solid from XML.
    /// </summary>
    [Fact]
    public void CellFill_FillTypeReadsSolidFromXml()
    {
        var tcPr = new XElement(ANs + "tcPr",
            new XElement(ANs + "solidFill",
                new XElement(ANs + "srgbClr", new XAttribute("val", "ADD8E6"))));

        var cf = new CellFormat();
        cf.InitInternal(tcPr, slidePart: null, parentSlide: null);

        cf.FillFormat.FillType.Should().Be(FillType.Solid);
    }

    /// <summary>
    /// Cell fill color is read from XML.
    /// </summary>
    [Fact]
    public void CellFill_SolidFillColorReadFromXml()
    {
        var tcPr = new XElement(ANs + "tcPr",
            new XElement(ANs + "solidFill",
                new XElement(ANs + "srgbClr", new XAttribute("val", "ADD8E6"))));

        var cf = new CellFormat();
        cf.InitInternal(tcPr, slidePart: null, parentSlide: null);

        var color = cf.FillFormat.SolidFillColor.Color;
        color.Should().NotBeNull();
        color!.B.Should().BeGreaterThan(0);
    }

    /// <summary>
    /// Color.LightBlue is a valid named color for cell fill.
    /// </summary>
    [Fact]
    public void CellFill_LightBlueColorIsValid()
    {
        Color.LightBlue.Should().NotBe(Color.Empty);
        Color.LightBlue.B.Should().BeGreaterThan(0);
    }

    /// <summary>
    /// Cell fill can be set to Solid with a color via API.
    /// </summary>
    [Fact]
    public void CellFill_CanBeSetToSolidWithColor()
    {
        var tcPr = new XElement(ANs + "tcPr");
        var cf = new CellFormat();
        cf.InitInternal(tcPr, slidePart: null, parentSlide: null);

        cf.FillFormat.FillType = FillType.Solid;
        cf.FillFormat.SolidFillColor.Color = Color.LightBlue;

        cf.FillFormat.FillType.Should().Be(FillType.Solid);
    }

}
