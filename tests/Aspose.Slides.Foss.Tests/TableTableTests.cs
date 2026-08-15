using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Exercises Table properties and methods through the full XML-backed InitInternal path.
/// </summary>
public sealed class TableTableTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private const float EmuPerPoint = 12700f;

    /// <summary>
    /// Builds a tbl XElement matching ShapeCollection.BuildTableXml format.
    /// </summary>
    private static XElement BuildTbl(double[] columnWidths, double[] rowHeights)
    {
        var tbl = new XElement(ANs + "tbl",
            new XElement(ANs + "tblPr",
                new XAttribute("firstRow", "1"),
                new XAttribute("bandRow", "1")),
            new XElement(ANs + "tblGrid"));

        var tblGrid = tbl.Element(ANs + "tblGrid")!;
        foreach (var w in columnWidths)
            tblGrid.Add(new XElement(ANs + "gridCol",
                new XAttribute("w", ((int)Math.Round(w * EmuPerPoint)).ToString())));

        foreach (var h in rowHeights)
        {
            var tr = new XElement(ANs + "tr",
                new XAttribute("h", ((int)Math.Round(h * EmuPerPoint)).ToString()));
            for (var c = 0; c < columnWidths.Length; c++)
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

    private static XElement WrapInGraphicFrame(XElement tbl)
    {
        return new XElement(PNs + "graphicFrame",
            new XElement(PNs + "nvGraphicFramePr",
                new XElement(PNs + "cNvPr", new XAttribute("id", "2"), new XAttribute("name", "Table 2")),
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

    private static Table CreateTable(double[] columnWidths, double[] rowHeights)
    {
        var tbl = BuildTbl(columnWidths, rowHeights);
        var frame = WrapInGraphicFrame(tbl);
        var table = new Table();
        table.InitInternal(frame, slidePart: null, parentSlide: null);
        return table;
    }

    // ── test_create_table ────────────────────────────────────────

    /// <summary>
    /// </summary>
    [Fact]
    public void CreateTable_RowAndColumnCountsMatch()
    {
        var table = CreateTable([100, 150, 200], [40, 40, 40]);

        table.Rows.Count.Should().Be(3);
        table.Columns.Count.Should().Be(3);
    }

    // ── test_cell_text ───────────────────────────────────────────

    /// <summary>
    /// </summary>
    [Fact]
    public void CellText_AllCellsHaveTextFrame()
    {
        var table = CreateTable([100, 100], [40, 40]);

        table.Rows[0][0].TextFrame.Should().NotBeNull();
        table.Rows[0][1].TextFrame.Should().NotBeNull();
        table.Rows[1][0].TextFrame.Should().NotBeNull();
        table.Rows[1][1].TextFrame.Should().NotBeNull();
    }

    // ── test_merge_cells ─────────────────────────────────────────

    /// <summary>
    /// </summary>
    [Fact]
    public void MergeCells_OriginCellHasColSpanAndIsMerged()
    {
        var table = CreateTable([100, 100, 100], [40, 40]);

        var cell1 = table.Rows[0][0];
        var cell2 = table.Rows[0][1];
        table.MergeCells(cell1, cell2, false);

        // Re-read from XML
        table.Rows[0][0].IsMergedCell.Should().BeTrue();
        table.Rows[0][0].ColSpan.Should().BeGreaterThanOrEqualTo(2);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void MergeCells_SpannedCellHasMergeFlag()
    {
        var table = CreateTable([100, 100, 100], [40, 40]);

        table.MergeCells(table.Rows[0][0], table.Rows[0][1], false);

        table.Rows[0][1].IsMergedCell.Should().BeTrue();
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void MergeCells_VerticalMergeSetsRowSpan()
    {
        var table = CreateTable([100, 100], [40, 40, 40]);

        table.MergeCells(table.Rows[0][0], table.Rows[1][0], false);

        table.Rows[0][0].RowSpan.Should().BeGreaterThanOrEqualTo(2);
        table.Rows[0][0].IsMergedCell.Should().BeTrue();
        table.Rows[1][0].IsMergedCell.Should().BeTrue();
    }

    // ── test_cell_borders ────────────────────────────────────────

    /// <summary>
    /// </summary>
    [Fact]
    public void CellBorders_AccessibleViaCellFormat()
    {
        var table = CreateTable([150], [50]);
        var fmt = table.Rows[0][0].CellFormat;

        fmt.BorderTop.Should().NotBeNull();
        fmt.BorderBottom.Should().NotBeNull();
        fmt.BorderLeft.Should().NotBeNull();
        fmt.BorderRight.Should().NotBeNull();
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void CellBorders_ReturnLineFormatInstances()
    {
        var table = CreateTable([150], [50]);
        var fmt = table.Rows[0][0].CellFormat;

        fmt.BorderTop.Should().BeAssignableTo<ILineFormat>();
        fmt.BorderBottom.Should().BeAssignableTo<ILineFormat>();
        fmt.BorderLeft.Should().BeAssignableTo<ILineFormat>();
        fmt.BorderRight.Should().BeAssignableTo<ILineFormat>();
    }

    // ── test_table_style_options ──────────────────────────────────

    /// <summary>
    /// </summary>
    [Fact]
    public void TableStyleOptions_FlagsPersistViaXml()
    {
        var table = CreateTable([120, 120], [40, 40, 40]);

        table.FirstRow = true;
        table.HorizontalBanding = true;
        table.VerticalBanding = false;

        table.FirstRow.Should().BeTrue();
        table.HorizontalBanding.Should().BeTrue();
        table.VerticalBanding.Should().BeFalse();
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void TableStyleOptions_InitialValuesFromXml()
    {
        // BuildTbl sets firstRow="1" and bandRow="1" by default
        var table = CreateTable([100], [40]);

        table.FirstRow.Should().BeTrue();
        table.HorizontalBanding.Should().BeTrue();
        table.VerticalBanding.Should().BeFalse();
        table.FirstCol.Should().BeFalse();
        table.LastRow.Should().BeFalse();
        table.LastCol.Should().BeFalse();
        table.RightToLeft.Should().BeFalse();
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void TableStyleOptions_SettingFalseRemovesAttribute()
    {
        var table = CreateTable([100], [40]);

        table.FirstRow.Should().BeTrue();
        table.FirstRow = false;
        table.FirstRow.Should().BeFalse();
    }

    // ── test_row_height ──────────────────────────────────────────

    /// <summary>
    /// </summary>
    [Fact]
    public void RowHeight_MatchesViaTable()
    {
        var table = CreateTable([200], [30, 50, 70]);

        table.Rows[0].Height.Should().BeApproximately(30, 0.01f);
        table.Rows[1].Height.Should().BeApproximately(50, 0.01f);
        table.Rows[2].Height.Should().BeApproximately(70, 0.01f);
    }

    // ── test_column_width ────────────────────────────────────────

    /// <summary>
    /// </summary>
    [Fact]
    public void ColumnWidth_MatchesViaTable()
    {
        var table = CreateTable([100, 200, 300], [40]);

        table.Columns[0].Width.Should().BeApproximately(100, 0.01f);
        table.Columns[1].Width.Should().BeApproximately(200, 0.01f);
        table.Columns[2].Width.Should().BeApproximately(300, 0.01f);
    }

    // ── test_cell_fill ───────────────────────────────────────────

    /// <summary>
    /// </summary>
    [Fact]
    public void CellFill_CanSetSolidFillViaTableCell()
    {
        var table = CreateTable([200], [60]);
        var cell = table.Rows[0][0];

        cell.CellFormat.FillFormat.FillType = FillType.Solid;
        cell.CellFormat.FillFormat.SolidFillColor.Color = Color.LightBlue;

        cell.CellFormat.FillFormat.FillType.Should().Be(FillType.Solid);
    }

    // ── StylePreset ──────────────────────────────────────────────

    /// <summary>
    /// StylePreset returns None when no style element is present.
    /// </summary>
    [Fact]
    public void StylePreset_DefaultIsNone()
    {
        // tblPr has no tableStyleId child
        var tbl = new XElement(ANs + "tbl",
            new XElement(ANs + "tblPr"),
            new XElement(ANs + "tblGrid"));
        var table = new Table();
        table.InitInternal(WrapInGraphicFrame(tbl), slidePart: null, parentSlide: null);

        table.StylePreset.Should().Be(TableStylePreset.None);
    }

    /// <summary>
    /// StylePreset round-trips through set/get.
    /// </summary>
    [Fact]
    public void StylePreset_RoundTrips()
    {
        var table = CreateTable([100], [40]);

        table.StylePreset = TableStylePreset.MediumStyle2Accent1;

        table.StylePreset.Should().Be(TableStylePreset.MediumStyle2Accent1);
    }

    /// <summary>
    /// Setting StylePreset to None removes the style element.
    /// </summary>
    [Fact]
    public void StylePreset_SetNoneRemovesElement()
    {
        var table = CreateTable([100], [40]);
        table.StylePreset = TableStylePreset.LightStyle1;

        table.StylePreset = TableStylePreset.None;

        table.StylePreset.Should().Be(TableStylePreset.None);
    }

    // ── AsIGraphicalObject / AsIBulkTextFormattable ──────────────

    [Fact]
    public void AsIGraphicalObject_ReturnsSelf()
    {
        var table = CreateTable([100], [40]);
        table.AsIGraphicalObject.Should().BeSameAs(table);
    }

    [Fact]
    public void AsIBulkTextFormattable_ReturnsSelf()
    {
        var table = CreateTable([100], [40]);
        table.AsIBulkTextFormattable.Should().BeSameAs(table);
    }

    // ── GraphicalObjectLock ──────────────────────────────────────

    [Fact]
    public void GraphicalObjectLock_ReturnsNull()
    {
        var table = CreateTable([100], [40]);
        table.GraphicalObjectLock.Should().BeNull();
    }

    // ── TableFormat ──────────────────────────────────────────────

    [Fact]
    public void TableFormat_IsNotNull()
    {
        var table = CreateTable([100], [40]);
        table.TableFormat.Should().NotBeNull();
    }
}
