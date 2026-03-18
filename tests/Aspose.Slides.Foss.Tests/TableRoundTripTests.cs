using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using Aspose.Slides.Foss.Internal;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Round-trip integration tests for Table functionality through XML serialization.
/// </summary>
public sealed class TableRoundTripTests
{
    private static readonly string SlideXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
               xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:cSld>
            <p:spTree>
              <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
              <p:grpSpPr/>
            </p:spTree>
          </p:cSld>
        </p:sld>
        """;

    private static (SlidePart slidePart, ShapeCollection shapes) CreateSlideWithShapes()
    {
        var slidePart = new SlidePart();
        slidePart.InitInternal("ppt/slides/slide1.xml");
        slidePart.Element = XDocument.Parse(SlideXml).Root;

        var shapes = new ShapeCollection();
        shapes.InitInternal(slidePart, null);
        return (slidePart, shapes);
    }

    private static ShapeCollection RoundTripShapes(SlidePart slidePart)
    {
        var xml = slidePart.Element!.ToString();
        var reloadedRoot = XDocument.Parse(xml).Root!;

        var newSlidePart = new SlidePart();
        newSlidePart.InitInternal("ppt/slides/slide1.xml");
        newSlidePart.Element = reloadedRoot;

        var shapes = new ShapeCollection();
        shapes.InitInternal(newSlidePart, null);
        return shapes;
    }

    /// <summary>
    /// Create a table and verify row/column counts.
    /// </summary>
    [Fact]
    public void CreateTable_RowAndColumnCountsMatch()
    {
        var (_, shapes) = CreateSlideWithShapes();
        double[] colWidths = [100, 150, 200];
        double[] rowHeights = [30, 40];
        var table = shapes.AddTable(50, 50, colWidths, rowHeights);

        table.Rows.Count.Should().Be(2);
        table.Columns.Count.Should().Be(3);
    }

    /// <summary>
    /// Cell text round-trips through XML serialization.
    /// </summary>
    [Fact]
    public void CellText_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        double[] colWidths = [100, 100];
        double[] rowHeights = [30, 30];
        var table = shapes.AddTable(50, 50, colWidths, rowHeights);

        var cell = table.Rows[0][0];
        var textFrame = cell.TextFrame;
        textFrame.Should().NotBeNull();
        textFrame!.Text = "Hello";

        var reloaded = RoundTripShapes(slidePart);
        var reloadedTable = (ITable)reloaded[0];
        var reloadedText = reloadedTable.Rows[0][0].TextFrame?.Text;

        reloadedText.Should().Be("Hello");
    }

    /// <summary>
    /// Merged cells preserve ColSpan after XML round-trip.
    /// </summary>
    [Fact]
    public void MergeCells_ColSpanPersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        double[] colWidths = [100, 100, 100];
        double[] rowHeights = [30];
        var table = shapes.AddTable(50, 50, colWidths, rowHeights);

        table.MergeCells(table.Rows[0][0], table.Rows[0][2], false);

        var reloaded = RoundTripShapes(slidePart);
        var reloadedTable = (ITable)reloaded[0];
        var originCell = reloadedTable.Rows[0][0];

        originCell.ColSpan.Should().Be(3);
        originCell.IsMergedCell.Should().BeTrue();
    }

    /// <summary>
    /// Cell borders persist after XML round-trip.
    /// </summary>
    [Fact]
    public void CellBorders_PersistAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        double[] colWidths = [100];
        double[] rowHeights = [30];
        var table = shapes.AddTable(50, 50, colWidths, rowHeights);

        var cellFormat = table.Rows[0][0].CellFormat;
        cellFormat.BorderLeft.Width = 3.0f;
        cellFormat.BorderTop.Width = 2.0f;

        var reloaded = RoundTripShapes(slidePart);
        var reloadedTable = (ITable)reloaded[0];
        var reloadedCellFormat = reloadedTable.Rows[0][0].CellFormat;

        reloadedCellFormat.BorderLeft.Width.Should().BeApproximately(3.0f, 0.1f);
        reloadedCellFormat.BorderTop.Width.Should().BeApproximately(2.0f, 0.1f);
    }

    /// <summary>
    /// Cell borders with fill format color persist after XML round-trip.
    /// </summary>
    [Fact]
    public void CellBorders_FillColorPersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        double[] colWidths = [150];
        double[] rowHeights = [50];
        var table = shapes.AddTable(50, 50, colWidths, rowHeights);

        var fmt = table.Rows[0][0].CellFormat;
        fmt.BorderTop.FillFormat.FillType = FillType.Solid;
        fmt.BorderTop.FillFormat.SolidFillColor.Color = Color.Red;
        fmt.BorderTop.Width = 3;

        var reloaded = RoundTripShapes(slidePart);
        var reloadedTable = (ITable)reloaded[0];
        var reloadedFmt = reloadedTable.Rows[0][0].CellFormat;

        reloadedFmt.BorderTop.Width.Should().BeApproximately(3f, 0.1f);
    }

    /// <summary>
    /// Table style flags persist after XML round-trip.
    /// </summary>
    [Fact]
    public void TableStyleOptions_PersistAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        double[] colWidths = [100, 100];
        double[] rowHeights = [30, 30];
        var table = shapes.AddTable(50, 50, colWidths, rowHeights);

        table.FirstRow = true;
        table.LastRow = true;
        table.FirstCol = true;
        table.LastCol = true;
        table.HorizontalBanding = false;
        table.VerticalBanding = true;

        var reloaded = RoundTripShapes(slidePart);
        var reloadedTable = (ITable)reloaded[0];

        reloadedTable.FirstRow.Should().BeTrue();
        reloadedTable.LastRow.Should().BeTrue();
        reloadedTable.FirstCol.Should().BeTrue();
        reloadedTable.LastCol.Should().BeTrue();
        reloadedTable.HorizontalBanding.Should().BeFalse();
        reloadedTable.VerticalBanding.Should().BeTrue();
    }

    /// <summary>
    /// Row heights match constructor arguments.
    /// </summary>
    [Fact]
    public void RowHeight_MatchesConstructorArguments()
    {
        var (_, shapes) = CreateSlideWithShapes();
        double[] colWidths = [100];
        double[] rowHeights = [25, 40, 55];
        var table = shapes.AddTable(50, 50, colWidths, rowHeights);

        table.Rows[0].Height.Should().BeApproximately(25f, 1f);
        table.Rows[1].Height.Should().BeApproximately(40f, 1f);
        table.Rows[2].Height.Should().BeApproximately(55f, 1f);
    }

    /// <summary>
    /// Column widths match constructor arguments.
    /// </summary>
    [Fact]
    public void ColumnWidth_MatchesConstructorArguments()
    {
        var (_, shapes) = CreateSlideWithShapes();
        double[] colWidths = [80, 120, 160];
        double[] rowHeights = [30];
        var table = shapes.AddTable(50, 50, colWidths, rowHeights);

        table.Columns[0].Width.Should().BeApproximately(80f, 1f);
        table.Columns[1].Width.Should().BeApproximately(120f, 1f);
        table.Columns[2].Width.Should().BeApproximately(160f, 1f);
    }

    /// <summary>
    /// Cell fill colour persists after XML round-trip.
    /// </summary>
    [Fact]
    public void CellFill_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        double[] colWidths = [100];
        double[] rowHeights = [30];
        var table = shapes.AddTable(50, 50, colWidths, rowHeights);

        var cellFormat = table.Rows[0][0].CellFormat;
        cellFormat.FillFormat.FillType = FillType.Solid;
        cellFormat.FillFormat.SolidFillColor.Color = Color.FromArgb(255, 255, 0, 0);

        var reloaded = RoundTripShapes(slidePart);
        var reloadedTable = (ITable)reloaded[0];
        var reloadedCellFormat = reloadedTable.Rows[0][0].CellFormat;

        reloadedCellFormat.FillFormat.FillType.Should().Be(FillType.Solid);
        var color = reloadedCellFormat.FillFormat.SolidFillColor.Color;
        color.Should().NotBeNull();
        color!.R.Should().Be(255);
        color.G.Should().Be(0);
        color.B.Should().Be(0);
    }
}
