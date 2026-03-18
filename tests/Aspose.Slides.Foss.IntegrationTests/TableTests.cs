using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using FluentAssertions;

namespace Aspose.Slides.Foss.IntegrationTests;

/// <summary>
/// Tests for Table: create, cell text, merge, borders, style options.
/// </summary>
public sealed class TableTests : IDisposable
{
    private readonly string _tempDir;

    public TableTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, recursive: true);
    }

    /// <summary>Find the Table shape on a slide (skip placeholders).</summary>
    private static ITable? FindTable(ISlide slide)
    {
        foreach (var shape in slide.Shapes!)
        {
            if (shape is Table table)
                return table;
        }
        return null;
    }

    /// <summary>Return the first slide with placeholders removed.</summary>
    private static ISlide BlankSlide(Presentation pres)
    {
        var slide = pres.Slides[0];
        slide.Shapes!.Clear();
        return slide;
    }

    [Fact]
    public void TestCreateTable()
    {
        // Create a table and verify row/column counts.
        using var pres = new Presentation();
        var slide = BlankSlide(pres);
        var table = slide.Shapes!.AddTable(50, 50, [100, 150, 200], [40, 40, 40]);
        table.Rows.Count.Should().Be(3);
        table.Columns.Count.Should().Be(3);
    }

    [Fact]
    public void TestCellText()
    {
        // Cell text round-trips through save/reload.
        using var pres = new Presentation();
        var slide = BlankSlide(pres);
        var table = slide.Shapes!.AddTable(50, 50, [100, 100], [40, 40]);
        table.Rows[0][0].TextFrame!.Text = "A";
        table.Rows[0][1].TextFrame!.Text = "B";
        table.Rows[1][0].TextFrame!.Text = "C";
        table.Rows[1][1].TextFrame!.Text = "D";

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var t2 = FindTable(pres2.Slides[0]);
        t2.Should().NotBeNull();
        t2!.Rows[0][0].TextFrame!.Text.Should().Be("A");
        t2.Rows[0][1].TextFrame!.Text.Should().Be("B");
        t2.Rows[1][0].TextFrame!.Text.Should().Be("C");
        t2.Rows[1][1].TextFrame!.Text.Should().Be("D");
    }

    [Fact]
    public void TestMergeCells()
    {
        // Merged cells preserve col_span after reload.
        using var pres = new Presentation();
        var slide = BlankSlide(pres);
        var table = slide.Shapes!.AddTable(50, 50, [100, 100, 100], [40, 40]);
        var cell1 = table.Rows[0][0];
        var cell2 = table.Rows[0][1];
        table.MergeCells(cell1, cell2, false);
        cell1.IsMergedCell.Should().BeTrue();
        cell1.ColSpan.Should().BeGreaterThanOrEqualTo(2);

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var t2 = FindTable(pres2.Slides[0]);
        t2.Should().NotBeNull();
        t2!.Rows[0][0].IsMergedCell.Should().BeTrue();
    }

    [Fact]
    public void TestCellBorders()
    {
        // Cell borders persist.
        using var pres = new Presentation();
        var slide = BlankSlide(pres);
        var table = slide.Shapes!.AddTable(50, 50, [150], [50]);
        var cell = table.Rows[0][0];
        cell.TextFrame!.Text = "Bordered";
        var fmt = cell.CellFormat;
        foreach (var border in new[] { fmt.BorderTop, fmt.BorderBottom, fmt.BorderLeft, fmt.BorderRight })
        {
            border.FillFormat.FillType = FillType.Solid;
            border.FillFormat.SolidFillColor.Color = Color.Red;
            border.Width = 3;
        }

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var t2 = FindTable(pres2.Slides[0]);
        t2.Should().NotBeNull();
        var fmt2 = t2!.Rows[0][0].CellFormat;
        fmt2.BorderTop.Width.Should().Be(3);
    }

    [Fact]
    public void TestTableStyleOptions()
    {
        // Table style flags persist.
        using var pres = new Presentation();
        var slide = BlankSlide(pres);
        var table = slide.Shapes!.AddTable(50, 50, [120, 120], [40, 40, 40]);
        table.FirstRow = true;
        table.HorizontalBanding = true;
        table.VerticalBanding = false;

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var t2 = FindTable(pres2.Slides[0]);
        t2.Should().NotBeNull();
        t2!.FirstRow.Should().BeTrue();
        t2.HorizontalBanding.Should().BeTrue();
    }

    [Fact]
    public void TestRowHeight()
    {
        // Row heights match constructor arguments.
        using var pres = new Presentation();
        var slide = BlankSlide(pres);
        var table = slide.Shapes!.AddTable(50, 50, [200], [30, 50, 70]);
        table.Rows[0].Height.Should().Be(30);
        table.Rows[1].Height.Should().Be(50);
        table.Rows[2].Height.Should().Be(70);
    }

    [Fact]
    public void TestColumnWidth()
    {
        // Column widths match constructor arguments.
        using var pres = new Presentation();
        var slide = BlankSlide(pres);
        var table = slide.Shapes!.AddTable(50, 50, [100, 200, 300], [40]);
        table.Columns[0].Width.Should().Be(100);
        table.Columns[1].Width.Should().Be(200);
        table.Columns[2].Width.Should().Be(300);
    }

    [Fact]
    public void TestCellFill()
    {
        // Cell fill colour persists.
        using var pres = new Presentation();
        var slide = BlankSlide(pres);
        var table = slide.Shapes!.AddTable(50, 50, [200], [60]);
        var cell = table.Rows[0][0];
        cell.CellFormat.FillFormat.FillType = FillType.Solid;
        cell.CellFormat.FillFormat.SolidFillColor.Color = Color.LightBlue;
        cell.TextFrame!.Text = "Blue";

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var t2 = FindTable(pres2.Slides[0]);
        t2.Should().NotBeNull();
        var cf2 = t2!.Rows[0][0].CellFormat;
        cf2.FillFormat.FillType.Should().Be(FillType.Solid);
    }
}
