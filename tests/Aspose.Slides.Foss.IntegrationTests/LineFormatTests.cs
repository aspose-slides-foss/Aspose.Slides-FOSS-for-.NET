using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using FluentAssertions;

namespace Aspose.Slides.Foss.IntegrationTests;

/// <summary>
/// Tests for LineFormat: colour, width, dash style.
/// </summary>
public sealed class LineFormatTests : IDisposable
{
    private readonly string _tempDir;

    public LineFormatTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, recursive: true);
    }

    private static ISlide ClearSlide(Presentation pres)
    {
        pres.Slides[0].Shapes!.Clear();
        return pres.Slides[0];
    }

    [Fact]
    public void TestLineColorAndWidth()
    {
        // Line colour and width persist after save/reload.
        using var pres = new Presentation();
        var slide = ClearSlide(pres);
        var shape = slide.Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        var lf = shape.LineFormat;
        lf.Width = 5;
        lf.FillFormat.FillType = FillType.Solid;
        lf.FillFormat.SolidFillColor.Color = Color.DarkRed;

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var lf2 = pres2.Slides[0].Shapes![0].LineFormat;
        lf2.Width.Should().Be(5);
        lf2.FillFormat.FillType.Should().Be(FillType.Solid);
        var c = lf2.FillFormat.SolidFillColor.Color;
        c!.R.Should().Be(Color.DarkRed.R);
    }

    [Fact]
    public void TestLineDashStyle()
    {
        // Dash style persists.
        using var pres = new Presentation();
        var slide = ClearSlide(pres);
        var shape = slide.Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        var lf = shape.LineFormat;
        lf.Width = 3;
        lf.DashStyle = LineDashStyle.Dash;
        lf.FillFormat.FillType = FillType.Solid;
        lf.FillFormat.SolidFillColor.Color = Color.Black;

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var lf2 = pres2.Slides[0].Shapes![0].LineFormat;
        lf2.DashStyle.Should().Be(LineDashStyle.Dash);
    }

    [Fact]
    public void TestMultipleDashStyles()
    {
        // Various dash styles can be set in-memory.
        LineDashStyle[] styles =
        [
            LineDashStyle.Solid,
            LineDashStyle.Dash,
            LineDashStyle.Dot,
            LineDashStyle.DashDot,
        ];

        using var pres = new Presentation();
        var slide = pres.Slides[0];
        foreach (var style in styles)
        {
            var shape = slide.Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 50);
            shape.LineFormat.DashStyle = style;
            shape.LineFormat.DashStyle.Should().Be(style);
        }
    }
}
