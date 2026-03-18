using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using FluentAssertions;

namespace Aspose.Slides.Foss.IntegrationTests;

/// <summary>
/// Tests for FillFormat: solid, gradient, pattern, picture, no-fill.
/// </summary>
public sealed class FillFormatTests : IDisposable
{
    private readonly string _tempDir;

    public FillFormatTests()
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
    public void TestSolidFill()
    {
        // Solid fill colour persists after save/reload.
        using var pres = new Presentation();
        var slide = ClearSlide(pres);
        var shape = slide.Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        shape.FillFormat.FillType = FillType.Solid;
        shape.FillFormat.SolidFillColor.Color = Color.FromArgb(255, 0, 128, 255);

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var ff = pres2.Slides[0].Shapes![0].FillFormat;
        ff.FillType.Should().Be(FillType.Solid);
        var c = ff.SolidFillColor.Color;
        c!.R.Should().Be(0);
        c.G.Should().Be(128);
        c.B.Should().Be(255);
    }

    [Fact]
    public void TestGradientFill()
    {
        // Gradient stops and angle persist.
        using var pres = new Presentation();
        var slide = ClearSlide(pres);
        var shape = slide.Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 300, 150);
        shape.FillFormat.FillType = FillType.Gradient;
        var gf = shape.FillFormat.GradientFormat;
        gf.GradientShape = GradientShape.Linear;
        gf.LinearGradientAngle = 45;
        gf.GradientStops.Add(0.0f, Color.Blue);
        gf.GradientStops.Add(1.0f, Color.Red);

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var ff2 = pres2.Slides[0].Shapes![0].FillFormat;
        ff2.FillType.Should().Be(FillType.Gradient);
        ff2.GradientFormat.GradientStops.Count.Should().BeGreaterThanOrEqualTo(2);
    }

    [Fact]
    public void TestPatternFill()
    {
        // Pattern style and colours persist.
        using var pres = new Presentation();
        var slide = ClearSlide(pres);
        var shape = slide.Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        shape.FillFormat.FillType = FillType.Pattern;
        var pf = shape.FillFormat.PatternFormat;
        pf.PatternStyle = PatternStyle.Percent50;
        pf.ForeColor.Color = Color.DarkBlue;
        pf.BackColor.Color = Color.LightYellow;

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var ff2 = pres2.Slides[0].Shapes![0].FillFormat;
        ff2.FillType.Should().Be(FillType.Pattern);
        ff2.PatternFormat.PatternStyle.Should().Be(PatternStyle.Percent50);
    }

    [Fact]
    public void TestNoFill()
    {
        // NO_FILL type persists.
        using var pres = new Presentation();
        var slide = ClearSlide(pres);
        var shape = slide.Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        shape.FillFormat.FillType = FillType.NoFill;

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        pres2.Slides[0].Shapes![0].FillFormat.FillType.Should().Be(FillType.NoFill);
    }

    [Fact]
    public void TestPictureFill()
    {
        // Picture fill with an image persists.
        using var pres = new Presentation();
        var slide = ClearSlide(pres);
        var shape = slide.Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 200);
        shape.FillFormat.FillType = FillType.Picture;
        var pff = shape.FillFormat.PictureFillFormat;
        pff.PictureFillMode = PictureFillMode.Stretch;
        var img = pres.Images.AddImage(TestHelpers.CreateTestPng(0, 255, 0));
        pff.Picture.Image = img;

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var ff2 = pres2.Slides[0].Shapes![0].FillFormat;
        ff2.FillType.Should().Be(FillType.Picture);
    }
}
