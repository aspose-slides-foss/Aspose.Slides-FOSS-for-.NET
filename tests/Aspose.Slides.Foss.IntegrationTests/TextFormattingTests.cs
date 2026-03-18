using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using FluentAssertions;

namespace Aspose.Slides.Foss.IntegrationTests;

/// <summary>
/// Tests for text formatting: bold, italic, underline, font, colour, alignment.
/// </summary>
public sealed class TextFormattingTests : IDisposable
{
    private readonly string _tempDir;

    public TextFormattingTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, recursive: true);
    }

    /// <summary>
    /// Helper: clear slide, add a rectangle with text and return (shape, portion_format).
    /// </summary>
    private static (IAutoShape Shape, IBasePortionFormat Format) Shaped(Presentation pres)
    {
        pres.Slides[0].Shapes!.Clear();
        var shape = pres.Slides[0].Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 400, 60);
        shape.TextFrame!.Text = "Sample";
        var fmt = shape.TextFrame!.Paragraphs[0].Portions[0].PortionFormat!;
        return (shape, fmt);
    }

    [Fact]
    public void TestBoldItalic()
    {
        // Bold and italic persist after save/reload.
        using var pres = new Presentation();
        var (_, fmt) = Shaped(pres);
        fmt.FontBold = NullableBool.True;
        fmt.FontItalic = NullableBool.True;

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var fmt2 = ((IAutoShape)pres2.Slides[0].Shapes![0]).TextFrame!.Paragraphs[0].Portions[0].PortionFormat!;
        fmt2.FontBold.Should().Be(NullableBool.True);
        fmt2.FontItalic.Should().Be(NullableBool.True);
    }

    [Fact]
    public void TestUnderline()
    {
        // Underline type persists.
        using var pres = new Presentation();
        var (_, fmt) = Shaped(pres);
        fmt.FontUnderline = TextUnderlineType.Single;

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var fmt2 = ((IAutoShape)pres2.Slides[0].Shapes![0]).TextFrame!.Paragraphs[0].Portions[0].PortionFormat!;
        fmt2.FontUnderline.Should().Be(TextUnderlineType.Single);
    }

    [Fact]
    public void TestStrikethrough()
    {
        // Strikethrough type persists.
        using var pres = new Presentation();
        var (_, fmt) = Shaped(pres);
        fmt.StrikethroughType = TextStrikethroughType.Single;

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var fmt2 = ((IAutoShape)pres2.Slides[0].Shapes![0]).TextFrame!.Paragraphs[0].Portions[0].PortionFormat!;
        fmt2.StrikethroughType.Should().Be(TextStrikethroughType.Single);
    }

    [Fact]
    public void TestFontSize()
    {
        // font_height persists.
        using var pres = new Presentation();
        var (_, fmt) = Shaped(pres);
        fmt.FontHeight = 28;

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var fmt2 = ((IAutoShape)pres2.Slides[0].Shapes![0]).TextFrame!.Paragraphs[0].Portions[0].PortionFormat!;
        fmt2.FontHeight.Should().Be(28);
    }

    [Fact]
    public void TestFontColor()
    {
        // Solid fill colour on portion text persists.
        using var pres = new Presentation();
        var (_, fmt) = Shaped(pres);
        fmt.FillFormat!.FillType = FillType.Solid;
        fmt.FillFormat!.SolidFillColor.Color = Color.Red;

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var fmt2 = ((IAutoShape)pres2.Slides[0].Shapes![0]).TextFrame!.Paragraphs[0].Portions[0].PortionFormat!;
        fmt2.FillFormat!.FillType.Should().Be(FillType.Solid);
        var c = fmt2.FillFormat!.SolidFillColor.Color;
        c!.R.Should().Be(255);
        c.G.Should().Be(0);
        c.B.Should().Be(0);
    }

    [Fact]
    public void TestLatinFont()
    {
        // latin_font persists.
        using var pres = new Presentation();
        var (_, fmt) = Shaped(pres);
        fmt.LatinFont = new FontData("Courier New");

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var fmt2 = ((IAutoShape)pres2.Slides[0].Shapes![0]).TextFrame!.Paragraphs[0].Portions[0].PortionFormat!;
        fmt2.LatinFont!.FontName.Should().Be("Courier New");
    }

    [Fact]
    public void TestParagraphAlignment()
    {
        // Paragraph alignment persists.
        using var pres = new Presentation();
        pres.Slides[0].Shapes!.Clear();
        var shape = pres.Slides[0].Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 400, 200);
        shape.TextFrame!.Text = "Centered";
        shape.TextFrame!.Paragraphs[0].ParagraphFormat.Alignment = TextAlignment.Center;

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var pf = ((IAutoShape)pres2.Slides[0].Shapes![0]).TextFrame!.Paragraphs[0].ParagraphFormat;
        pf.Alignment.Should().Be(TextAlignment.Center);
    }
}
