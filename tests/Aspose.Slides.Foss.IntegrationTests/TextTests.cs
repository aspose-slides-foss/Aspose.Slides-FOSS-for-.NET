using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.IntegrationTests;

/// <summary>
/// Tests for TextFrame, Paragraph, and Portion CRUD.
/// </summary>
public sealed class TextTests : IDisposable
{
    private readonly string _tempDir;

    public TextTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, recursive: true);
    }

    [Fact]
    public void TestTextFrameText()
    {
        // Setting text_frame.text and reading it back.
        using var pres = new Presentation();
        var shape = pres.Slides[0].Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 300, 100);
        shape.TextFrame!.Text = "Hello, World!";
        shape.TextFrame!.Text.Should().Be("Hello, World!");
    }

    [Fact]
    public void TestOverwriteText()
    {
        // Overwriting text replaces the previous value.
        using var pres = new Presentation();
        var shape = pres.Slides[0].Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 300, 100);
        shape.TextFrame!.Text = "First";
        shape.TextFrame!.Text = "Second";
        shape.TextFrame!.Text.Should().Be("Second");
    }

    [Fact]
    public void TestParagraphsCount()
    {
        // Setting text creates exactly one paragraph.
        using var pres = new Presentation();
        var shape = pres.Slides[0].Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 300, 100);
        shape.TextFrame!.Text = "Line";
        shape.TextFrame!.Paragraphs.Count.Should().BeGreaterThanOrEqualTo(1);
    }

    [Fact]
    public void TestParagraphText()
    {
        // Reading and modifying paragraph text.
        using var pres = new Presentation();
        var shape = pres.Slides[0].Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 300, 100);
        shape.TextFrame!.Text = "Original";
        var para = shape.TextFrame!.Paragraphs[0];
        para.Text.Should().Be("Original");
        para.Text = "Modified";
        para.Text.Should().Be("Modified");
    }

    [Fact]
    public void TestPortionsCount()
    {
        // A simple text creates at least one portion.
        using var pres = new Presentation();
        var shape = pres.Slides[0].Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 300, 100);
        shape.TextFrame!.Text = "Hello";
        shape.TextFrame!.Paragraphs[0].Portions.Count.Should().BeGreaterThanOrEqualTo(1);
    }

    [Fact]
    public void TestAddPortion()
    {
        // Adding a Portion appends text.
        using var pres = new Presentation();
        var shape = pres.Slides[0].Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 400, 100);
        shape.TextFrame!.Text = "Hello ";
        var newPortion = new Portion("World!");
        shape.TextFrame!.Paragraphs[0].Portions.Add(newPortion);
        shape.TextFrame!.Text.Should().Contain("World!");
    }

    [Fact]
    public void TestTextPersists()
    {
        // Text survives a save/reload cycle.
        using var pres = new Presentation();
        pres.Slides[0].Shapes!.Clear();
        var shape = pres.Slides[0].Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 300, 100);
        shape.TextFrame!.Text = "Persistent text";

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        ((IAutoShape)pres2.Slides[0].Shapes![0]).TextFrame!.Text.Should().Be("Persistent text");
    }

    [Fact]
    public void TestAddTextFrame()
    {
        // add_text_frame on a shape created without text.
        using var pres = new Presentation();
        var shape = pres.Slides[0].Shapes!.AddAutoShape(
            ShapeType.Rectangle, 50, 50, 300, 100, false);
        shape.AddTextFrame("via add_text_frame");
        shape.TextFrame!.Text.Should().Be("via add_text_frame");
    }
}
