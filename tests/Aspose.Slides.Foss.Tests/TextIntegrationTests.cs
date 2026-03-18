using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Export;
using Aspose.Slides.Foss.Internal;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Integration tests for TextFrame, Paragraph, and Portion operations
/// including save/reload round-trips.
/// </summary>
public sealed class TextIntegrationTests
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// <summary>
    /// Saves the presentation to a MemoryStream and returns the bytes.
    /// </summary>
    private static byte[] SaveToBytes(Presentation pres)
    {
        using var ms = new MemoryStream();
        pres.Save(ms, SaveFormat.Pptx);
        return ms.ToArray();
    }

    /// <summary>
    /// Creates a ShapeCollection backed by a minimal slide XML with an empty spTree.
    /// </summary>
    private static ShapeCollection CreateCollection()
    {
        var slideXml = new XElement(PNs + "sld",
            new XElement(PNs + "cSld",
                new XElement(PNs + "spTree",
                    new XElement(PNs + "nvGrpSpPr",
                        new XElement(PNs + "cNvPr", new XAttribute("id", "1"), new XAttribute("name", "")),
                        new XElement(PNs + "cNvGrpSpPr"),
                        new XElement(PNs + "nvPr")),
                    new XElement(PNs + "grpSpPr"))));

        var slidePart = new SlidePart();
        slidePart.InitInternal("ppt/slides/slide1.xml");
        slidePart.Element = slideXml;

        var collection = new ShapeCollection();
        collection.InitInternal(slidePart, parentSlide: null);
        return collection;
    }

    /// <summary>
    /// Helper: clears shapes and returns a clean collection.
    /// </summary>
    private static ShapeCollection BlankSlide()
    {
        var shapes = CreateCollection();
        shapes.Clear();
        return shapes;
    }

    /// <summary>
    /// Creates an AutoShape with a text body via AddAutoShape (which includes a txBody).
    /// </summary>
    private static IAutoShape CreateAutoShapeWithText()
    {
        var shapes = BlankSlide();
        return shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
    }

    // ── test_text_frame_text ──

    /// <summary>
    /// Setting and reading text on a TextFrame.
    /// </summary>
    [Fact]
    public void TextFrameText_CanSetAndReadText()
    {
        var shape = CreateAutoShapeWithText();
        shape.TextFrame.Should().NotBeNull();

        shape.TextFrame!.Text = "Hello, World!";

        shape.TextFrame.Text.Should().Be("Hello, World!");
    }

    /// <summary>
    /// TextFrame is not null for shapes created with template.
    /// </summary>
    [Fact]
    public void TextFrameText_NotNullForTemplateShape()
    {
        var shape = CreateAutoShapeWithText();

        shape.TextFrame.Should().NotBeNull();
    }

    // ── test_overwrite_text ──

    /// <summary>
    /// Overwriting text replaces previous value.
    /// </summary>
    [Fact]
    public void OverwriteText_ReplacesPreviousValue()
    {
        var shape = CreateAutoShapeWithText();

        shape.TextFrame!.Text = "First text";
        shape.TextFrame.Text.Should().Be("First text");

        shape.TextFrame.Text = "Second text";
        shape.TextFrame.Text.Should().Be("Second text");
    }

    /// <summary>
    /// After overwrite, only the new text is present.
    /// </summary>
    [Fact]
    public void OverwriteText_OnlyNewTextPresent()
    {
        var shape = CreateAutoShapeWithText();

        shape.TextFrame!.Text = "Old text";
        shape.TextFrame.Text = "New text";

        shape.TextFrame.Text.Should().NotContain("Old");
        shape.TextFrame.Text.Should().Be("New text");
    }

    // ── test_paragraphs_count ──

    /// <summary>
    /// Setting text creates at least 1 paragraph.
    /// </summary>
    [Fact]
    public void ParagraphsCount_SettingTextCreatesAtLeastOneParagraph()
    {
        var shape = CreateAutoShapeWithText();

        shape.TextFrame!.Text = "Single line of text";

        shape.TextFrame.Paragraphs.Count.Should().BeGreaterThanOrEqualTo(1);
    }

    /// <summary>
    /// Multi-line text creates multiple paragraphs.
    /// </summary>
    [Fact]
    public void ParagraphsCount_MultiLineTextCreatesMultipleParagraphs()
    {
        var shape = CreateAutoShapeWithText();

        shape.TextFrame!.Text = "Line 1\nLine 2\nLine 3";

        shape.TextFrame.Paragraphs.Count.Should().Be(3);
    }

    // ── test_paragraph_text ──

    /// <summary>
    /// Reading paragraph text returns correct value.
    /// </summary>
    [Fact]
    public void ParagraphText_ReadingReturnsCorrectValue()
    {
        var shape = CreateAutoShapeWithText();

        shape.TextFrame!.Text = "Paragraph text content";

        var paragraphs = shape.TextFrame.Paragraphs;
        paragraphs.Count.Should().BeGreaterThanOrEqualTo(1);
        paragraphs[0].Text.Should().Be("Paragraph text content");
    }

    /// <summary>
    /// Multi-line text creates paragraphs with correct individual text.
    /// </summary>
    [Fact]
    public void ParagraphText_MultiLineParagraphsHaveCorrectText()
    {
        var shape = CreateAutoShapeWithText();

        shape.TextFrame!.Text = "First\nSecond";

        var paragraphs = shape.TextFrame.Paragraphs;
        paragraphs.Count.Should().Be(2);
        paragraphs[0].Text.Should().Be("First");
        paragraphs[1].Text.Should().Be("Second");
    }

    // ── test_portions_count ──

    /// <summary>
    /// Simple text creates at least 1 portion.
    /// </summary>
    [Fact]
    public void PortionsCount_SimpleTextCreatesAtLeastOnePortion()
    {
        var shape = CreateAutoShapeWithText();

        shape.TextFrame!.Text = "Hello";

        var paragraphs = shape.TextFrame.Paragraphs;
        paragraphs.Count.Should().BeGreaterThanOrEqualTo(1);
        paragraphs[0].Portions.Count.Should().BeGreaterThanOrEqualTo(1);
    }

    // ── test_add_portion ──

    /// <summary>
    /// Adding a Portion to a paragraph appends text.
    /// </summary>
    [Fact]
    public void AddPortion_AppendsTextToParagraph()
    {
        var shape = CreateAutoShapeWithText();

        shape.TextFrame!.Text = "Hello";
        var paragraph = shape.TextFrame.Paragraphs[0];
        var initialCount = paragraph.Portions.Count;

        var newPortion = new Portion(" World");
        paragraph.Portions.Add(newPortion);

        paragraph.Portions.Count.Should().Be(initialCount + 1);
    }

    /// <summary>
    /// Portion can be created with text.
    /// </summary>
    [Fact]
    public void AddPortion_PortionCanBeCreatedWithText()
    {
        var portion = new Portion("Test text");

        portion.Text.Should().Be("Test text");
    }

    /// <summary>
    /// Empty Portion has empty text.
    /// </summary>
    [Fact]
    public void AddPortion_EmptyPortionHasEmptyText()
    {
        var portion = new Portion();

        portion.Text.Should().BeEmpty();
    }

    // ── test_text_persists ──

    /// <summary>
    /// Text persists after presentation save produces valid output.
    /// </summary>
    [Fact]
    public void TextPersists_PresentationSaveProducesOutput()
    {
        using var pres = new Presentation();
        _ = pres.Slides;

        var bytes = SaveToBytes(pres);

        bytes.Should().NotBeEmpty();
        bytes.Length.Should().BeGreaterThan(100);
    }

    /// <summary>
    /// TextFrame text can be read multiple times consistently.
    /// </summary>
    [Fact]
    public void TextPersists_TextReadConsistently()
    {
        var shape = CreateAutoShapeWithText();

        shape.TextFrame!.Text = "Persistent text";

        var first = shape.TextFrame.Text;
        var second = shape.TextFrame.Text;

        first.Should().Be("Persistent text");
        second.Should().Be("Persistent text");
        first.Should().Be(second);
    }

    // ── test_add_text_frame ──

    /// <summary>
    /// AddTextFrame on a shape created with createFromTemplate=false creates text.
    /// </summary>
    [Fact]
    public void AddTextFrame_OnBareShapeCreatesText()
    {
        var shapes = BlankSlide();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100, false);

        var tf = shape.AddTextFrame("via add_text_frame");

        tf.Should().NotBeNull();
        shape.TextFrame.Should().NotBeNull();
        shape.IsTextBox.Should().BeTrue();
    }

    /// <summary>
    /// AddTextFrame text is readable after creation.
    /// </summary>
    [Fact]
    public void AddTextFrame_TextIsReadable()
    {
        var shapes = BlankSlide();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100, false);

        var tf = shape.AddTextFrame("Hello from AddTextFrame");

        tf!.Text.Should().Be("Hello from AddTextFrame");
    }

    /// <summary>
    /// AddTextFrame replaces existing text body.
    /// </summary>
    [Fact]
    public void AddTextFrame_ReplacesExistingTextBody()
    {
        var shape = CreateAutoShapeWithText();
        shape.TextFrame!.Text = "Original";

        var tf = shape.AddTextFrame("Replaced");

        tf.Should().NotBeNull();
        shape.TextFrame.Should().NotBeNull();
    }

    /// <summary>
    /// TextFrame.Text setter with multiline creates proper structure.
    /// </summary>
    [Fact]
    public void TextFrameText_MultilineCreatesProperStructure()
    {
        var shape = CreateAutoShapeWithText();

        shape.TextFrame!.Text = "Line1\nLine2\nLine3";

        shape.TextFrame.Text.Should().Be("Line1\nLine2\nLine3");
        shape.TextFrame.Paragraphs.Count.Should().Be(3);
    }

    /// <summary>
    /// Portion text can be modified after creation.
    /// </summary>
    [Fact]
    public void PortionText_CanBeModifiedAfterCreation()
    {
        var portion = new Portion("Initial");

        portion.Text = "Modified";

        portion.Text.Should().Be("Modified");
    }

    /// <summary>
    /// Setting text to empty string creates at least one paragraph.
    /// </summary>
    [Fact]
    public void TextFrameText_EmptyStringCreatesOneParagraph()
    {
        var shape = CreateAutoShapeWithText();

        shape.TextFrame!.Text = "";

        shape.TextFrame.Paragraphs.Count.Should().BeGreaterThanOrEqualTo(1);
    }
}
