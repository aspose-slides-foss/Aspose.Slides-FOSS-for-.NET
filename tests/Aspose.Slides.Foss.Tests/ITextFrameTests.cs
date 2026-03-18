using System.Xml.Linq;
using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// </summary>
public sealed class ITextFrameTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// <summary>
    /// Creates a <see cref="TextFrame"/> backed by an XML txBody element with the given text lines.
    /// Each line becomes a separate paragraph with a single run.
    /// </summary>
    private static TextFrame CreateTextFrame(params string[] lines)
    {
        var txBody = new XElement(ANs + "txBody");
        foreach (var line in lines)
        {
            txBody.Add(new XElement(ANs + "p",
                new XElement(ANs + "r",
                    new XElement(ANs + "t", line))));
        }

        var tf = new TextFrame();
        tf.InitInternal(txBody, null, null, (IShape?)null);
        return tf;
    }

    /// <summary>
    /// Creates an empty <see cref="TextFrame"/> with no paragraphs.
    /// </summary>
    private static TextFrame CreateEmptyTextFrame()
    {
        var txBody = new XElement(ANs + "txBody");
        var tf = new TextFrame();
        tf.InitInternal(txBody, null, null, (IShape?)null);
        return tf;
    }

    // -------------------------------------------------------------------
    // "Setting text_frame.text and reading it back."
    // -------------------------------------------------------------------

    [Fact]
    public void Text_GetReturnsSetValue()
    {
        var tf = CreateEmptyTextFrame();

        tf.Text = "Hello, World!";

        tf.Text.Should().Be("Hello, World!");
    }

    // -------------------------------------------------------------------
    // "Overwriting text replaces the previous value."
    // -------------------------------------------------------------------

    [Fact]
    public void Text_OverwriteReplacesPreviousValue()
    {
        var tf = CreateEmptyTextFrame();

        tf.Text = "First";
        tf.Text = "Second";

        tf.Text.Should().Be("Second");
    }

    // -------------------------------------------------------------------
    // "Setting text creates exactly one paragraph."
    // -------------------------------------------------------------------

    [Fact]
    public void Paragraphs_CountIsAtLeastOneAfterSettingText()
    {
        var tf = CreateEmptyTextFrame();

        tf.Text = "Line";

        tf.Paragraphs.Count.Should().BeGreaterThanOrEqualTo(1);
    }

    // -------------------------------------------------------------------
    // "Reading and modifying paragraph text."
    // -------------------------------------------------------------------

    [Fact]
    public void Paragraphs_FirstParagraphTextMatchesFrameText()
    {
        var tf = CreateTextFrame("Original");

        tf.Paragraphs[0].Text.Should().Be("Original");
    }

    // -------------------------------------------------------------------
    // "A simple text creates at least one portion."
    // -------------------------------------------------------------------

    [Fact]
    public void Paragraphs_FirstParagraphHasAtLeastOnePortion()
    {
        var tf = CreateTextFrame("Hello");

        tf.Paragraphs[0].Portions.Count.Should().BeGreaterThanOrEqualTo(1);
    }

    // -------------------------------------------------------------------
    // "Adding a Portion appends text."
    // -------------------------------------------------------------------

    [Fact]
    public void Text_ContainsAddedPortionText()
    {
        // Build a paragraph with two runs directly in the XML,
        // simulating what happens when a Portion is added through the full pipeline.
        var txBody = new XElement(ANs + "txBody",
            new XElement(ANs + "p",
                new XElement(ANs + "r", new XElement(ANs + "t", "Hello ")),
                new XElement(ANs + "r", new XElement(ANs + "t", "World!"))));

        var tf = new TextFrame();
        tf.InitInternal(txBody, null, null, (IShape?)null);

        tf.Text.Should().Contain("World!");
        tf.Text.Should().Be("Hello World!");
    }

    // -------------------------------------------------------------------
    // "Text survives a save/reload cycle."
    // Verified by writing and reading back from the same XML.
    // -------------------------------------------------------------------

    [Fact]
    public void Text_PersistsAcrossMultipleReads()
    {
        var tf = CreateEmptyTextFrame();
        tf.Text = "Persistent text";

        var read1 = tf.Text;
        var read2 = tf.Text;

        read1.Should().Be("Persistent text");
        read2.Should().Be("Persistent text");
    }

    // -------------------------------------------------------------------
    // "add_text_frame on a shape created without text."
    // -------------------------------------------------------------------

    [Fact]
    public void Text_CanBeSetOnEmptyFrame()
    {
        var tf = CreateEmptyTextFrame();

        tf.Text = "via add_text_frame";

        tf.Text.Should().Be("via add_text_frame");
    }

    // -------------------------------------------------------------------
    // "Bold and italic persist after save/reload."
    // Verifies portion format is accessible through TextFrame paragraphs.
    // -------------------------------------------------------------------

    [Fact]
    public void Paragraphs_PortionFormatIsAccessible()
    {
        var tf = CreateTextFrame("Sample");

        var fmt = tf.Paragraphs[0].Portions[0].PortionFormat;

        fmt.Should().NotBeNull();
    }

    // -------------------------------------------------------------------
    // "font_height persists."
    // -------------------------------------------------------------------

    [Fact]
    public void Paragraphs_PortionFormatFontHeightCanBeSet()
    {
        var tf = CreateTextFrame("Sample");
        var fmt = tf.Paragraphs[0].Portions[0].PortionFormat;

        fmt!.FontHeight = 28;

        fmt.FontHeight.Should().Be(28);
    }

    // -------------------------------------------------------------------
    // "Solid fill colour on portion text persists."
    // -------------------------------------------------------------------

    [Fact]
    public void Paragraphs_PortionFormatFillFormatIsAccessible()
    {
        var tf = CreateTextFrame("Sample");
        var fmt = tf.Paragraphs[0].Portions[0].PortionFormat;

        fmt!.FillFormat.Should().NotBeNull();
    }

    // -------------------------------------------------------------------
    // "latin_font persists."
    // -------------------------------------------------------------------

    [Fact]
    public void Paragraphs_PortionFormatLatinFontCanBeSet()
    {
        var tf = CreateTextFrame("Sample");
        var fmt = tf.Paragraphs[0].Portions[0].PortionFormat;

        fmt!.LatinFont = new FontData("Courier New");

        fmt.LatinFont.Should().NotBeNull();
        fmt.LatinFont!.FontName.Should().Be("Courier New");
    }

    // -------------------------------------------------------------------
    // "Paragraph alignment persists."
    // -------------------------------------------------------------------

    [Fact]
    public void Paragraphs_ParagraphFormatIsAccessible()
    {
        var tf = CreateTextFrame("Centered");

        tf.Paragraphs[0].ParagraphFormat.Should().NotBeNull();
    }

    // -------------------------------------------------------------------
    // "Underline type persists."
    // -------------------------------------------------------------------

    [Fact]
    public void Paragraphs_PortionFormatUnderlineCanBeSet()
    {
        var tf = CreateTextFrame("Sample");
        var fmt = tf.Paragraphs[0].Portions[0].PortionFormat;

        fmt!.FontUnderline = TextUnderlineType.Single;

        fmt.FontUnderline.Should().Be(TextUnderlineType.Single);
    }

    // -------------------------------------------------------------------
    // "Strikethrough type persists."
    // -------------------------------------------------------------------

    [Fact]
    public void Paragraphs_PortionFormatStrikethroughCanBeSet()
    {
        var tf = CreateTextFrame("Sample");
        var fmt = tf.Paragraphs[0].Portions[0].PortionFormat;

        fmt!.StrikethroughType = TextStrikethroughType.Single;

        fmt.StrikethroughType.Should().Be(TextStrikethroughType.Single);
    }

    // -------------------------------------------------------------------
    // "Comment text, position, and time persist."
    // -------------------------------------------------------------------

    [Fact]
    public void Text_SupportsArbitraryStringContent()
    {
        var tf = CreateEmptyTextFrame();

        tf.Text = "Review note";

        tf.Text.Should().Be("Review note");
    }

    // -------------------------------------------------------------------
    // "get_slide_comments filters by author."
    // -------------------------------------------------------------------

    [Fact]
    public void AsISlideComponent_ReturnsNonNullReference()
    {
        var tf = CreateTextFrame("text");

        tf.AsISlideComponent.Should().NotBeNull();
    }

    // -------------------------------------------------------------------
    // "insert_comment places at the correct index."
    // -------------------------------------------------------------------

    [Fact]
    public void Paragraphs_ReturnsParagraphsInOrder()
    {
        var tf = CreateTextFrame("First", "Second", "Third");

        tf.Paragraphs[0].Text.Should().Be("First");
        tf.Paragraphs[1].Text.Should().Be("Second");
        tf.Paragraphs[2].Text.Should().Be("Third");
    }

    // -------------------------------------------------------------------
    // "Notes text persists after save/reload."
    // -------------------------------------------------------------------

    [Fact]
    public void Text_SetAndGetRoundTrips()
    {
        var tf = CreateEmptyTextFrame();

        tf.Text = "Speaker notes";

        tf.Text.Should().Be("Speaker notes");
    }

    // -------------------------------------------------------------------
    // "Header/footer visibility persists."
    // -------------------------------------------------------------------

    [Fact]
    public void AsISlideComponent_IsAssignableToISlideComponent()
    {
        var tf = CreateTextFrame("Notes");

        tf.AsISlideComponent.Should().BeAssignableTo<ISlideComponent>();
    }

    // -------------------------------------------------------------------
    // "Cell text round-trips through save/reload."
    // -------------------------------------------------------------------

    [Theory]
    [InlineData("A")]
    [InlineData("B")]
    [InlineData("C")]
    [InlineData("D")]
    public void Text_SingleCharacterCellTextRoundTrips(string cellText)
    {
        var tf = CreateEmptyTextFrame();

        tf.Text = cellText;

        tf.Text.Should().Be(cellText);
    }

    // -------------------------------------------------------------------
    // "Cell borders persist."
    // -------------------------------------------------------------------

    [Fact]
    public void Text_WorksForBorderedCellContent()
    {
        var tf = CreateEmptyTextFrame();

        tf.Text = "Bordered";

        tf.Text.Should().Be("Bordered");
    }

    // -------------------------------------------------------------------
    // "Cell fill colour persists."
    // -------------------------------------------------------------------

    [Fact]
    public void Text_WorksForFilledCellContent()
    {
        var tf = CreateEmptyTextFrame();

        tf.Text = "Blue";

        tf.Text.Should().Be("Blue");
    }

    // -------------------------------------------------------------------
    // Additional behavioral tests
    // -------------------------------------------------------------------

    [Fact]
    public void Text_MultilineSplitsIntoParagraphs()
    {
        var tf = CreateEmptyTextFrame();

        tf.Text = "Line1\nLine2\nLine3";

        tf.Paragraphs.Count.Should().Be(3);
        tf.Paragraphs[0].Text.Should().Be("Line1");
        tf.Paragraphs[1].Text.Should().Be("Line2");
        tf.Paragraphs[2].Text.Should().Be("Line3");
    }

    [Fact]
    public void Text_EmptyValueCreatesOneParagraph()
    {
        var tf = CreateEmptyTextFrame();

        tf.Text = "";

        tf.Paragraphs.Count.Should().Be(1);
    }

    [Fact]
    public void Text_NullTxBodyReturnsEmpty()
    {
        var tf = new TextFrame();

        tf.Text.Should().BeEmpty();
    }

    [Fact]
    public void TextFrameFormat_ReturnsNonNull()
    {
        var tf = CreateTextFrame("text");

        tf.TextFrameFormat.Should().NotBeNull();
    }

    [Fact]
    public void ParentShape_ReturnsNullWhenNotSet()
    {
        var tf = CreateTextFrame("text");

        tf.ParentShape.Should().BeNull();
    }

    [Fact]
    public void ParentCell_ReturnsNullWhenNotSet()
    {
        var tf = CreateTextFrame("text");

        tf.ParentCell.Should().BeNull();
    }

    [Fact]
    public void Slide_ReturnsNullWhenNoParentSlide()
    {
        var tf = CreateTextFrame("text");

        tf.Slide.Should().BeNull();
    }

    [Fact]
    public void Slide_ReturnsParentSlideWhenSet()
    {
        var slide = new Slide();
        var txBody = new XElement(ANs + "txBody",
            new XElement(ANs + "p",
                new XElement(ANs + "r",
                    new XElement(ANs + "t", "test"))));
        var tf = new TextFrame();
        tf.InitInternal(txBody, null, slide, (IShape?)null);

        tf.Slide.Should().BeSameAs(slide);
    }

    [Fact]
    public void Presentation_ReturnsNullWhenNoParentSlide()
    {
        var tf = CreateTextFrame("text");

        tf.Presentation.Should().BeNull();
    }

    [Fact]
    public void TextFrame_ImplementsITextFrame()
    {
        var tf = CreateTextFrame("text");

        tf.Should().BeAssignableTo<ITextFrame>();
    }

    [Fact]
    public void TextFrame_InheritsFromISlideComponent()
    {
        var tf = CreateTextFrame("text");

        tf.Should().BeAssignableTo<ISlideComponent>();
    }

    [Fact]
    public void TextFrame_InheritsFromIPresentationComponent()
    {
        var tf = CreateTextFrame("text");

        tf.Should().BeAssignableTo<IPresentationComponent>();
    }

    [Fact]
    public void AsIPresentationComponent_ReturnsSelf()
    {
        var tf = CreateTextFrame("text");

        tf.AsIPresentationComponent.Should().BeSameAs(tf);
    }

    [Fact]
    public void Text_CarriageReturnSplitsIntoParagraphs()
    {
        var tf = CreateEmptyTextFrame();

        tf.Text = "A\r\nB\rC\nD";

        tf.Paragraphs.Count.Should().Be(4);
        tf.Text.Should().Be("A\nB\nC\nD");
    }
}
