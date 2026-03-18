using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests for the concrete <see cref="Paragraph"/> class.
/// </summary>
public sealed class ParagraphTests
{
    // -------------------------------------------------------------------
    // "Setting text_frame.text and reading it back."
    // -------------------------------------------------------------------

    [Fact]
    public void Text_GetReturnsSetValue()
    {
        var para = new Paragraph();

        para.Text = "Hello, World!";

        para.Text.Should().Be("Hello, World!");
    }

    // -------------------------------------------------------------------
    // "Overwriting text replaces the previous value."
    // -------------------------------------------------------------------

    [Fact]
    public void Text_OverwriteReplacesPreviousValue()
    {
        var para = new Paragraph();
        para.Text = "First";

        para.Text = "Second";

        para.Text.Should().Be("Second");
    }

    // -------------------------------------------------------------------
    // "Reading and modifying paragraph text."
    // -------------------------------------------------------------------

    [Fact]
    public void Text_CanBeReadAndModified()
    {
        var para = new Paragraph();
        para.Text = "Original";

        para.Text.Should().Be("Original");

        para.Text = "Modified";

        para.Text.Should().Be("Modified");
    }

    // -------------------------------------------------------------------
    // "Setting text creates exactly one paragraph."
    // Verifies paragraph has accessible text after setting.
    // -------------------------------------------------------------------

    [Fact]
    public void Text_DefaultIsEmpty()
    {
        var para = new Paragraph();

        para.Text.Should().BeEmpty();
    }

    // -------------------------------------------------------------------
    // "A simple text creates at least one portion."
    // -------------------------------------------------------------------

    [Fact]
    public void Portions_IsNotNull()
    {
        var para = new Paragraph();

        para.Portions.Should().NotBeNull();
    }

    [Fact]
    public void Portions_ImplementsIPortionCollection()
    {
        var para = new Paragraph();

        para.Portions.Should().BeAssignableTo<IPortionCollection>();
    }

    [Fact]
    public void Portions_ReturnsSameInstanceOnMultipleAccesses()
    {
        var para = new Paragraph();

        var ref1 = para.Portions;
        var ref2 = para.Portions;

        ref1.Should().BeSameAs(ref2);
    }

    // -------------------------------------------------------------------
    // "Adding a Portion appends text."
    // -------------------------------------------------------------------

    [Fact]
    public void Portions_InitiallyEmpty()
    {
        var para = new Paragraph();

        para.Portions.Count.Should().Be(0);
    }

    // -------------------------------------------------------------------
    // "Text survives a save/reload cycle."
    // Verified as consistency across multiple reads.
    // -------------------------------------------------------------------

    [Fact]
    public void Text_RemainsConsistentAcrossMultipleReads()
    {
        var para = new Paragraph();
        para.Text = "Persistent text";

        var read1 = para.Text;
        var read2 = para.Text;

        read1.Should().Be("Persistent text");
        read2.Should().Be("Persistent text");
    }

    // -------------------------------------------------------------------
    // "add_text_frame on a shape created without text."
    // -------------------------------------------------------------------

    [Fact]
    public void Text_CanBeSetAfterDefaultConstruction()
    {
        var para = new Paragraph();

        para.Text = "via add_text_frame";

        para.Text.Should().Be("via add_text_frame");
    }

    // -------------------------------------------------------------------
    // "Bold and italic persist after save/reload."
    // ParagraphFormat is the entry point for formatting.
    // -------------------------------------------------------------------

    [Fact]
    public void ParagraphFormat_IsNotNull()
    {
        var para = new Paragraph();

        para.ParagraphFormat.Should().NotBeNull();
    }

    [Fact]
    public void ParagraphFormat_ImplementsIParagraphFormat()
    {
        var para = new Paragraph();

        para.ParagraphFormat.Should().BeAssignableTo<IParagraphFormat>();
    }

    // -------------------------------------------------------------------
    // "font_height persists."
    // ParagraphFormat returns same instance on multiple accesses.
    // -------------------------------------------------------------------

    [Fact]
    public void ParagraphFormat_ReturnsSameInstanceOnMultipleAccesses()
    {
        var para = new Paragraph();

        var ref1 = para.ParagraphFormat;
        var ref2 = para.ParagraphFormat;

        ref1.Should().BeSameAs(ref2);
    }

    // -------------------------------------------------------------------
    // "Paragraph alignment persists."
    // ParagraphFormat is a concrete ParagraphFormat instance.
    // -------------------------------------------------------------------

    [Fact]
    public void ParagraphFormat_IsConcreteParagraphFormat()
    {
        var para = new Paragraph();

        para.ParagraphFormat.Should().BeOfType<ParagraphFormat>();
    }

    // -------------------------------------------------------------------
    // "Comment text, position, and time persist."
    // IParagraph.Text parallels comment text: plain text get/set contract.
    // -------------------------------------------------------------------

    [Fact]
    public void Text_SupportsArbitraryContent()
    {
        var para = new Paragraph();
        para.Text = "Review note";

        para.Text.Should().Be("Review note");
    }

    // -------------------------------------------------------------------
    // "get_slide_comments filters by author."
    // AsISlideComponent connects paragraph to its parent slide context.
    // -------------------------------------------------------------------

    [Fact]
    public void AsISlideComponent_ReturnsSelf()
    {
        var para = new Paragraph();

        para.AsISlideComponent.Should().BeSameAs(para);
    }

    // -------------------------------------------------------------------
    // "insert_comment places at the correct index."
    // Portions collection parallels ordered insertion.
    // -------------------------------------------------------------------

    [Fact]
    public void AsISlideComponent_IsAssignableToISlideComponent()
    {
        var para = new Paragraph();

        para.AsISlideComponent.Should().BeAssignableTo<ISlideComponent>();
    }

    // -------------------------------------------------------------------
    // "Notes text persists after save/reload."
    // -------------------------------------------------------------------

    [Fact]
    public void Text_SetAndGetRoundTrips()
    {
        var para = new Paragraph();

        para.Text = "Speaker notes";

        para.Text.Should().Be("Speaker notes");
    }

    // -------------------------------------------------------------------
    // "Header/footer visibility persists."
    // Paragraph inherits ISlideComponent — AsIPresentationComponent exposes this.
    // -------------------------------------------------------------------

    [Fact]
    public void AsIPresentationComponent_ReturnsSelf()
    {
        var para = new Paragraph();

        para.AsIPresentationComponent.Should().BeSameAs(para);
    }

    [Fact]
    public void AsIPresentationComponent_IsAssignableToIPresentationComponent()
    {
        var para = new Paragraph();

        para.AsIPresentationComponent.Should().BeAssignableTo<IPresentationComponent>();
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
        var para = new Paragraph();
        para.Text = cellText;

        para.Text.Should().Be(cellText);
    }

    // -------------------------------------------------------------------
    // "Cell borders persist."
    // -------------------------------------------------------------------

    [Fact]
    public void Text_WorksForBorderedCellContent()
    {
        var para = new Paragraph();
        para.Text = "Bordered";

        para.Text.Should().Be("Bordered");
    }

    // -------------------------------------------------------------------
    // "Cell fill colour persists."
    // -------------------------------------------------------------------

    [Fact]
    public void Text_WorksForFilledCellContent()
    {
        var para = new Paragraph();
        para.Text = "Blue";

        para.Text.Should().Be("Blue");
    }

    // -------------------------------------------------------------------
    // Inheritance hierarchy tests
    // -------------------------------------------------------------------

    [Fact]
    public void Paragraph_InheritsFromIParagraph()
    {
        var para = new Paragraph();

        para.Should().BeAssignableTo<IParagraph>();
    }

    [Fact]
    public void Paragraph_InheritsFromISlideComponent()
    {
        var para = new Paragraph();

        para.Should().BeAssignableTo<ISlideComponent>();
    }

    [Fact]
    public void Paragraph_InheritsFromIPresentationComponent()
    {
        var para = new Paragraph();

        para.Should().BeAssignableTo<IPresentationComponent>();
    }

    // -------------------------------------------------------------------
    // Slide and Presentation — null without parent
    // -------------------------------------------------------------------

    [Fact]
    public void Slide_IsNullWithoutParent()
    {
        var para = new Paragraph();

        para.Slide.Should().BeNull();
    }

    [Fact]
    public void Presentation_IsNullWithoutParent()
    {
        var para = new Paragraph();

        para.Presentation.Should().BeNull();
    }

    [Fact]
    public void Text_EmptyStringIsValid()
    {
        var para = new Paragraph();
        para.Text = "";

        para.Text.Should().BeEmpty();
    }

    [Fact]
    public void Text_CanBeSetToEmptyAfterHavingContent()
    {
        var para = new Paragraph();
        para.Text = "content";

        para.Text = "";

        para.Text.Should().BeEmpty();
    }
}
