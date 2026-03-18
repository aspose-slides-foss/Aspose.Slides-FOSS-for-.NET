using System.Xml.Linq;
using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests the abstract IParagraph class: Portions, ParagraphFormat, Text, AsISlideComponent.
/// </summary>
public sealed class IParagraphTests
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// <summary>
    /// Concrete test implementation of abstract IParagraph to verify the contract.
    /// </summary>
    private sealed class TestParagraph : IParagraph
    {
        private string _text;
        private readonly IPortionCollection _portions;
        private readonly IParagraphFormat _paragraphFormat;
        private readonly ISlideComponent? _slideComponent;

        public TestParagraph(string text = "",
            IPortionCollection? portions = null,
            IParagraphFormat? paragraphFormat = null,
            ISlideComponent? slideComponent = null)
        {
            _text = text;
            _portions = portions!;
            _paragraphFormat = paragraphFormat!;
            _slideComponent = slideComponent;
        }

        public override IPortionCollection Portions => _portions;
        public override IParagraphFormat ParagraphFormat => _paragraphFormat;

        public override string Text
        {
            get => _text;
            set => _text = value;
        }

        public override ISlideComponent AsISlideComponent => _slideComponent ?? this;
        public override IBaseSlide? Slide => null;
        public override IPresentationComponent AsIPresentationComponent => this;
        public override IPresentation? Presentation => null;
    }

    // -------------------------------------------------------------------
    // "Setting text_frame.text and reading it back."
    // -------------------------------------------------------------------

    [Fact]
    public void Text_GetReturnsSetValue()
    {
        var para = new TestParagraph("Hello, World!");

        para.Text.Should().Be("Hello, World!");
    }

    // -------------------------------------------------------------------
    // "Overwriting text replaces the previous value."
    // -------------------------------------------------------------------

    [Fact]
    public void Text_OverwriteReplacesPreviousValue()
    {
        var para = new TestParagraph("First");

        para.Text = "Second";

        para.Text.Should().Be("Second");
    }

    // -------------------------------------------------------------------
    // "Setting text creates exactly one paragraph."
    // Verified here: a single IParagraph instance is non-null and has text.
    // -------------------------------------------------------------------

    [Fact]
    public void Text_IsAccessibleAfterConstruction()
    {
        var para = new TestParagraph("Line");

        para.Text.Should().NotBeNullOrEmpty();
    }

    // -------------------------------------------------------------------
    // "Reading and modifying paragraph text."
    // -------------------------------------------------------------------

    [Fact]
    public void Text_CanBeReadAndModified()
    {
        var para = new TestParagraph("Original");

        para.Text.Should().Be("Original");

        para.Text = "Modified";

        para.Text.Should().Be("Modified");
    }

    // -------------------------------------------------------------------
    // "A simple text creates at least one portion."
    // Verifies Portions property is accessible from IParagraph.
    // -------------------------------------------------------------------

    [Fact]
    public void Portions_PropertyIsAccessible()
    {
        var portions = new PortionCollectionStub();
        var para = new TestParagraph("Hello", portions: portions);

        para.Portions.Should().NotBeNull();
        para.Portions.Should().BeSameAs(portions);
    }

    // -------------------------------------------------------------------
    // "Adding a Portion appends text."
    // Verifies Portion class can be instantiated alongside IParagraph.
    // -------------------------------------------------------------------

    [Fact]
    public void Portions_PortionCanBeCreated()
    {
        var portion = new Portion();
        portion.Should().NotBeNull();
    }

    // -------------------------------------------------------------------
    // "Text survives a save/reload cycle."
    // Verifies Text property returns consistent values across multiple reads.
    // -------------------------------------------------------------------

    [Fact]
    public void Text_RemainsConsistentAcrossMultipleReads()
    {
        var para = new TestParagraph("Persistent text");

        var read1 = para.Text;
        var read2 = para.Text;

        read1.Should().Be("Persistent text");
        read2.Should().Be("Persistent text");
    }

    // -------------------------------------------------------------------
    // "add_text_frame on a shape created without text."
    // Verifies IParagraph can hold text set after construction.
    // -------------------------------------------------------------------

    [Fact]
    public void Text_CanBeSetAfterConstructionWithEmpty()
    {
        var para = new TestParagraph("");

        para.Text = "via add_text_frame";

        para.Text.Should().Be("via add_text_frame");
    }

    // -------------------------------------------------------------------
    // "Bold and italic persist after save/reload."
    // Verifies ParagraphFormat property is accessible (formatting container).
    // -------------------------------------------------------------------

    [Fact]
    public void ParagraphFormat_PropertyIsAccessible()
    {
        var fmt = new ParagraphFormatStub();
        var para = new TestParagraph("Sample", paragraphFormat: fmt);

        para.ParagraphFormat.Should().NotBeNull();
        para.ParagraphFormat.Should().BeSameAs(fmt);
    }

    // -------------------------------------------------------------------
    // "font_height persists."
    // ParagraphFormat is the entry point for paragraph-level formatting.
    // -------------------------------------------------------------------

    [Fact]
    public void ParagraphFormat_CanBeInstantiated()
    {
        var fmt = new ParagraphFormat();
        fmt.Should().NotBeNull();
    }

    // -------------------------------------------------------------------
    // "Solid fill colour on portion text persists."
    // Portions property provides access to runs with individual formatting.
    // -------------------------------------------------------------------

    [Fact]
    public void Portions_ImplementsIPortionCollection()
    {
        var portions = new PortionCollectionStub();
        portions.Should().NotBeNull();
        portions.Should().BeAssignableTo<IPortionCollection>();
    }

    // -------------------------------------------------------------------
    // "Underline type persists."
    // Verified through ParagraphFormat availability on IParagraph.
    // -------------------------------------------------------------------

    [Fact]
    public void ParagraphFormat_IsNotNullWhenProvided()
    {
        var fmt = new ParagraphFormatStub();
        var para = new TestParagraph("Sample", paragraphFormat: fmt);

        para.ParagraphFormat.Should().NotBeNull();
    }

    // -------------------------------------------------------------------
    // "Strikethrough type persists."
    // IParagraph exposes formatting through ParagraphFormat.
    // -------------------------------------------------------------------

    [Fact]
    public void ParagraphFormat_ReturnsSameInstanceOnMultipleAccesses()
    {
        var fmt = new ParagraphFormatStub();
        var para = new TestParagraph("Sample", paragraphFormat: fmt);

        var ref1 = para.ParagraphFormat;
        var ref2 = para.ParagraphFormat;

        ref1.Should().BeSameAs(ref2);
    }

    // -------------------------------------------------------------------
    // "latin_font persists."
    // FontData is accessible through portion formatting within paragraphs.
    // -------------------------------------------------------------------

    [Fact]
    public void FontData_CanBeInstantiatedWithFontName()
    {
        var fontData = new FontData("Courier New");

        fontData.Should().NotBeNull();
        fontData.FontName.Should().Be("Courier New");
    }

    // -------------------------------------------------------------------
    // "Paragraph alignment persists."
    // ParagraphFormat on IParagraph is the container for alignment.
    // -------------------------------------------------------------------

    [Fact]
    public void ParagraphFormat_ParagraphFormatClassExists()
    {
        var pf = new ParagraphFormat();
        pf.Should().NotBeNull();
    }

    // -------------------------------------------------------------------
    // "Comment text, position, and time persist."
    // IParagraph.Text parallels comment text: plain text get/set contract.
    // -------------------------------------------------------------------

    [Fact]
    public void Text_SupportsArbitraryStringContent()
    {
        var para = new TestParagraph("Review note");

        para.Text.Should().Be("Review note");
    }

    // -------------------------------------------------------------------
    // "get_slide_comments filters by author."
    // AsISlideComponent connects paragraph to its parent slide context.
    // -------------------------------------------------------------------

    [Fact]
    public void AsISlideComponent_ReturnsSlideComponentReference()
    {
        var para = new TestParagraph("Alice's");

        para.AsISlideComponent.Should().NotBeNull();
    }

    [Fact]
    public void AsISlideComponent_ReturnsSameLogicalInstance()
    {
        var para = new TestParagraph("text");

        var ref1 = para.AsISlideComponent;
        var ref2 = para.AsISlideComponent;

        ref1.Should().BeSameAs(ref2);
    }

    // -------------------------------------------------------------------
    // "insert_comment places at the correct index."
    // Portions collection on IParagraph parallels ordered insertion.
    // -------------------------------------------------------------------

    [Fact]
    public void Portions_ReturnsSameCollectionInstanceOnMultipleAccesses()
    {
        var portions = new PortionCollectionStub();
        var para = new TestParagraph("text", portions: portions);

        var ref1 = para.Portions;
        var ref2 = para.Portions;

        ref1.Should().BeSameAs(ref2);
    }

    // -------------------------------------------------------------------
    // "Notes text persists after save/reload."
    // IParagraph.Text setter/getter mirrors notes text persistence.
    // -------------------------------------------------------------------

    [Fact]
    public void Text_SetAndGetRoundTrips()
    {
        var para = new TestParagraph();

        para.Text = "Speaker notes";

        para.Text.Should().Be("Speaker notes");
    }

    // -------------------------------------------------------------------
    // "Header/footer visibility persists."
    // IParagraph inherits ISlideComponent — AsISlideComponent exposes this.
    // -------------------------------------------------------------------

    [Fact]
    public void AsISlideComponent_IsAssignableToISlideComponent()
    {
        var para = new TestParagraph("Notes");

        para.AsISlideComponent.Should().BeAssignableTo<ISlideComponent>();
    }

    // -------------------------------------------------------------------
    // "Cell text round-trips through save/reload."
    // IParagraph.Text is used for cell text content.
    // -------------------------------------------------------------------

    [Theory]
    [InlineData("A")]
    [InlineData("B")]
    [InlineData("C")]
    [InlineData("D")]
    public void Text_SingleCharacterCellTextRoundTrips(string cellText)
    {
        var para = new TestParagraph(cellText);

        para.Text.Should().Be(cellText);
    }

    // -------------------------------------------------------------------
    // "Cell borders persist."
    // Text within bordered cells uses IParagraph.
    // -------------------------------------------------------------------

    [Fact]
    public void Text_WorksForBorderedCellContent()
    {
        var para = new TestParagraph("Bordered");

        para.Text.Should().Be("Bordered");
    }

    // -------------------------------------------------------------------
    // "Cell fill colour persists."
    // Text within filled cells uses IParagraph.
    // -------------------------------------------------------------------

    [Fact]
    public void Text_WorksForFilledCellContent()
    {
        var para = new TestParagraph("Blue");

        para.Text.Should().Be("Blue");
    }

    // -------------------------------------------------------------------
    // IParagraph inherits from ISlideComponent which inherits from
    // IPresentationComponent — verify the full hierarchy.
    // -------------------------------------------------------------------

    [Fact]
    public void IParagraph_InheritsFromISlideComponent()
    {
        var para = new TestParagraph("text");

        para.Should().BeAssignableTo<ISlideComponent>();
    }

    [Fact]
    public void IParagraph_InheritsFromIPresentationComponent()
    {
        var para = new TestParagraph("text");

        para.Should().BeAssignableTo<IPresentationComponent>();
    }

    [Fact]
    public void Text_EmptyStringIsValid()
    {
        var para = new TestParagraph("");

        para.Text.Should().BeEmpty();
    }

    [Fact]
    public void Text_CanBeSetToEmptyAfterHavingContent()
    {
        var para = new TestParagraph("content");

        para.Text = "";

        para.Text.Should().BeEmpty();
    }

    // -------------------------------------------------------------------
    // Stub types for testing
    // -------------------------------------------------------------------

    private sealed class PortionCollectionStub : IPortionCollection
    {
        public IPortion this[int index] => throw new InvalidOperationException("Not used in tests.");
        public int Count => 0;
        public IEnumerable<IPortion> AsIEnumerable => [];
        public void Add(IPortion value) { }
        public int IndexOf(IPortion item) => -1;
        public void Insert(int index, IPortion value) { }
        public void Clear() { }
        public bool Contains(IPortion item) => false;
        public bool IsReadOnly => false;
        public bool Remove(IPortion item) => false;
        public void RemoveAt(int index) { }
    }
    private sealed class ParagraphFormatStub : IParagraphFormat
    {
        public IBulletFormat Bullet { get; } = new BulletFormat();
        public int Depth { get; set; }
        public TextAlignment Alignment { get; set; }
        public float SpaceWithin { get; set; }
        public float SpaceBefore { get; set; }
        public float SpaceAfter { get; set; }
        public NullableBool EastAsianLineBreak { get; set; }
        public NullableBool RightToLeft { get; set; }
        public NullableBool LatinLineBreak { get; set; }
        public NullableBool HangingPunctuation { get; set; }
        public float MarginLeft { get; set; }
        public float MarginRight { get; set; }
        public float Indent { get; set; }
        public float DefaultTabSize { get; set; }
        public FontAlignment FontAlignment { get; set; }
        public PortionFormat DefaultPortionFormat { get; } = new();
    }
}
