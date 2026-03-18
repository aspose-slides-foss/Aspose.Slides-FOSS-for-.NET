using System.Xml.Linq;
using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests the Portion class: construction, text get/set, portion format,
/// slide component and presentation component properties.
/// </summary>
public sealed class PortionTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    // -------------------------------------------------------------------
    // "Adding a Portion appends text."
    // Verifies Portion can be constructed with text.
    // -------------------------------------------------------------------

    [Fact]
    public void Constructor_WithText_SetsTextProperty()
    {
        var portion = new Portion("World!");

        portion.Text.Should().Be("World!");
    }

    [Fact]
    public void Constructor_NoArgs_SetsEmptyText()
    {
        var portion = new Portion();

        portion.Text.Should().Be("");
    }

    // -------------------------------------------------------------------
    // "Setting text_frame.text and reading it back."
    // Verifies text can be set and read back on a portion.
    // -------------------------------------------------------------------

    [Fact]
    public void Text_SetAndGet_RoundTrips()
    {
        var portion = new Portion();

        portion.Text = "Hello, World!";

        portion.Text.Should().Be("Hello, World!");
    }

    // -------------------------------------------------------------------
    // "Overwriting text replaces the previous value."
    // -------------------------------------------------------------------

    [Fact]
    public void Text_Overwrite_ReplacesPreviousValue()
    {
        var portion = new Portion("First");

        portion.Text = "Second";

        portion.Text.Should().Be("Second");
    }

    // -------------------------------------------------------------------
    // "Reading and modifying paragraph text."
    // Verifies text modification at the portion level.
    // -------------------------------------------------------------------

    [Fact]
    public void Text_Modify_UpdatesValue()
    {
        var portion = new Portion("Original");

        portion.Text = "Modified";

        portion.Text.Should().Be("Modified");
    }

    // -------------------------------------------------------------------
    // "A simple text creates at least one portion."
    // Verifies portion is a valid IPortion.
    // -------------------------------------------------------------------

    [Fact]
    public void Portion_IsIPortion()
    {
        var portion = new Portion("Hello");

        portion.Should().BeAssignableTo<IPortion>();
    }

    // -------------------------------------------------------------------
    // "Adding a Portion appends text."
    // Verifies Portion can be added to a PortionCollection and text is accessible.
    // -------------------------------------------------------------------

    [Fact]
    public void Add_ToCollection_TextAccessibleViaCollection()
    {
        var collection = new PortionCollection();
        var p1 = new Portion("Hello ");
        var p2 = new Portion("World!");
        collection.Add(p1);
        collection.Add(p2);

        collection.Count.Should().Be(2);
        ((Portion)collection[0]).Text.Should().Be("Hello ");
        ((Portion)collection[1]).Text.Should().Be("World!");
    }

    // -------------------------------------------------------------------
    // "Bold and italic persist after save/reload."
    // Verifies PortionFormat returns a non-null format object
    // and bold/italic can be set.
    // -------------------------------------------------------------------

    [Fact]
    public void PortionFormat_ReturnsNonNullFormat()
    {
        var portion = new Portion("Sample");

        portion.PortionFormat.Should().NotBeNull();
    }

    [Fact]
    public void PortionFormat_BoldItalic_CanBeSet()
    {
        var portion = new Portion("Sample");
        var fmt = portion.PortionFormat!;

        fmt.FontBold = NullableBool.True;
        fmt.FontItalic = NullableBool.True;

        fmt.FontBold.Should().Be(NullableBool.True);
        fmt.FontItalic.Should().Be(NullableBool.True);
    }

    // -------------------------------------------------------------------
    // "Underline type persists."
    // -------------------------------------------------------------------

    [Fact]
    public void PortionFormat_Underline_CanBeSet()
    {
        var portion = new Portion("Sample");
        var fmt = portion.PortionFormat!;

        fmt.FontUnderline = TextUnderlineType.Single;

        fmt.FontUnderline.Should().Be(TextUnderlineType.Single);
    }

    // -------------------------------------------------------------------
    // "Strikethrough type persists."
    // -------------------------------------------------------------------

    [Fact]
    public void PortionFormat_Strikethrough_CanBeSet()
    {
        var portion = new Portion("Sample");
        var fmt = portion.PortionFormat!;

        fmt.StrikethroughType = TextStrikethroughType.Single;

        fmt.StrikethroughType.Should().Be(TextStrikethroughType.Single);
    }

    // -------------------------------------------------------------------
    // "font_height persists."
    // -------------------------------------------------------------------

    [Fact]
    public void PortionFormat_FontHeight_CanBeSet()
    {
        var portion = new Portion("Sample");
        var fmt = portion.PortionFormat!;

        fmt.FontHeight = 28;

        fmt.FontHeight.Should().Be(28);
    }

    // -------------------------------------------------------------------
    // "Solid fill colour on portion text persists."
    // -------------------------------------------------------------------

    [Fact]
    public void PortionFormat_FillFormat_SolidColorCanBeSet()
    {
        var portion = new Portion("Sample");
        var fmt = portion.PortionFormat!;

        fmt.FillFormat!.FillType = FillType.Solid;
        fmt.FillFormat!.SolidFillColor.Color = Aspose.Slides.Foss.Drawing.Color.Red;

        // Re-read via a fresh PortionFormat access to verify XML persistence
        var fmt2 = portion.PortionFormat!;
        fmt2.FillFormat!.FillType.Should().Be(FillType.Solid);
        var c = fmt2.FillFormat!.SolidFillColor.Color;
        c!.R.Should().Be(255);
        c.G.Should().Be(0);
        c.B.Should().Be(0);
    }

    // -------------------------------------------------------------------
    // "latin_font persists."
    // -------------------------------------------------------------------

    [Fact]
    public void PortionFormat_LatinFont_CanBeSet()
    {
        var portion = new Portion("Sample");
        var fmt = portion.PortionFormat!;

        fmt.LatinFont = new FontData("Courier New");

        var fmt2 = portion.PortionFormat!;
        fmt2.LatinFont.Should().NotBeNull();
        fmt2.LatinFont!.FontName.Should().Be("Courier New");
    }

    // -------------------------------------------------------------------
    // Component properties
    // -------------------------------------------------------------------

    [Fact]
    public void AsISlideComponent_ReturnsSelf()
    {
        var portion = new Portion("test");

        portion.AsISlideComponent.Should().BeSameAs(portion);
    }

    [Fact]
    public void AsIPresentationComponent_ReturnsSelf()
    {
        var portion = new Portion("test");

        portion.AsIPresentationComponent.Should().BeSameAs(portion);
    }

    [Fact]
    public void Slide_WithoutParent_ReturnsNull()
    {
        var portion = new Portion("test");

        portion.Slide.Should().BeNull();
    }

    [Fact]
    public void Presentation_WithoutParent_ReturnsNull()
    {
        var portion = new Portion("test");

        portion.Presentation.Should().BeNull();
    }

    // -------------------------------------------------------------------
    // "Text survives a save/reload cycle."
    // Verifies text set via Portion persists in the XML element.
    // -------------------------------------------------------------------

    [Fact]
    public void Text_PersistsInXmlElement()
    {
        var rElement = new XElement(ANs + "r",
            new XElement(ANs + "t", "Initial"));

        var portion = new Portion();
        portion.InitInternal(rElement, pElement: null, txBodyElement: null, slidePart: null, parentSlide: null);

        portion.Text.Should().Be("Initial");
        portion.Text = "Persistent text";
        portion.Text.Should().Be("Persistent text");

        // Verify the XML was actually updated
        rElement.Element(ANs + "t")!.Value.Should().Be("Persistent text");
    }

    // -------------------------------------------------------------------
    // "add_text_frame on a shape created without text."
    // Verifies Portion initialized from existing XML element reads text.
    // -------------------------------------------------------------------

    [Fact]
    public void InitInternal_ReadsTextFromExistingElement()
    {
        var rElement = new XElement(ANs + "r",
            new XElement(ANs + "t", "via add_text_frame"));

        var portion = new Portion();
        portion.InitInternal(rElement, pElement: null, txBodyElement: null, slidePart: null, parentSlide: null);

        portion.Text.Should().Be("via add_text_frame");
    }

    // -------------------------------------------------------------------
    // Verifies format properties persist in the XML element across
    // multiple PortionFormat accesses.
    // -------------------------------------------------------------------

    [Fact]
    public void PortionFormat_PropertiesPersistAcrossAccesses()
    {
        var rElement = new XElement(ANs + "r",
            new XElement(ANs + "t", "Sample"));

        var portion = new Portion();
        portion.InitInternal(rElement, pElement: null, txBodyElement: null, slidePart: null, parentSlide: null);

        // Set via first access
        portion.PortionFormat!.FontBold = NullableBool.True;

        // Verify via second access (re-reads from same XML element)
        portion.PortionFormat!.FontBold.Should().Be(NullableBool.True);
    }
}
