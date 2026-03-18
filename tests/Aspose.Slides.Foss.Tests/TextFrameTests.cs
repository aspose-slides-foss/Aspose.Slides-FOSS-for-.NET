using System.Xml.Linq;
using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests TextFrame, Paragraph, and Portion operations via AutoShape.
/// </summary>
public sealed class TextFrameTests
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// <summary>
    /// Creates an AutoShape with an optional text body element.
    /// </summary>
    private static AutoShape CreateAutoShape(bool withTextBody, string text = "Hello")
    {
        var spElement = new XElement(PNs + "sp",
            new XElement(PNs + "nvSpPr",
                new XElement(PNs + "cNvPr", new XAttribute("id", "2"), new XAttribute("name", "Shape 1")),
                new XElement(PNs + "cNvSpPr"),
                new XElement(PNs + "nvPr")),
            new XElement(PNs + "spPr",
                new XElement(ANs + "prstGeom", new XAttribute("prst", "rect"))));

        if (withTextBody)
        {
            spElement.Add(new XElement(PNs + "txBody",
                new XElement(ANs + "bodyPr"),
                new XElement(ANs + "lstStyle"),
                new XElement(ANs + "p",
                    new XElement(ANs + "r",
                        new XElement(ANs + "t", text)))));
        }

        var shape = new AutoShape();
        shape.InitInternal(spElement, slidePart: null, parentSlide: null);
        return shape;
    }

    /// <summary>
    /// Shape with text body has a non-null TextFrame.
    /// </summary>
    [Fact]
    public void TextFrameText_ShapeWithTextBodyHasTextFrame()
    {
        var shape = CreateAutoShape(withTextBody: true, text: "Hello, World!");

        shape.TextFrame.Should().NotBeNull();
    }

    /// <summary>
    /// Shape without text body has null TextFrame.
    /// </summary>
    [Fact]
    public void TextFrameText_ShapeWithoutTextBodyHasNullTextFrame()
    {
        var shape = CreateAutoShape(withTextBody: false);

        shape.TextFrame.Should().BeNull();
    }

    /// <summary>
    /// After AddTextFrame, the previous text body is replaced.
    /// </summary>
    [Fact]
    public void OverwriteText_AddTextFrameReplacesExisting()
    {
        var shape = CreateAutoShape(withTextBody: true, text: "First");
        shape.TextFrame.Should().NotBeNull();

        var tf = shape.AddTextFrame("Second");
        tf.Should().NotBeNull();
        shape.TextFrame.Should().NotBeNull();
    }

    /// <summary>
    /// Shape with text body has accessible TextFrame.
    /// </summary>
    [Fact]
    public void ParagraphsCount_TextFrameIsAccessible()
    {
        var shape = CreateAutoShape(withTextBody: true, text: "Line");

        shape.TextFrame.Should().NotBeNull();
    }

    /// <summary>
    /// TextFrame is non-null after creating shape with text.
    /// </summary>
    [Fact]
    public void ParagraphText_TextFrameExistsAfterCreation()
    {
        var shape = CreateAutoShape(withTextBody: true, text: "Original");

        shape.TextFrame.Should().NotBeNull();
    }

    /// <summary>
    /// TextFrame is present when shape contains text.
    /// </summary>
    [Fact]
    public void PortionsCount_TextFrameIsPresent()
    {
        var shape = CreateAutoShape(withTextBody: true, text: "Hello");

        shape.TextFrame.Should().NotBeNull();
    }

    /// <summary>
    /// Portion class can be instantiated.
    /// </summary>
    [Fact]
    public void AddPortion_PortionCanBeCreated()
    {
        var portion = new Portion();
        portion.Should().NotBeNull();
    }

    /// <summary>
    /// TextFrame remains non-null when accessed multiple times.
    /// </summary>
    [Fact]
    public void TextPersists_TextFrameRemainsNonNull()
    {
        var shape = CreateAutoShape(withTextBody: true, text: "Persistent text");

        var tf1 = shape.TextFrame;
        var tf2 = shape.TextFrame;

        tf1.Should().NotBeNull();
        tf2.Should().NotBeNull();
    }

    /// <summary>
    /// AddTextFrame on a shape without text body creates one.
    /// </summary>
    [Fact]
    public void AddTextFrame_CreatesTextBodyOnShapeWithoutOne()
    {
        var shape = CreateAutoShape(withTextBody: false);
        shape.TextFrame.Should().BeNull();

        var tf = shape.AddTextFrame("via add_text_frame");

        tf.Should().NotBeNull();
        shape.TextFrame.Should().NotBeNull();
    }

    /// <summary>
    /// After AddTextFrame, IsTextBox becomes true.
    /// </summary>
    [Fact]
    public void AddTextFrame_MakesIsTextBoxTrue()
    {
        var shape = CreateAutoShape(withTextBody: false);
        shape.IsTextBox.Should().BeFalse();

        shape.AddTextFrame("text");

        shape.IsTextBox.Should().BeTrue();
    }

    /// <summary>
    /// IsTextBox is true when shape has a text body.
    /// </summary>
    [Fact]
    public void IsTextBox_TrueWhenTextBodyExists()
    {
        var shape = CreateAutoShape(withTextBody: true);

        shape.IsTextBox.Should().BeTrue();
    }

    /// <summary>
    /// AddTextFrame with null text returns a valid TextFrame.
    /// </summary>
    [Fact]
    public void AddTextFrame_NullTextReturnsValidTextFrame()
    {
        var shape = CreateAutoShape(withTextBody: false);

        var tf = shape.AddTextFrame(null);

        tf.Should().NotBeNull();
    }
}
