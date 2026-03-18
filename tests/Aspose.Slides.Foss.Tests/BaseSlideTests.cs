using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Internal;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests for <see cref="BaseSlide"/>: shapes, name, slide ID, presentation.
/// </summary>
public sealed class BaseSlideTests
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";

    /// <summary>
    /// Creates a BaseSlide backed by a minimal slide XML with an empty spTree.
    /// </summary>
    private static BaseSlide CreateBaseSlide(string partName = "ppt/slides/slide1.xml")
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
        slidePart.InitInternal(partName);
        slidePart.Element = slideXml;

        var slide = new TestableBaseSlide(slidePart);
        return slide;
    }

    // ── Shapes property ──

    [Fact]
    public void Shapes_ReturnsNonNullCollection()
    {
        var slide = CreateBaseSlide();

        slide.Shapes.Should().NotBeNull();
    }

    [Fact]
    public void Shapes_ReturnsSameInstanceOnMultipleAccesses()
    {
        var slide = CreateBaseSlide();

        var first = slide.Shapes;
        var second = slide.Shapes;

        first.Should().BeSameAs(second);
    }

    [Fact]
    public void Shapes_InitiallyEmpty()
    {
        var slide = CreateBaseSlide();

        slide.Shapes!.ToArray().Should().BeEmpty();
    }

    [Fact]
    public void Shapes_CanAddAutoShape()
    {
        var slide = CreateBaseSlide();

        var shape = ((ShapeCollection)slide.Shapes!).AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);

        slide.Shapes!.ToArray().Should().HaveCount(1);
        shape.Should().BeAssignableTo<IAutoShape>();
    }

    [Fact]
    public void Shapes_IsNullWhenNoSlidePart()
    {
        var slide = new BaseSlide();

        slide.Shapes.Should().BeNull();
    }

    // ── Name property ──

    [Fact]
    public void Name_DefaultsToEmpty()
    {
        var slide = new BaseSlide();

        slide.Name.Should().BeEmpty();
    }

    [Fact]
    public void Name_CanBeSetAndRead()
    {
        var slide = new BaseSlide();

        slide.Name = "MySlide";

        slide.Name.Should().Be("MySlide");
    }

    [Fact]
    public void Name_CanBeOverwritten()
    {
        var slide = new BaseSlide();

        slide.Name = "First";
        slide.Name = "Second";

        slide.Name.Should().Be("Second");
    }

    // ── SlideId property ──

    [Fact]
    public void SlideId_DefaultsToZero()
    {
        var slide = new BaseSlide();

        slide.SlideId.Should().Be(0);
    }

    [Fact]
    public void SlideId_ExtractsFromPartName()
    {
        var slide = CreateBaseSlide("ppt/slides/slide3.xml");

        slide.SlideId.Should().Be(3);
    }

    [Fact]
    public void SlideId_ExtractsFromLayoutPartName()
    {
        var slide = CreateBaseSlide("ppt/slideLayouts/slideLayout5.xml");

        slide.SlideId.Should().Be(5);
    }

    [Fact]
    public void SlideId_ExtractsFromMasterPartName()
    {
        var slide = CreateBaseSlide("ppt/slideMasters/slideMaster2.xml");

        slide.SlideId.Should().Be(2);
    }

    // ── Presentation property ──

    [Fact]
    public void Presentation_DefaultsToNull()
    {
        var slide = new BaseSlide();

        slide.Presentation.Should().BeNull();
    }

    // ── IBaseSlide contract ──

    [Fact]
    public void ImplementsIBaseSlide()
    {
        var slide = new BaseSlide();

        slide.Should().BeAssignableTo<IBaseSlide>();
    }

    /// <summary>
    /// Test helper that exposes internal BaseSlide fields for testing.
    /// </summary>
    private sealed class TestableBaseSlide : BaseSlide
    {
        public TestableBaseSlide(SlidePart slidePart)
        {
            _slidePart = slidePart;
            _partName = slidePart.PartName;
        }
    }
}
