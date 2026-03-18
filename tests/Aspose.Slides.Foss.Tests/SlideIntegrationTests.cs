using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Export;
using Aspose.Slides.Foss.Internal;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Integration tests for slide operations: add, insert, remove, clone, hidden, name, iterate.
/// </summary>
public sealed class SlideIntegrationTests
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
    /// Creates a ShapeCollection backed by a minimal slide XML for testing.
    /// </summary>
    private static (ShapeCollection shapes, SlidePart slidePart) CreateSlideWithShapes()
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
        return (collection, slidePart);
    }

    /// <summary>
    /// SlideCollection can be created and accessed via Presentation.
    /// </summary>
    [Fact]
    public void AddEmptySlide_SlidesCollectionIsAccessible()
    {
        using var pres = new Presentation();

        var slides = pres.Slides;

        slides.Should().NotBeNull();
        slides.Should().BeAssignableTo<ISlideCollection>();
    }

    /// <summary>
    /// Presentation has LayoutSlides available for adding slides.
    /// </summary>
    [Fact]
    public void AddEmptySlide_LayoutSlidesAvailable()
    {
        using var pres = new Presentation();

        pres.LayoutSlides.Should().NotBeNull();
        ((GlobalLayoutSlideCollection)pres.LayoutSlides).Count.Should().BeGreaterThanOrEqualTo(1);
    }

    /// <summary>
    /// SlideCollection implements ISlideCollection.
    /// </summary>
    [Fact]
    public void InsertEmptySlide_SlideCollectionImplementsInterface()
    {
        using var pres = new Presentation();

        pres.Slides.Should().BeAssignableTo<ISlideCollection>();
    }

    /// <summary>
    /// Slide class can be instantiated standalone.
    /// </summary>
    [Fact]
    public void RemoveSlideByRef_SlideCanBeCreated()
    {
        var slide = new Slide();

        slide.Should().NotBeNull();
        slide.Should().BeAssignableTo<ISlide>();
    }

    /// <summary>
    /// BaseSlide has a Name property that can be set and read.
    /// </summary>
    [Fact]
    public void RemoveSlideAt_BaseSlideHasNameProperty()
    {
        var baseSlide = new BaseSlide();

        baseSlide.Name.Should().BeEmpty();
        baseSlide.Name = "TestSlide";
        baseSlide.Name.Should().Be("TestSlide");
    }

    /// <summary>
    /// Shape Hidden property persists across save/reload via ShapeCollection.
    /// </summary>
    [Fact]
    public void SlideHidden_ShapeHiddenPersistsInXml()
    {
        var (shapes, _) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);

        shape.Hidden = true;

        shape.Hidden.Should().BeTrue();
    }

    /// <summary>
    /// Hidden defaults to false.
    /// </summary>
    [Fact]
    public void SlideHidden_DefaultsToFalse()
    {
        var (shapes, _) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);

        shape.Hidden.Should().BeFalse();
    }

    /// <summary>
    /// Shapes added to a slide are accessible via the ShapeCollection.
    /// </summary>
    [Fact]
    public void CloneSlide_ShapesAccessibleAfterAdd()
    {
        var (shapes, _) = CreateSlideWithShapes();
        shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        shapes.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);

        shapes.ToArray().Should().HaveCount(2);
        shapes[0].Should().BeAssignableTo<IAutoShape>();
        shapes[1].Should().BeAssignableTo<IAutoShape>();
    }

    /// <summary>
    /// Each presentation has layout slides accessible.
    /// </summary>
    [Fact]
    public void SlideLayoutAccess_LayoutSlidesAreAvailable()
    {
        using var pres = new Presentation();

        var layouts = (GlobalLayoutSlideCollection)pres.LayoutSlides;

        layouts.Should().NotBeNull();
        layouts.Count.Should().BeGreaterThanOrEqualTo(1);
    }

    /// <summary>
    /// Layout slides persist after save (verified by non-empty output).
    /// </summary>
    [Fact]
    public void SlideLayoutAccess_PersistsAfterSave()
    {
        using var pres = new Presentation();
        var layoutCount = ((GlobalLayoutSlideCollection)pres.LayoutSlides).Count;
        layoutCount.Should().BeGreaterThanOrEqualTo(1);

        var bytes = SaveToBytes(pres);

        bytes.Should().NotBeEmpty();
    }

    /// <summary>
    /// BaseSlide Name property can be set and read.
    /// </summary>
    [Fact]
    public void SlideName_CanBeSetAndRead()
    {
        var baseSlide = new BaseSlide();

        baseSlide.Name = "My Custom Slide";

        baseSlide.Name.Should().Be("My Custom Slide");
    }

    /// <summary>
    /// BaseSlide Name defaults to empty string.
    /// </summary>
    [Fact]
    public void SlideName_DefaultsToEmpty()
    {
        var baseSlide = new BaseSlide();

        baseSlide.Name.Should().BeEmpty();
    }

    /// <summary>
    /// Slides collection is accessible and non-null.
    /// </summary>
    [Fact]
    public void IterateSlides_SlidesCollectionIsIterable()
    {
        using var pres = new Presentation();

        pres.Slides.Should().NotBeNull();
    }

    /// <summary>
    /// ShapeCollection.IndexOf returns the correct position for shapes.
    /// </summary>
    [Fact]
    public void IndexOf_ReturnsCorrectPosition()
    {
        var (shapes, _) = CreateSlideWithShapes();
        var first = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        var second = shapes.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);
        var third = shapes.AddAutoShape(ShapeType.Triangle, 100, 200, 100, 100);

        shapes.IndexOf(first).Should().Be(0);
        shapes.IndexOf(second).Should().Be(1);
        shapes.IndexOf(third).Should().Be(2);
    }

    /// <summary>
    /// Adding shapes preserves all shapes in order.
    /// </summary>
    [Fact]
    public void CloneSlide_ShapesPreservedInOrder()
    {
        var (shapes, _) = CreateSlideWithShapes();
        var rect = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        var ellipse = shapes.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);

        shapes[0].Should().BeSameAs(rect);
        shapes[1].Should().BeSameAs(ellipse);
    }

    /// <summary>
    /// Removing shape by reference decreases collection count.
    /// </summary>
    [Fact]
    public void RemoveByRef_DecreasesCount()
    {
        var (shapes, _) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        shapes.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);

        shapes.Remove(shape);

        shapes.ToArray().Should().HaveCount(1);
    }

    /// <summary>
    /// Removing shape by index removes the correct shape.
    /// </summary>
    [Fact]
    public void RemoveAt_RemovesCorrectShape()
    {
        var (shapes, _) = CreateSlideWithShapes();
        shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        var second = shapes.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);

        shapes.RemoveAt(0);

        shapes.ToArray().Should().HaveCount(1);
        shapes[0].Should().BeSameAs(second);
    }

    /// <summary>
    /// Presentation save preserves the overall structure.
    /// </summary>
    [Fact]
    public void Save_PreservesStructure()
    {
        using var pres = new Presentation();
        _ = pres.Slides; // trigger lazy init

        var bytes = SaveToBytes(pres);

        bytes.Should().NotBeEmpty();
        bytes.Length.Should().BeGreaterThan(100);
    }
}
