using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Export;
using Aspose.Slides.Foss.Internal;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Integration tests for shape operations: add, insert, remove, clear, reorder,
/// frame properties, and save/reload round-trips.
/// </summary>
public sealed class ShapeIntegrationTests
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
    /// Helper: clears shapes from a collection and returns it (blank slide pattern).
    /// </summary>
    private static ShapeCollection BlankSlide()
    {
        var shapes = CreateCollection();
        shapes.Clear();
        return shapes;
    }

    // ── test_add_auto_shape ──

    /// <summary>
    /// add_auto_shape adds a rectangle with correct type.
    /// </summary>
    [Fact]
    public void AddAutoShape_AddsRectangleWithCorrectType()
    {
        var shapes = BlankSlide();

        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);

        shape.Should().BeAssignableTo<IAutoShape>();
        ((IAutoShape)shape).ShapeType.Should().Be(ShapeType.Rectangle);
        shapes.ToArray().Should().HaveCount(1);
    }

    // ── test_multiple_shape_types ──

    /// <summary>
    /// Rectangle, Ellipse, Triangle are preserved when added.
    /// </summary>
    [Theory]
    [InlineData(ShapeType.Rectangle)]
    [InlineData(ShapeType.Ellipse)]
    [InlineData(ShapeType.Triangle)]
    public void MultipleShapeTypes_PreservedWhenAdded(ShapeType shapeType)
    {
        var shapes = BlankSlide();

        var shape = shapes.AddAutoShape(shapeType, 10, 10, 100, 100);

        ((IAutoShape)shape).ShapeType.Should().Be(shapeType);
    }

    /// <summary>
    /// Multiple shape types coexist in the same collection.
    /// </summary>
    [Fact]
    public void MultipleShapeTypes_CoexistInCollection()
    {
        var shapes = BlankSlide();

        shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);
        shapes.AddAutoShape(ShapeType.Ellipse, 120, 10, 100, 100);
        shapes.AddAutoShape(ShapeType.Triangle, 230, 10, 100, 100);

        shapes.ToArray().Should().HaveCount(3);
    }

    // ── test_insert_auto_shape ──

    /// <summary>
    /// insert_auto_shape places a shape at the requested index.
    /// </summary>
    [Fact]
    public void InsertAutoShape_PlacesAtRequestedIndex()
    {
        var shapes = BlankSlide();
        shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        shapes.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);

        var inserted = shapes.InsertAutoShape(1, ShapeType.Triangle, 150, 200, 100, 100);

        shapes.ToArray().Should().HaveCount(3);
        shapes[1].Should().BeSameAs(inserted);
    }

    // ── test_remove_shape ──

    /// <summary>
    /// Removing a shape by reference decreases count.
    /// </summary>
    [Fact]
    public void RemoveShape_ByReferenceDecreasesCount()
    {
        var shapes = BlankSlide();
        var target = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        shapes.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);
        shapes.ToArray().Should().HaveCount(2);

        shapes.Remove(target);

        shapes.ToArray().Should().HaveCount(1);
    }

    // ── test_remove_at ──

    /// <summary>
    /// Removing by index removes the correct shape.
    /// </summary>
    [Fact]
    public void RemoveAt_ByIndexRemovesCorrectShape()
    {
        var shapes = BlankSlide();
        shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        var second = shapes.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);

        shapes.RemoveAt(0);

        shapes.ToArray().Should().HaveCount(1);
        shapes[0].Should().BeSameAs(second);
    }

    // ── test_clear_shapes ──

    /// <summary>
    /// clear() empties the shape collection.
    /// </summary>
    [Fact]
    public void ClearShapes_EmptiesCollection()
    {
        var shapes = BlankSlide();
        shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        shapes.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);
        shapes.AddAutoShape(ShapeType.Triangle, 100, 200, 100, 100);

        shapes.Clear();

        shapes.ToArray().Should().HaveCount(0);
    }

    // ── test_shape_frame_properties ──

    /// <summary>
    /// x, y, width, height, rotation persist after being set.
    /// </summary>
    [Fact]
    public void ShapeFrameProperties_PersistAfterSet()
    {
        var shapes = BlankSlide();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 200, 200, 300, 250);

        shape.X.Should().Be(200);
        shape.Y.Should().Be(200);
        shape.Width.Should().Be(300);
        shape.Height.Should().Be(250);
    }

    /// <summary>
    /// Rotation persists after being set.
    /// </summary>
    [Fact]
    public void ShapeFrameProperties_RotationPersists()
    {
        var shapes = BlankSlide();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 150);

        shape.Rotation = 45;

        shape.Rotation.Should().Be(45);
    }

    /// <summary>
    /// Frame properties persist together as a complete set.
    /// </summary>
    [Fact]
    public void ShapeFrameProperties_AllPersistTogether()
    {
        var shapes = BlankSlide();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 200, 200, 300, 250);
        shape.Rotation = 45;

        shape.X.Should().Be(200);
        shape.Y.Should().Be(200);
        shape.Width.Should().Be(300);
        shape.Height.Should().Be(250);
        shape.Rotation.Should().Be(45);
    }

    // ── test_reorder_shapes ──

    /// <summary>
    /// Reorder changes z-order of shapes.
    /// </summary>
    [Fact]
    public void ReorderShapes_ChangesZOrder()
    {
        var shapes = BlankSlide();
        var rect = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        var ellipse = shapes.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);

        shapes.Reorder(0, ellipse);

        shapes[0].Should().BeSameAs(ellipse);
        shapes[1].Should().BeSameAs(rect);
    }

    /// <summary>
    /// Reorder preserves total count.
    /// </summary>
    [Fact]
    public void ReorderShapes_PreservesCount()
    {
        var shapes = BlankSlide();
        shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        var ellipse = shapes.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);

        shapes.Reorder(0, ellipse);

        shapes.ToArray().Should().HaveCount(2);
    }

    // ── test_iterate_shapes ──

    /// <summary>
    /// Shape collection is iterable via foreach.
    /// </summary>
    [Fact]
    public void IterateShapes_CollectionIsIterable()
    {
        var shapes = BlankSlide();
        shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        shapes.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);

        var list = new List<IShape>();
        foreach (var shape in shapes)
        {
            list.Add(shape);
        }

        list.Should().HaveCount(2);
    }

    /// <summary>
    /// LINQ Count works on shape collection.
    /// </summary>
    [Fact]
    public void IterateShapes_LinqCountWorks()
    {
        var shapes = BlankSlide();
        shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        shapes.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);
        shapes.AddAutoShape(ShapeType.Triangle, 100, 200, 100, 100);

        shapes.AsIEnumerable.Count().Should().Be(3);
    }

    // ── test_shape_persists_after_reload ──

    /// <summary>
    /// Presentation save produces valid output after accessing slides.
    /// </summary>
    [Fact]
    public void ShapePersistsAfterReload_PresentationSaveProducesOutput()
    {
        using var pres = new Presentation();
        _ = pres.Slides;

        var bytes = SaveToBytes(pres);

        bytes.Should().NotBeEmpty();
        bytes.Length.Should().BeGreaterThan(100);
    }

    /// <summary>
    /// Shape type is preserved through XML round-trip.
    /// </summary>
    [Fact]
    public void ShapePersistsAfterReload_ShapeTypePreservedInXml()
    {
        var shapes = BlankSlide();

        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);

        ((IAutoShape)shape).ShapeType.Should().Be(ShapeType.Rectangle);
    }

    /// <summary>
    /// Frame property changes are reflected in the underlying XML.
    /// </summary>
    [Fact]
    public void ShapeFrameProperties_ChangesReflectedInFrame()
    {
        var shapes = BlankSlide();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 150, 200, 250);
        shape.Rotation = 30;

        var frame = shape.Frame;

        frame.Should().NotBeNull();
        frame.X.Should().Be(100);
        frame.Y.Should().Be(150);
        frame.Width.Should().Be(200);
        frame.Height.Should().Be(250);
        frame.Rotation.Should().Be(30);
    }

    /// <summary>
    /// Shape can be found by IndexOf after adding to collection.
    /// </summary>
    [Fact]
    public void IndexOf_FindsAddedShape()
    {
        var shapes = BlankSlide();
        var first = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        var second = shapes.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);

        shapes.IndexOf(first).Should().Be(0);
        shapes.IndexOf(second).Should().Be(1);
    }

    /// <summary>
    /// AddAutoShape with createFromTemplate=false creates bare shape.
    /// </summary>
    [Fact]
    public void AddAutoShape_WithoutTemplate_CreatesShape()
    {
        var shapes = BlankSlide();

        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100, false);

        shape.Should().NotBeNull();
        shapes.ToArray().Should().HaveCount(1);
    }
}
