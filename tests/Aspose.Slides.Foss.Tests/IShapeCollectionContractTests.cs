using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Internal;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests IShapeCollection contract: AddAutoShape, AddConnector, AddPictureFrame, AddTable,
/// InsertAutoShape, InsertConnector, InsertPictureShape, InsertTable,
/// Remove, RemoveAt, Clear, Reorder, IndexOf, ToArray, AsICollection, AsIEnumerable, ParentGroup.
/// </summary>
public sealed class IShapeCollectionContractTests
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace RNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>
    /// Creates a ShapeCollection backed by a minimal slide XML with an spTree,
    /// so that Add/Insert/Remove operations work against real XML.
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

    // ── AddAutoShape ──

    /// <summary>
    /// </summary>
    [Fact]
    public void AddAutoShape_IncreasesCount()
    {
        var shapes = CreateCollection();

        shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);

        shapes.ToArray().Should().HaveCount(1);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void AddAutoShape_ReturnsIAutoShape()
    {
        var shapes = CreateCollection();

        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);

        shape.Should().BeAssignableTo<IAutoShape>();
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void AddAutoShape_ShapeIsAccessibleInCollection()
    {
        var shapes = CreateCollection();

        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);

        shapes[0].Should().BeSameAs(shape);
    }

    /// <summary>
    /// </summary>
    [Theory]
    [InlineData(ShapeType.Rectangle)]
    [InlineData(ShapeType.Ellipse)]
    [InlineData(ShapeType.Triangle)]
    public void AddAutoShape_AcceptsVariousShapeTypes(ShapeType shapeType)
    {
        var shapes = CreateCollection();

        var shape = shapes.AddAutoShape(shapeType, 10, 10, 100, 100);

        shape.Should().NotBeNull();
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void AddAutoShape_MultipleShapesCoexist()
    {
        var shapes = CreateCollection();

        shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);
        shapes.AddAutoShape(ShapeType.Ellipse, 10, 10, 100, 100);
        shapes.AddAutoShape(ShapeType.Triangle, 10, 10, 100, 100);

        shapes.ToArray().Should().HaveCount(3);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void AddAutoShape_WithCreateFromTemplateFalse_ReturnsShape()
    {
        var shapes = CreateCollection();

        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 300, 100, false);

        shape.Should().NotBeNull();
        shapes.ToArray().Should().HaveCount(1);
    }

    // ── InsertAutoShape ──

    /// <summary>
    /// </summary>
    [Fact]
    public void InsertAutoShape_PlacesAtRequestedIndex()
    {
        var shapes = CreateCollection();
        shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        shapes.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);

        var inserted = shapes.InsertAutoShape(1, ShapeType.Triangle, 150, 200, 100, 100);

        shapes.ToArray().Should().HaveCount(3);
        shapes[1].Should().BeSameAs(inserted);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void InsertAutoShape_AtZero_Prepends()
    {
        var shapes = CreateCollection();
        shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);

        var inserted = shapes.InsertAutoShape(0, ShapeType.Ellipse, 10, 10, 100, 100);

        shapes[0].Should().BeSameAs(inserted);
    }

    /// <summary>
    /// InsertAutoShape with createFromTemplate parameter.
    /// </summary>
    [Fact]
    public void InsertAutoShape_WithCreateFromTemplateFalse_ReturnsShape()
    {
        var shapes = CreateCollection();
        shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);

        var shape = shapes.InsertAutoShape(0, ShapeType.Rectangle, 10, 10, 50, 50, false);

        shape.Should().NotBeNull();
        shapes.ToArray().Should().HaveCount(2);
    }

    // ── AddConnector ──

    /// <summary>
    /// </summary>
    [Fact]
    public void AddConnector_IncreasesCount()
    {
        var shapes = CreateCollection();

        shapes.AddConnector(ShapeType.StraightConnector1, 100, 100, 300, 200);

        shapes.ToArray().Should().HaveCount(1);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void AddConnector_ReturnsIConnector()
    {
        var shapes = CreateCollection();

        var conn = shapes.AddConnector(ShapeType.StraightConnector1, 100, 100, 300, 200);

        conn.Should().BeAssignableTo<IConnector>();
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void AddConnector_PersistsInCollection()
    {
        var shapes = CreateCollection();

        shapes.AddConnector(ShapeType.StraightConnector1, 100, 100, 300, 200);

        shapes.ToArray().Length.Should().BeGreaterThanOrEqualTo(1);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void AddConnector_BentConnector3_ReturnsConnector()
    {
        var shapes = CreateCollection();

        var conn = shapes.AddConnector(ShapeType.BentConnector3, 50, 50, 300, 200);

        conn.Should().NotBeNull();
    }

    /// <summary>
    /// AddConnector with createFromTemplate parameter.
    /// </summary>
    [Fact]
    public void AddConnector_WithCreateFromTemplateFalse_ReturnsConnector()
    {
        var shapes = CreateCollection();

        var conn = shapes.AddConnector(ShapeType.StraightConnector1, 100, 100, 300, 200, false);

        conn.Should().NotBeNull();
        shapes.ToArray().Should().HaveCount(1);
    }

    // ── InsertConnector ──

    /// <summary>
    /// InsertConnector places connector at requested index.
    /// </summary>
    [Fact]
    public void InsertConnector_PlacesAtRequestedIndex()
    {
        var shapes = CreateCollection();
        shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 100, 60);
        shapes.AddAutoShape(ShapeType.Rectangle, 350, 200, 100, 60);

        var conn = shapes.InsertConnector(1, ShapeType.BentConnector3, 0, 0, 1, 1);

        shapes.ToArray().Should().HaveCount(3);
        shapes[1].Should().BeSameAs(conn);
    }

    /// <summary>
    /// InsertConnector with createFromTemplate parameter.
    /// </summary>
    [Fact]
    public void InsertConnector_WithCreateFromTemplateFalse_Works()
    {
        var shapes = CreateCollection();
        shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 100, 60);

        var conn = shapes.InsertConnector(0, ShapeType.StraightConnector1, 0, 0, 1, 1, false);

        conn.Should().NotBeNull();
        shapes[0].Should().BeSameAs(conn);
    }

    // ── AddPictureFrame ──

    /// <summary>
    /// </summary>
    [Fact]
    public void AddPictureFrame_IncreasesCount()
    {
        var shapes = CreateCollection();
        var mockImage = new PPImage();

        shapes.AddPictureFrame(ShapeType.Rectangle, 50, 50, 100, 100, mockImage);

        shapes.ToArray().Should().HaveCount(1);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void AddPictureFrame_ReturnsIPictureFrame()
    {
        var shapes = CreateCollection();
        var mockImage = new PPImage();

        var frame = shapes.AddPictureFrame(ShapeType.Rectangle, 50, 50, 100, 100, mockImage);

        frame.Should().BeAssignableTo<IPictureFrame>();
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void AddPictureFrame_IsAccessibleInCollection()
    {
        var shapes = CreateCollection();
        var mockImage = new PPImage();

        var frame = shapes.AddPictureFrame(ShapeType.Rectangle, 50, 50, 200, 200, mockImage);

        shapes[0].Should().BeSameAs(frame);
    }

    // ── InsertPictureFrame ──

    /// <summary>
    /// InsertPictureFrame places at specified index.
    /// </summary>
    [Fact]
    public void InsertPictureFrame_PlacesAtRequestedIndex()
    {
        var shapes = CreateCollection();
        shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 100, 100);
        var mockImage = new PPImage();

        var frame = shapes.InsertPictureFrame(0, ShapeType.Rectangle, 10, 10, 50, 50, mockImage);

        shapes[0].Should().BeSameAs(frame);
        shapes.ToArray().Should().HaveCount(2);
    }

    // ── AddTable ──

    /// <summary>
    /// </summary>
    [Fact]
    public void AddTable_IncreasesCount()
    {
        var shapes = CreateCollection();

        shapes.AddTable(50, 50, [100, 150, 200], [40, 40, 40]);

        shapes.ToArray().Should().HaveCount(1);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void AddTable_ReturnsITable()
    {
        var shapes = CreateCollection();

        var table = shapes.AddTable(50, 50, [100, 150, 200], [40, 40, 40]);

        table.Should().BeAssignableTo<ITable>();
    }

    // ── InsertTable ──

    /// <summary>
    /// InsertTable places table at specified index.
    /// </summary>
    [Fact]
    public void InsertTable_PlacesAtRequestedIndex()
    {
        var shapes = CreateCollection();
        shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 100, 100);

        var table = shapes.InsertTable(0, 10, 10, [100], [40]);

        shapes[0].Should().BeSameAs(table);
        shapes.ToArray().Should().HaveCount(2);
    }

    // ── Remove ──

    /// <summary>
    /// </summary>
    [Fact]
    public void Remove_DecreasesCount()
    {
        var shapes = CreateCollection();
        var s = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        shapes.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);
        shapes.ToArray().Should().HaveCount(2);

        shapes.Remove(s);

        shapes.ToArray().Should().HaveCount(1);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void Remove_LeavesCorrectShapeRemaining()
    {
        var shapes = CreateCollection();
        var first = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        var second = shapes.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);

        shapes.Remove(first);

        shapes[0].Should().BeSameAs(second);
    }

    // ── RemoveAt ──

    /// <summary>
    /// </summary>
    [Fact]
    public void RemoveAt_RemovesByIndex()
    {
        var shapes = CreateCollection();
        shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        shapes.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);

        shapes.RemoveAt(0);

        shapes.ToArray().Should().HaveCount(1);
    }

    /// <summary>
    /// RemoveAt at middle index removes the correct element.
    /// </summary>
    [Fact]
    public void RemoveAt_MiddleIndex_RemovesCorrectElement()
    {
        var shapes = CreateCollection();
        shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 50, 50);
        var middle = shapes.AddAutoShape(ShapeType.Ellipse, 20, 20, 50, 50);
        shapes.AddAutoShape(ShapeType.Triangle, 30, 30, 50, 50);

        shapes.RemoveAt(1);

        shapes.ToArray().Should().HaveCount(2);
        shapes.ToArray().Should().NotContain(middle);
    }

    // ── Clear ──

    /// <summary>
    /// </summary>
    [Fact]
    public void Clear_EmptiesCollection()
    {
        var shapes = CreateCollection();
        shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        shapes.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);

        shapes.Clear();

        shapes.ToArray().Should().HaveCount(0);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void Clear_OnEmptyCollection_DoesNotThrow()
    {
        var shapes = CreateCollection();

        var act = () => shapes.Clear();

        act.Should().NotThrow();
    }

    // ── Reorder ──

    /// <summary>
    /// </summary>
    [Fact]
    public void Reorder_ChangesZOrder()
    {
        var shapes = CreateCollection();
        var rect = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        var ellipse = shapes.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);

        shapes.Reorder(0, ellipse);

        shapes[0].Should().BeSameAs(ellipse);
        shapes[1].Should().BeSameAs(rect);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void Reorder_PreservesCount()
    {
        var shapes = CreateCollection();
        shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        var ellipse = shapes.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);

        shapes.Reorder(0, ellipse);

        shapes.ToArray().Should().HaveCount(2);
    }

    /// <summary>
    /// Reorder with multiple shapes preserves relative order.
    /// </summary>
    [Fact]
    public void Reorder_MultipleShapes_PreservesRelativeOrder()
    {
        var shapes = CreateCollection();
        var a = shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 50, 50);
        var b = shapes.AddAutoShape(ShapeType.Ellipse, 20, 20, 50, 50);
        var c = shapes.AddAutoShape(ShapeType.Triangle, 30, 30, 50, 50);

        shapes.Reorder(0, [b, c]);

        shapes[0].Should().BeSameAs(b);
        shapes[1].Should().BeSameAs(c);
        shapes[2].Should().BeSameAs(a);
    }

    // ── IndexOf ──

    /// <summary>
    /// </summary>
    [Fact]
    public void IndexOf_ReturnsCorrectPosition()
    {
        var shapes = CreateCollection();
        var first = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        var second = shapes.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);

        shapes.IndexOf(first).Should().Be(0);
        shapes.IndexOf(second).Should().Be(1);
    }

    /// <summary>
    /// IndexOf returns -1 for a shape not in the collection.
    /// </summary>
    [Fact]
    public void IndexOf_ReturnsNegativeOneForMissingShape()
    {
        var shapes = CreateCollection();
        shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);

        var other = CreateCollection();
        var external = other.AddAutoShape(ShapeType.Ellipse, 10, 10, 50, 50);

        shapes.IndexOf(external).Should().Be(-1);
    }

    // ── ToArray ──

    /// <summary>
    /// </summary>
    [Fact]
    public void ToArray_ReturnsAllShapes()
    {
        var shapes = CreateCollection();
        shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        shapes.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);

        var array = shapes.ToArray();

        array.Should().HaveCount(2);
    }

    /// <summary>
    /// ToArray returns empty array for empty collection.
    /// </summary>
    [Fact]
    public void ToArray_ReturnsEmptyForEmptyCollection()
    {
        var shapes = CreateCollection();

        shapes.ToArray().Should().BeEmpty();
    }

    /// <summary>
    /// ToArray with startIndex and count returns correct subset.
    /// </summary>
    [Fact]
    public void ToArray_WithRange_ReturnsSubset()
    {
        var shapes = CreateCollection();
        shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 50, 50);
        var second = shapes.AddAutoShape(ShapeType.Ellipse, 20, 20, 50, 50);
        shapes.AddAutoShape(ShapeType.Triangle, 30, 30, 50, 50);

        var subset = shapes.ToArray(1, 1);

        subset.Should().HaveCount(1);
        subset[0].Should().BeSameAs(second);
    }

    // ── AsICollection ──

    /// <summary>
    /// AsICollection exposes shapes as IList.
    /// </summary>
    [Fact]
    public void AsICollection_ReturnsAllShapes()
    {
        var shapes = CreateCollection();
        shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        shapes.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);

        IList<IShape> list = shapes.AsICollection;

        list.Should().HaveCount(2);
    }

    /// <summary>
    /// AsICollection is assignable to IList.
    /// </summary>
    [Fact]
    public void AsICollection_IsAssignableToIList()
    {
        var shapes = CreateCollection();

        shapes.AsICollection.Should().BeAssignableTo<IList<IShape>>();
    }

    // ── AsIEnumerable ──

    /// <summary>
    /// AsIEnumerable exposes shapes as IEnumerable.
    /// </summary>
    [Fact]
    public void AsIEnumerable_ReturnsAllShapes()
    {
        var shapes = CreateCollection();
        shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        shapes.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);

        shapes.AsIEnumerable.Count().Should().Be(2);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void AsIEnumerable_SupportsForEach()
    {
        var shapes = CreateCollection();
        shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        shapes.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);

        var list = new List<IShape>();
        foreach (var shape in shapes)
        {
            list.Add(shape);
        }

        list.Should().HaveCount(2);
    }

    // ── ParentGroup ──

    /// <summary>
    /// ParentGroup defaults to null for a standalone collection.
    /// </summary>
    [Fact]
    public void ParentGroup_DefaultsToNull()
    {
        var shapes = CreateCollection();

        shapes.ParentGroup.Should().BeNull();
    }

    // ── Indexer ──

    /// <summary>
    /// Indexer returns shape at specified position.
    /// </summary>
    [Fact]
    public void Indexer_ReturnsShapeAtPosition()
    {
        var shapes = CreateCollection();
        var first = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        var second = shapes.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);

        shapes[0].Should().BeSameAs(first);
        shapes[1].Should().BeSameAs(second);
    }

    // ── Shape frame properties persist through collection (test_shape_frame_properties) ──

    /// <summary>
    /// </summary>
    [Fact]
    public void AddAutoShape_ReturnedShapeIsSameAsRetrieved()
    {
        var shapes = CreateCollection();

        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 200, 200, 300, 250);

        shapes[0].Should().BeSameAs(shape);
    }

    // ── Mixed shape types in collection ──

    /// <summary>
    /// </summary>
    [Fact]
    public void Collection_HoldsMixedShapeTypes()
    {
        var shapes = CreateCollection();
        shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 100, 60);
        shapes.AddAutoShape(ShapeType.Rectangle, 350, 200, 100, 60);
        shapes.AddConnector(ShapeType.BentConnector3, 0, 0, 1, 1);

        shapes.ToArray().Should().HaveCount(3);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void Collection_HoldsAllShapeTypes()
    {
        var shapes = CreateCollection();
        var mockImage = new PPImage();

        shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 100, 100);
        shapes.AddConnector(ShapeType.StraightConnector1, 0, 0, 100, 100);
        shapes.AddPictureFrame(ShapeType.Rectangle, 50, 50, 100, 100, mockImage);
        shapes.AddTable(50, 50, [100], [40]);

        shapes.ToArray().Should().HaveCount(4);
    }

    // ── Implements IShapeCollection ──

    /// <summary>
    /// ShapeCollection implements IShapeCollection interface.
    /// </summary>
    [Fact]
    public void ShapeCollection_ImplementsIShapeCollection()
    {
        var shapes = CreateCollection();

        shapes.Should().BeAssignableTo<IShapeCollection>();
    }

    /// <summary>
    /// ShapeCollection implements IEnumerable of IShape.
    /// </summary>
    [Fact]
    public void ShapeCollection_ImplementsIEnumerableOfIShape()
    {
        var shapes = CreateCollection();

        shapes.Should().BeAssignableTo<IEnumerable<IShape>>();
    }

    // ── RemoveAt after insert (test_remove_slide_at pattern) ──

    /// <summary>
    /// </summary>
    [Fact]
    public void RemoveAt_AfterAdd_RestoresCount()
    {
        var shapes = CreateCollection();
        shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        shapes.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);
        shapes.ToArray().Should().HaveCount(2);

        shapes.RemoveAt(1);

        shapes.ToArray().Should().HaveCount(1);
    }

    // ── Clone-like behavior (test_clone_slide) ──

    /// <summary>
    /// </summary>
    [Fact]
    public void MultipleAdds_AllShapesAccessibleByIndex()
    {
        var shapes = CreateCollection();
        shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        shapes.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);
        shapes.AddConnector(ShapeType.StraightConnector1, 0, 0, 100, 50);

        shapes[0].Should().NotBeNull();
        shapes[1].Should().NotBeNull();
        shapes[2].Should().NotBeNull();
    }

    // ── Clear and re-add (test_bent_connector_adjustments pattern) ──

    /// <summary>
    /// </summary>
    [Fact]
    public void Clear_ThenAdd_WorksCorrectly()
    {
        var shapes = CreateCollection();
        shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        shapes.Clear();
        shapes.ToArray().Should().HaveCount(0);

        shapes.AddConnector(ShapeType.BentConnector3, 50, 50, 300, 200);

        shapes.ToArray().Should().HaveCount(1);
    }
}
