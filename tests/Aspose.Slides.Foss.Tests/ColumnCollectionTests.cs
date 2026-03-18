using System.Xml.Linq;
using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests the ColumnCollection public API: AddClone, InsertClone, RemoveAt,
/// AsICollection, AsIEnumerable.
/// </summary>
public sealed class ColumnCollectionTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const float EmuPerPoint = 12700f;

    /// <summary>
    /// Creates a ColumnCollection backed by a table XML element with the given column widths.
    /// Each column gets one cell per row (one row by default).
    /// </summary>
    private static ColumnCollection CreateCollection(params float[] widths)
    {
        var tblGrid = new XElement(ANs + "tblGrid");
        var tbl = new XElement(ANs + "tbl", tblGrid);

        var tr = new XElement(ANs + "tr", new XAttribute("h", "370840"));
        tbl.Add(tr);

        foreach (var w in widths)
        {
            var emu = (long)Math.Round(w * EmuPerPoint);
            tblGrid.Add(new XElement(ANs + "gridCol", new XAttribute("w", emu)));
            tr.Add(new XElement(ANs + "tc",
                new XElement(ANs + "txBody",
                    new XElement(ANs + "bodyPr"),
                    new XElement(ANs + "lstStyle"),
                    new XElement(ANs + "p",
                        new XElement(ANs + "endParaRPr"))),
                new XElement(ANs + "tcPr")));
        }

        var collection = new ColumnCollection();
        collection.InitInternal(tbl, tblGrid, slidePart: null, parentSlide: null, table: null);
        return collection;
    }

    // ---------------------------------------------------------------
    // and test_remove_at (shapes.remove_at)
    // ---------------------------------------------------------------

    /// <summary>
    /// Adding three items then removing the middle one leaves two items.
    /// </summary>
    [Fact]
    public void RemoveAt_MiddleElement_LeavesTwo()
    {
        var collection = CreateCollection(100, 200, 300);
        collection.Count.Should().Be(3);

        collection.RemoveAt(1, false);

        collection.Count.Should().Be(2);
    }

    /// <summary>
    /// remove_at(0) removes the first element by index.
    /// </summary>
    [Fact]
    public void RemoveAt_FirstElement_LeavesOne()
    {
        var collection = CreateCollection(100, 200);

        collection.RemoveAt(0, false);

        collection.Count.Should().Be(1);
    }

    /// <summary>
    /// remove_at on the last element reduces count.
    /// </summary>
    [Fact]
    public void RemoveAt_LastElement_LeavesOne()
    {
        var collection = CreateCollection(100, 200);

        collection.RemoveAt(1, false);

        collection.Count.Should().Be(1);
    }

    /// <summary>
    /// After removing the middle column, remaining columns are accessible.
    /// </summary>
    [Fact]
    public void RemoveAt_RemainingColumnsAreAccessible()
    {
        var collection = CreateCollection(100, 200, 300);

        collection.RemoveAt(1, false);

        collection[0].Width.Should().Be(100);
        collection[1].Width.Should().Be(300);
    }

    /// <summary>
    /// RemoveAt with invalid index throws.
    /// </summary>
    [Fact]
    public void RemoveAt_InvalidIndex_Throws()
    {
        var collection = CreateCollection(100);

        var act = () => collection.RemoveAt(5, false);

        act.Should().Throw<IndexOutOfRangeException>();
    }

    // ---------------------------------------------------------------
    // ---------------------------------------------------------------

    /// <summary>
    /// AddClone duplicates a column, increasing the count.
    /// </summary>
    [Fact]
    public void AddClone_IncreasesCount()
    {
        var collection = CreateCollection(100);
        var template = collection[0];

        collection.AddClone(template, false);

        collection.Count.Should().Be(2);
    }

    /// <summary>
    /// The cloned column preserves the template's width.
    /// </summary>
    [Fact]
    public void AddClone_PreservesWidth()
    {
        var collection = CreateCollection(150);
        var template = collection[0];

        var cloned = collection.AddClone(template, false);

        cloned.Should().ContainSingle().Which.Width.Should().Be(150);
    }

    /// <summary>
    /// AddClone appends to the end of the collection.
    /// </summary>
    [Fact]
    public void AddClone_AppendsToEnd()
    {
        var collection = CreateCollection(100, 200);
        var template = collection[0];

        collection.AddClone(template, false);

        collection.Count.Should().Be(3);
        collection[2].Width.Should().Be(100);
    }

    // ---------------------------------------------------------------
    // ---------------------------------------------------------------

    /// <summary>
    /// InsertClone places a cloned column at the requested index.
    /// </summary>
    [Fact]
    public void InsertClone_PlacesAtCorrectIndex()
    {
        var collection = CreateCollection(100, 300);
        var template = collection[0];

        collection.InsertClone(1, template, false);

        collection.Count.Should().Be(3);
        collection[1].Width.Should().Be(100);
    }

    /// <summary>
    /// InsertClone at index 0 prepends the column.
    /// </summary>
    [Fact]
    public void InsertClone_AtZero_Prepends()
    {
        var collection = CreateCollection(200);
        var template = collection[0];

        collection.InsertClone(0, template, false);

        collection.Count.Should().Be(2);
        collection[0].Width.Should().Be(200);
    }

    // ---------------------------------------------------------------
    // AsICollection – collection snapshot
    // ---------------------------------------------------------------

    /// <summary>
    /// AsICollection returns a list with the correct count.
    /// </summary>
    [Fact]
    public void AsICollection_ReturnsCorrectCount()
    {
        var collection = CreateCollection(100, 200, 300);

        var list = collection.AsICollection;

        list.Should().HaveCount(3);
    }

    /// <summary>
    /// AsICollection returns an independent copy (not the same reference).
    /// </summary>
    [Fact]
    public void AsICollection_ReturnsIndependentCopy()
    {
        var collection = CreateCollection(100, 200);

        var list1 = collection.AsICollection;
        var list2 = collection.AsICollection;

        list1.Should().NotBeSameAs(list2);
    }

    // ---------------------------------------------------------------
    // AsIEnumerable – enumeration
    // ---------------------------------------------------------------

    /// <summary>
    /// AsIEnumerable yields all columns.
    /// </summary>
    [Fact]
    public void AsIEnumerable_YieldsAllColumns()
    {
        var collection = CreateCollection(100, 200, 300);

        var items = collection.AsIEnumerable.ToList();

        items.Should().HaveCount(3);
    }

    /// <summary>
    /// Collection is iterable via foreach.
    /// </summary>
    [Fact]
    public void Collection_IsIterable()
    {
        var collection = CreateCollection(100, 200);

        var widths = new List<float>();
        foreach (var col in collection)
            widths.Add(col.Width);

        widths.Should().Equal(100f, 200f);
    }
}
