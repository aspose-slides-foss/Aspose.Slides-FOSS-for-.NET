using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Export;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests the CellCollection public API contract: AsICollection, AsIEnumerable,
/// AsIPresentationComponent, AsISlideComponent, Presentation, Slide.
/// </summary>
public sealed class CellCollectionTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private sealed class RowStub : IRow
    {
        public float Height { get; init; }
        public float MinimalHeight { get; set; }
        public IRowFormat RowFormat => throw new InvalidOperationException("Not used in tests.");
        public ICellCollection AsICellCollection => throw new InvalidOperationException("Not used in tests.");
        public IBulkTextFormattable AsIBulkTextFormattable => throw new InvalidOperationException("Not used in tests.");
        public ICell this[int index] => throw new InvalidOperationException("Not used in tests.");
        public int Count => 0;
        public IBaseSlide? Slide => null;
        public IPresentation? Presentation => null;
        public IList<ICell> AsICollection => throw new InvalidOperationException("Not used in tests.");
        public IEnumerable<ICell> AsIEnumerable => throw new InvalidOperationException("Not used in tests.");
        public ISlideComponent AsISlideComponent => throw new InvalidOperationException("Not used in tests.");
        public IPresentationComponent AsIPresentationComponent => throw new InvalidOperationException("Not used in tests.");
        public void SetTextFormat(IBasePortionFormat source) => throw new InvalidOperationException("Not used in tests.");
        public void SetTextFormat(IParagraphFormat source) => throw new InvalidOperationException("Not used in tests.");
        public void SetTextFormat(ITextFrameFormat source) => throw new InvalidOperationException("Not used in tests.");
    }

    private sealed class ColumnStub : IColumn
    {
        public float Width { get; set; }
        public IColumnFormat ColumnFormat => throw new InvalidOperationException("Not used in tests.");
        public ICellCollection AsICellCollection => throw new InvalidOperationException("Not used in tests.");
        public IBulkTextFormattable AsIBulkTextFormattable => throw new InvalidOperationException("Not used in tests.");
        public ICell this[int index] => throw new InvalidOperationException("Not used in tests.");
        public int Count => 0;
        public IBaseSlide? Slide => null;
        public IPresentation? Presentation => null;
        public IList<ICell> AsICollection => throw new InvalidOperationException("Not used in tests.");
        public IEnumerable<ICell> AsIEnumerable => throw new InvalidOperationException("Not used in tests.");
        public ISlideComponent AsISlideComponent => throw new InvalidOperationException("Not used in tests.");
        public IPresentationComponent AsIPresentationComponent => throw new InvalidOperationException("Not used in tests.");
        public void SetTextFormat(IBasePortionFormat source) => throw new InvalidOperationException("Not used in tests.");
        public void SetTextFormat(IParagraphFormat source) => throw new InvalidOperationException("Not used in tests.");
        public void SetTextFormat(ITextFrameFormat source) => throw new InvalidOperationException("Not used in tests.");
    }

    private sealed class RowCollectionStub : IRowCollection
    {
        private readonly IReadOnlyList<IRow> _rows;
        public RowCollectionStub(IReadOnlyList<IRow> rows) => _rows = rows;
        public IRow this[int index] => _rows[index];
        public int Count => _rows.Count;
        public IList<IRow> AsICollection => throw new InvalidOperationException("Not used in tests.");
        public IEnumerable<IRow> AsIEnumerable => throw new InvalidOperationException("Not used in tests.");
        public IList<IRow> AddClone(IRow templ, bool withAttachedRows) => throw new InvalidOperationException("Not used in tests.");
        public IList<IRow> InsertClone(int index, IRow templ, bool withAttachedRows) => throw new InvalidOperationException("Not used in tests.");
        public void RemoveAt(int firstRowIndex, bool withAttachedRows) => throw new InvalidOperationException("Not used in tests.");
    }

    private sealed class ColumnCollectionStub : IColumnCollection
    {
        private readonly IReadOnlyList<IColumn> _columns;
        public ColumnCollectionStub(IReadOnlyList<IColumn> columns) => _columns = columns;
        public IColumn this[int index] => _columns[index];
        public int Count => _columns.Count;
        public IList<IColumn> AsICollection => throw new InvalidOperationException("Not used in tests.");
        public IEnumerable<IColumn> AsIEnumerable => throw new InvalidOperationException("Not used in tests.");
        public IList<IColumn> AddClone(IColumn templ, bool withAttachedColumns) => throw new InvalidOperationException("Not used in tests.");
        public IList<IColumn> InsertClone(int index, IColumn templ, bool withAttachedColumns) => throw new InvalidOperationException("Not used in tests.");
        public void RemoveAt(int index, bool withAttachedRows) => throw new InvalidOperationException("Not used in tests.");
    }

    private sealed class TableStub : ITable
    {
        public IRowCollection Rows { get; init; } = null!;
        public IColumnCollection Columns { get; init; } = null!;
        public ITableFormat TableFormat => throw new InvalidOperationException("Not used in tests.");
        public TableStylePreset StylePreset { get; set; }
        public bool RightToLeft { get; set; }
        public bool FirstRow { get; set; }
        public bool FirstCol { get; set; }
        public bool LastRow { get; set; }
        public bool LastCol { get; set; }
        public bool HorizontalBanding { get; set; }
        public bool VerticalBanding { get; set; }
        public IGraphicalObject AsIGraphicalObject => throw new InvalidOperationException("Not used in tests.");
        public IBulkTextFormattable AsIBulkTextFormattable => throw new InvalidOperationException("Not used in tests.");
        public ICell MergeCells(ICell cell1, ICell cell2, bool allowSplitting) => throw new InvalidOperationException("Not used in tests.");
        public IGraphicalObjectLock? GraphicalObjectLock => null;
        public bool IsTextHolder => false;
        public IPlaceholder? Placeholder => null;
        public ICustomData CustomData => null!;
        public ILineFormat LineFormat => throw new InvalidOperationException("Not used in tests.");
        public IThreeDFormat ThreeDFormat => throw new InvalidOperationException("Not used in tests.");
        public IEffectFormat EffectFormat => throw new InvalidOperationException("Not used in tests.");
        public IFillFormat FillFormat => throw new InvalidOperationException("Not used in tests.");
        public IShapeFrame RawFrame { get; set; } = null!;
        public IShapeFrame Frame { get; set; } = null!;
        public bool Hidden { get; set; }
        public float Rotation { get; set; }
        public float X { get; set; }
        public float Y { get; set; }
        public float Width { get; set; }
        public float Height { get; set; }
        public string AlternativeText { get; set; } = "";
        public string AlternativeTextTitle { get; set; } = "";
        public string Name { get; set; } = "";
        public bool IsDecorative { get; set; }
        public int UniqueId => 0;
        public int OfficeInteropShapeId => 0;
        public bool IsGrouped => false;
        public IGroupShape? ParentGroup => null;
        public IShapeStyle? ShapeStyle => null;
        public int ConnectionSiteCount => 0;
        public int ZOrderPosition => 0;
        public ShapeType ShapeType { get; set; }
        public IBaseSlide? Slide => null;
        public IPresentation? Presentation => null;
        public ISlideComponent AsISlideComponent => throw new InvalidOperationException("Not used in tests.");
        public IPresentationComponent AsIPresentationComponent => throw new InvalidOperationException("Not used in tests.");
        public void SetTextFormat(IBasePortionFormat source) => throw new InvalidOperationException("Not used in tests.");
        public void SetTextFormat(IParagraphFormat source) => throw new InvalidOperationException("Not used in tests.");
        public void SetTextFormat(ITextFrameFormat source) => throw new InvalidOperationException("Not used in tests.");
    }

    private sealed class BaseSlideStub : IBaseSlide
    {
        public IPresentation? Presentation { get; set; }
        public IShapeCollection? Shapes => null;
        public string Name { get; set; } = string.Empty;
        public int SlideId => 0;
    }

    private sealed class PresentationStub : IPresentation
    {
        public override DateTime CurrentDateTime { get; set; }
        public override ISlideCollection Slides => null!;
        public override INotesSize NotesSize => null!;
        public override IGlobalLayoutSlideCollection LayoutSlides => null!;
        public override IMasterSlideCollection Masters => null!;
        public override ICommentAuthorCollection CommentAuthors => null!;
        public override IDocumentProperties DocumentProperties => null!;
        public override IImageCollection Images => null!;
        public override SourceFormat SourceFormat => SourceFormat.Pptx;
        public override int FirstSlideNumber { get; set; }
        public override ISectionCollection Sections => null!;
        public override IPresentationComponent AsIPresentationComponent => this;
        public override void Save(string fname, SaveFormat format) => throw new InvalidOperationException("Not used in tests.");
        public override void Save(Stream stream, SaveFormat format) => throw new InvalidOperationException("Not used in tests.");
        public override void Save(string fname, SaveFormat format, ISaveOptions options) => throw new InvalidOperationException("Not used in tests.");
        public override void Save(Stream stream, SaveFormat format, ISaveOptions options) => throw new InvalidOperationException("Not used in tests.");
        public override void Save(string fname, int[] slides, SaveFormat format) => throw new InvalidOperationException("Not used in tests.");
        public override void Save(string fname, int[] slides, SaveFormat format, ISaveOptions options) => throw new InvalidOperationException("Not used in tests.");
        public override void Save(Stream stream, int[] slides, SaveFormat format) => throw new InvalidOperationException("Not used in tests.");
        public override void Save(Stream stream, int[] slides, SaveFormat format, ISaveOptions options) => throw new InvalidOperationException("Not used in tests.");
        public override void Save(ISaveOptions options) => throw new InvalidOperationException("Not used in tests.");
    }

    private static XElement MakeTcElement(string text = "")
    {
        return new XElement(ANs + "tc",
            new XElement(ANs + "txBody",
                new XElement(ANs + "bodyPr"),
                new XElement(ANs + "lstStyle"),
                new XElement(ANs + "p",
                    new XElement(ANs + "r",
                        new XElement(ANs + "t", text)))));
    }

    private static Cell CreateCell(string text = "", int rowIndex = 0, int colIndex = 0,
        IBaseSlide? parentSlide = null)
    {
        IRow[] rows = [new RowStub { Height = 40, MinimalHeight = 20 }];
        IColumn[] columns = [new ColumnStub { Width = 100 }];
        var table = new TableStub { Rows = new RowCollectionStub(rows), Columns = new ColumnCollectionStub(columns) };
        var tc = MakeTcElement(text);
        var cell = new Cell();
        cell.InitInternal(tc, rowIndex, colIndex, slidePart: null, parentSlide: parentSlide, table);
        return cell;
    }

    private static CellCollection CreateCollection(
        int cellCount = 0,
        IBaseSlide? parentSlide = null)
    {
        var cells = new List<Cell>();
        IRow[] rows = [new RowStub { Height = 40, MinimalHeight = 20 }];
        IColumn[] columns = [new ColumnStub { Width = 100 }];
        var table = new TableStub { Rows = new RowCollectionStub(rows), Columns = new ColumnCollectionStub(columns) };

        for (int i = 0; i < cellCount; i++)
        {
            var tc = MakeTcElement($"Cell{i}");
            var cell = new Cell();
            cell.InitInternal(tc, 0, i, slidePart: null, parentSlide: parentSlide, table);
            cells.Add(cell);
        }

        var collection = new CellCollection();
        collection.InitInternal(cells, slidePart: null, parentSlide: parentSlide);
        return collection;
    }

    // ── AsICollection ────────────────────────────────────────────────

    /// <summary>
    /// AsICollection returns a list with the correct count.
    /// </summary>
    [Fact]
    public void AsICollection_ReturnsListWithCorrectCount()
    {
        var collection = CreateCollection(cellCount: 3);

        var list = collection.AsICollection;

        list.Should().HaveCount(3);
    }

    /// <summary>
    /// AsICollection returns an independent copy.
    /// </summary>
    [Fact]
    public void AsICollection_ReturnsIndependentCopy()
    {
        var collection = CreateCollection(cellCount: 2);

        var list1 = collection.AsICollection;
        var list2 = collection.AsICollection;

        list1.Should().NotBeSameAs(list2);
    }

    /// <summary>
    /// AsICollection returns empty list for empty collection.
    /// </summary>
    [Fact]
    public void AsICollection_EmptyForEmptyCollection()
    {
        var collection = CreateCollection(cellCount: 0);

        collection.AsICollection.Should().BeEmpty();
    }

    /// <summary>
    /// AsICollection elements match indexer access.
    /// </summary>
    [Fact]
    public void AsICollection_ElementsMatchIndexer()
    {
        var collection = CreateCollection(cellCount: 2);

        var list = collection.AsICollection;

        list[0].Should().BeSameAs(collection[0]);
        list[1].Should().BeSameAs(collection[1]);
    }

    // ── AsIEnumerable ────────────────────────────────────────────────

    /// <summary>
    /// AsIEnumerable yields all cells.
    /// </summary>
    [Fact]
    public void AsIEnumerable_YieldsAllCells()
    {
        var collection = CreateCollection(cellCount: 3);

        var items = collection.AsIEnumerable.ToList();

        items.Should().HaveCount(3);
    }

    /// <summary>
    /// AsIEnumerable yields empty sequence for empty collection.
    /// </summary>
    [Fact]
    public void AsIEnumerable_EmptyForEmptyCollection()
    {
        var collection = CreateCollection(cellCount: 0);

        collection.AsIEnumerable.Should().BeEmpty();
    }

    /// <summary>
    /// AsIEnumerable returns a fresh enumerable each time.
    /// </summary>
    [Fact]
    public void AsIEnumerable_ReturnsFreshEnumerableEachTime()
    {
        var collection = CreateCollection(cellCount: 2);

        var enum1 = collection.AsIEnumerable;
        var enum2 = collection.AsIEnumerable;

        enum1.Should().NotBeSameAs(enum2);
    }

    /// <summary>
    /// AsIEnumerable elements match those from AsICollection.
    /// </summary>
    [Fact]
    public void AsIEnumerable_ElementsMatchAsICollection()
    {
        var collection = CreateCollection(cellCount: 3);

        var fromEnum = collection.AsIEnumerable.ToList();
        var fromList = collection.AsICollection;

        fromEnum.Should().Equal(fromList);
    }

    // ── AsISlideComponent ────────────────────────────────────────────

    /// <summary>
    /// AsISlideComponent returns the collection itself.
    /// </summary>
    [Fact]
    public void AsISlideComponent_ReturnsSelf()
    {
        var collection = CreateCollection();

        collection.AsISlideComponent.Should().BeSameAs(collection);
    }

    /// <summary>
    /// AsISlideComponent is of type ISlideComponent.
    /// </summary>
    [Fact]
    public void AsISlideComponent_IsISlideComponent()
    {
        var collection = CreateCollection();

        collection.AsISlideComponent.Should().BeAssignableTo<ISlideComponent>();
    }

    // ── AsIPresentationComponent ─────────────────────────────────────

    /// <summary>
    /// AsIPresentationComponent returns the collection itself.
    /// </summary>
    [Fact]
    public void AsIPresentationComponent_ReturnsSelf()
    {
        var collection = CreateCollection();

        collection.AsIPresentationComponent.Should().BeSameAs(collection);
    }

    /// <summary>
    /// AsIPresentationComponent is of type IPresentationComponent.
    /// </summary>
    [Fact]
    public void AsIPresentationComponent_IsIPresentationComponent()
    {
        var collection = CreateCollection();

        collection.AsIPresentationComponent.Should().BeAssignableTo<IPresentationComponent>();
    }

    // ── Slide ────────────────────────────────────────────────────────

    /// <summary>
    /// Slide returns the parent slide set during init.
    /// </summary>
    [Fact]
    public void Slide_ReturnsParentSlide()
    {
        var slide = new BaseSlideStub();
        var collection = CreateCollection(cellCount: 1, parentSlide: slide);

        collection.Slide.Should().BeSameAs(slide);
    }

    /// <summary>
    /// Slide is null when no parent slide is set.
    /// </summary>
    [Fact]
    public void Slide_NullWhenNoParentSlide()
    {
        var collection = CreateCollection();

        collection.Slide.Should().BeNull();
    }

    // ── Presentation ─────────────────────────────────────────────────

    /// <summary>
    /// Presentation returns the presentation from the parent slide.
    /// </summary>
    [Fact]
    public void Presentation_ReturnsPresentationFromSlide()
    {
        var pres = new PresentationStub();
        var slide = new BaseSlideStub { Presentation = pres };
        var collection = CreateCollection(cellCount: 1, parentSlide: slide);

        collection.Presentation.Should().BeSameAs(pres);
    }

    /// <summary>
    /// Presentation is null when no parent slide is set.
    /// </summary>
    [Fact]
    public void Presentation_NullWhenNoParentSlide()
    {
        var collection = CreateCollection();

        collection.Presentation.Should().BeNull();
    }

    /// <summary>
    /// Presentation is null when parent slide has no presentation.
    /// </summary>
    [Fact]
    public void Presentation_NullWhenSlideHasNoPresentation()
    {
        var slide = new BaseSlideStub { Presentation = null };
        var collection = CreateCollection(cellCount: 1, parentSlide: slide);

        collection.Presentation.Should().BeNull();
    }

    // ── Count / Indexer / Enumeration ────────────────────────────────

    /// <summary>
    /// Count returns the number of cells.
    /// </summary>
    [Fact]
    public void Count_ReturnsNumberOfCells()
    {
        var collection = CreateCollection(cellCount: 3);

        collection.Count.Should().Be(3);
    }

    /// <summary>
    /// Count is zero for empty collection.
    /// </summary>
    [Fact]
    public void Count_ZeroForEmptyCollection()
    {
        var collection = CreateCollection(cellCount: 0);

        collection.Count.Should().Be(0);
    }

    /// <summary>
    /// Indexer returns the correct cell.
    /// </summary>
    [Fact]
    public void Indexer_ReturnsCorrectCell()
    {
        var collection = CreateCollection(cellCount: 3);

        collection[0].Should().NotBeNull();
        collection[1].Should().NotBeNull();
        collection[2].Should().NotBeNull();
    }

    /// <summary>
    /// GetEnumerator yields all cells via foreach.
    /// </summary>
    [Fact]
    public void GetEnumerator_YieldsAllCells()
    {
        var collection = CreateCollection(cellCount: 2);

        var items = new List<ICell>();
        foreach (var cell in collection)
            items.Add(cell);

        items.Should().HaveCount(2);
    }

    /// <summary>
    /// LINQ works with the collection.
    /// </summary>
    [Fact]
    public void Enumerable_SupportsLinq()
    {
        var collection = CreateCollection(cellCount: 3);

        var count = collection.Count();

        count.Should().Be(3);
    }
}
