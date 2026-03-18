using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Export;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests the Cell public API contract.
/// </summary>
public sealed class CellTests
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
        public override ISectionCollection Sections => null!;
        public override ICommentAuthorCollection CommentAuthors => null!;
        public override IDocumentProperties DocumentProperties => null!;
        public override IImageCollection Images => null!;
        public override SourceFormat SourceFormat => SourceFormat.Pptx;
        public override int FirstSlideNumber { get; set; }
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

    /// <summary>
    /// Creates a Cell backed by a tc XML element within a table stub.
    /// </summary>
    private static Cell CreateCell(
        XElement? tcElement = null,
        int rowIndex = 0,
        int colIndex = 0,
        IReadOnlyList<IRow>? rows = null,
        IReadOnlyList<IColumn>? columns = null,
        IBaseSlide? parentSlide = null)
    {
        rows ??= [new RowStub { Height = 40, MinimalHeight = 20 }];
        columns ??= [new ColumnStub { Width = 100 }];
        var table = new TableStub { Rows = new RowCollectionStub(rows), Columns = new ColumnCollectionStub(columns) };

        tcElement ??= new XElement(ANs + "tc",
            new XElement(ANs + "txBody",
                new XElement(ANs + "bodyPr"),
                new XElement(ANs + "lstStyle"),
                new XElement(ANs + "p",
                    new XElement(ANs + "r",
                        new XElement(ANs + "t", "")))));

        var cell = new Cell();
        cell.InitInternal(tcElement, rowIndex, colIndex, slidePart: null, parentSlide: parentSlide, table);
        return cell;
    }

    private static XElement MakeTcElement(string text = "", string? gridSpan = null, string? rowSpan = null,
        string? hMerge = null, string? vMerge = null)
    {
        var tc = new XElement(ANs + "tc",
            new XElement(ANs + "txBody",
                new XElement(ANs + "bodyPr"),
                new XElement(ANs + "lstStyle"),
                new XElement(ANs + "p",
                    new XElement(ANs + "r",
                        new XElement(ANs + "t", text)))));

        if (gridSpan is not null) tc.Add(new XAttribute("gridSpan", gridSpan));
        if (rowSpan is not null) tc.Add(new XAttribute("rowSpan", rowSpan));
        if (hMerge is not null) tc.Add(new XAttribute("hMerge", hMerge));
        if (vMerge is not null) tc.Add(new XAttribute("vMerge", vMerge));

        return tc;
    }

    // ── TextFrame ──────────────────────────────────────────────────────

    /// <summary>
    /// Cell with txBody element returns a non-null TextFrame.
    /// </summary>
    [Fact]
    public void TextFrame_NotNullWhenTxBodyPresent()
    {
        var tc = MakeTcElement("A");
        var cell = CreateCell(tc);

        cell.TextFrame.Should().NotBeNull();
    }

    /// <summary>
    /// Cell without txBody element returns null TextFrame.
    /// </summary>
    [Fact]
    public void TextFrame_NullWhenNoTxBody()
    {
        var tc = new XElement(ANs + "tc");
        var cell = CreateCell(tc);

        cell.TextFrame.Should().BeNull();
    }

    // ── ColSpan / RowSpan ──────────────────────────────────────────────

    /// <summary>
    /// ColSpan reads from gridSpan XML attribute.
    /// </summary>
    [Fact]
    public void ColSpan_ReadsFromGridSpanAttribute()
    {
        var tc = MakeTcElement(gridSpan: "2");
        IColumn[] columns =
        [
            new ColumnStub { Width = 100 },
            new ColumnStub { Width = 100 },
        ];
        var cell = CreateCell(tc, columns: columns);

        cell.ColSpan.Should().Be(2);
    }

    /// <summary>
    /// ColSpan defaults to 1 when no gridSpan attribute.
    /// </summary>
    [Fact]
    public void ColSpan_DefaultsToOne()
    {
        var cell = CreateCell();

        cell.ColSpan.Should().Be(1);
    }

    /// <summary>
    /// RowSpan reads from rowSpan XML attribute.
    /// </summary>
    [Fact]
    public void RowSpan_ReadsFromRowSpanAttribute()
    {
        var tc = MakeTcElement(rowSpan: "3");
        IRow[] rows =
        [
            new RowStub { Height = 40, MinimalHeight = 20 },
            new RowStub { Height = 40, MinimalHeight = 20 },
            new RowStub { Height = 40, MinimalHeight = 20 },
        ];
        var cell = CreateCell(tc, rows: rows);

        cell.RowSpan.Should().Be(3);
    }

    /// <summary>
    /// RowSpan defaults to 1 when no rowSpan attribute.
    /// </summary>
    [Fact]
    public void RowSpan_DefaultsToOne()
    {
        var cell = CreateCell();

        cell.RowSpan.Should().Be(1);
    }

    // ── IsMergedCell ──────────────────────────────────────────────────

    /// <summary>
    /// IsMergedCell is true when gridSpan > 1.
    /// </summary>
    [Fact]
    public void IsMergedCell_TrueWhenGridSpanGreaterThanOne()
    {
        var tc = MakeTcElement(gridSpan: "2");
        IColumn[] columns =
        [
            new ColumnStub { Width = 100 },
            new ColumnStub { Width = 100 },
        ];
        var cell = CreateCell(tc, columns: columns);

        cell.IsMergedCell.Should().BeTrue();
    }

    /// <summary>
    /// IsMergedCell is true when rowSpan > 1.
    /// </summary>
    [Fact]
    public void IsMergedCell_TrueWhenRowSpanGreaterThanOne()
    {
        var tc = MakeTcElement(rowSpan: "2");
        IRow[] rows =
        [
            new RowStub { Height = 40, MinimalHeight = 20 },
            new RowStub { Height = 40, MinimalHeight = 20 },
        ];
        var cell = CreateCell(tc, rows: rows);

        cell.IsMergedCell.Should().BeTrue();
    }

    /// <summary>
    /// IsMergedCell is true when hMerge attribute is set.
    /// </summary>
    [Fact]
    public void IsMergedCell_TrueWhenHMerge()
    {
        var tc = MakeTcElement(hMerge: "1");
        var cell = CreateCell(tc);

        cell.IsMergedCell.Should().BeTrue();
    }

    /// <summary>
    /// IsMergedCell is true when vMerge attribute is set.
    /// </summary>
    [Fact]
    public void IsMergedCell_TrueWhenVMerge()
    {
        var tc = MakeTcElement(vMerge: "true");
        var cell = CreateCell(tc);

        cell.IsMergedCell.Should().BeTrue();
    }

    /// <summary>
    /// IsMergedCell is false for a simple non-merged cell.
    /// </summary>
    [Fact]
    public void IsMergedCell_FalseForSimpleCell()
    {
        var cell = CreateCell();

        cell.IsMergedCell.Should().BeFalse();
    }

    // ── CellFormat ────────────────────────────────────────────────────

    /// <summary>
    /// CellFormat is accessible and not null.
    /// </summary>
    [Fact]
    public void CellFormat_IsNotNull()
    {
        var cell = CreateCell();

        cell.CellFormat.Should().NotBeNull();
    }

    // ── Width / Height ────────────────────────────────────────────────

    /// <summary>
    /// Width returns the column width for a single-span cell.
    /// </summary>
    [Fact]
    public void Width_ReturnsSingleColumnWidth()
    {
        IColumn[] columns = [new ColumnStub { Width = 200 }];
        var cell = CreateCell(columns: columns);

        cell.Width.Should().Be(200);
    }

    /// <summary>
    /// Width accounts for column span.
    /// </summary>
    [Fact]
    public void Width_AccountsForColSpan()
    {
        var tc = MakeTcElement(gridSpan: "2");
        IColumn[] columns =
        [
            new ColumnStub { Width = 100 },
            new ColumnStub { Width = 200 },
            new ColumnStub { Width = 300 },
        ];
        var cell = CreateCell(tc, colIndex: 0, columns: columns);

        cell.Width.Should().Be(300); // 100 + 200
    }

    /// <summary>
    /// Height returns the row height for a single-span cell.
    /// </summary>
    [Fact]
    public void Height_ReturnsSingleRowHeight()
    {
        IRow[] rows = [new RowStub { Height = 50 }];
        var cell = CreateCell(rows: rows);

        cell.Height.Should().Be(50);
    }

    /// <summary>
    /// Height accounts for row span.
    /// </summary>
    [Fact]
    public void Height_AccountsForRowSpan()
    {
        var tc = MakeTcElement(rowSpan: "2");
        IRow[] rows =
        [
            new RowStub { Height = 30, MinimalHeight = 15 },
            new RowStub { Height = 50, MinimalHeight = 25 },
        ];
        var cell = CreateCell(tc, rowIndex: 0, rows: rows);

        cell.Height.Should().Be(80); // 30 + 50
    }

    // ── MinimalHeight ─────────────────────────────────────────────────

    /// <summary>
    /// MinimalHeight returns the row's minimal height for a single-span cell.
    /// </summary>
    [Fact]
    public void MinimalHeight_ReturnsSingleRowMinimalHeight()
    {
        IRow[] rows = [new RowStub { Height = 50, MinimalHeight = 20 }];
        var cell = CreateCell(rows: rows);

        cell.MinimalHeight.Should().Be(20);
    }

    // ── OffsetX / OffsetY ─────────────────────────────────────────────

    /// <summary>
    /// OffsetX is zero for a cell in the first column.
    /// </summary>
    [Fact]
    public void OffsetX_ZeroForFirstColumn()
    {
        var cell = CreateCell(colIndex: 0);

        cell.OffsetX.Should().Be(0);
    }

    /// <summary>
    /// OffsetX sums preceding column widths.
    /// </summary>
    [Fact]
    public void OffsetX_SumsPrecedingColumnWidths()
    {
        IColumn[] columns =
        [
            new ColumnStub { Width = 100 },
            new ColumnStub { Width = 200 },
            new ColumnStub { Width = 300 },
        ];
        var cell = CreateCell(colIndex: 2, columns: columns);

        cell.OffsetX.Should().Be(300); // 100 + 200
    }

    /// <summary>
    /// OffsetY is zero for a cell in the first row.
    /// </summary>
    [Fact]
    public void OffsetY_ZeroForFirstRow()
    {
        var cell = CreateCell(rowIndex: 0);

        cell.OffsetY.Should().Be(0);
    }

    /// <summary>
    /// OffsetY sums preceding row heights.
    /// </summary>
    [Fact]
    public void OffsetY_SumsPrecedingRowHeights()
    {
        IRow[] rows =
        [
            new RowStub { Height = 30 },
            new RowStub { Height = 50 },
            new RowStub { Height = 70 },
        ];
        var cell = CreateCell(rowIndex: 2, rows: rows);

        cell.OffsetY.Should().Be(80); // 30 + 50
    }

    // ── FirstRowIndex / FirstColumnIndex ──────────────────────────────

    /// <summary>
    /// FirstRowIndex returns the row index passed during init.
    /// </summary>
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void FirstRowIndex_ReturnsInitializedValue(int rowIndex)
    {
        IRow[] rows =
        [
            new RowStub { Height = 40 },
            new RowStub { Height = 40 },
            new RowStub { Height = 40 },
        ];
        var cell = CreateCell(rowIndex: rowIndex, rows: rows);

        cell.FirstRowIndex.Should().Be(rowIndex);
    }

    /// <summary>
    /// FirstColumnIndex returns the column index passed during init.
    /// </summary>
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void FirstColumnIndex_ReturnsInitializedValue(int colIndex)
    {
        IColumn[] columns =
        [
            new ColumnStub { Width = 100 },
            new ColumnStub { Width = 100 },
            new ColumnStub { Width = 100 },
        ];
        var cell = CreateCell(colIndex: colIndex, columns: columns);

        cell.FirstColumnIndex.Should().Be(colIndex);
    }

    // ── FirstRow / FirstColumn ────────────────────────────────────────

    /// <summary>
    /// FirstRow returns the correct row from the table.
    /// </summary>
    [Fact]
    public void FirstRow_ReturnsCorrectRow()
    {
        var row0 = new RowStub { Height = 30 };
        var row1 = new RowStub { Height = 50 };
        IRow[] rows = [row0, row1];
        var cell = CreateCell(rowIndex: 1, rows: rows);

        cell.FirstRow.Should().BeSameAs(row1);
    }

    /// <summary>
    /// FirstColumn returns the correct column from the table.
    /// </summary>
    [Fact]
    public void FirstColumn_ReturnsCorrectColumn()
    {
        var col0 = new ColumnStub { Width = 100 };
        var col1 = new ColumnStub { Width = 200 };
        IColumn[] columns = [col0, col1];
        var cell = CreateCell(colIndex: 1, columns: columns);

        cell.FirstColumn.Should().BeSameAs(col1);
    }

    // ── Table ─────────────────────────────────────────────────────────

    /// <summary>
    /// Table property returns the parent table.
    /// </summary>
    [Fact]
    public void Table_ReturnsParentTable()
    {
        var cell = CreateCell();

        cell.Table.Should().NotBeNull();
    }

    // ── Margins ───────────────────────────────────────────────────────

    /// <summary>
    /// Default margins are set correctly (left/right = 91440 EMU, top/bottom = 45720 EMU).
    /// </summary>
    [Fact]
    public void Margins_HaveDefaults()
    {
        var cell = CreateCell();

        cell.MarginLeft.Should().BeApproximately(91440f / 12700f, 0.01f);
        cell.MarginRight.Should().BeApproximately(91440f / 12700f, 0.01f);
        cell.MarginTop.Should().BeApproximately(45720f / 12700f, 0.01f);
        cell.MarginBottom.Should().BeApproximately(45720f / 12700f, 0.01f);
    }

    /// <summary>
    /// Margins can be set and read back.
    /// </summary>
    [Fact]
    public void Margins_CanBeSetAndRetrieved()
    {
        var cell = CreateCell();

        cell.MarginLeft = 10;
        cell.MarginRight = 12;
        cell.MarginTop = 5;
        cell.MarginBottom = 8;

        cell.MarginLeft.Should().BeApproximately(10, 0.1f);
        cell.MarginRight.Should().BeApproximately(12, 0.1f);
        cell.MarginTop.Should().BeApproximately(5, 0.1f);
        cell.MarginBottom.Should().BeApproximately(8, 0.1f);
    }

    /// <summary>
    /// Margins read from XML tcPr attributes.
    /// </summary>
    [Fact]
    public void Margins_ReadFromXml()
    {
        var tc = MakeTcElement();
        tc.Add(new XElement(ANs + "tcPr",
            new XAttribute("marL", "127000"),  // 10 points
            new XAttribute("marR", "254000"),  // 20 points
            new XAttribute("marT", "63500"),   // 5 points
            new XAttribute("marB", "0")));     // 0 points
        var cell = CreateCell(tc);

        cell.MarginLeft.Should().BeApproximately(10, 0.01f);
        cell.MarginRight.Should().BeApproximately(20, 0.01f);
        cell.MarginTop.Should().BeApproximately(5, 0.01f);
        cell.MarginBottom.Should().BeApproximately(0, 0.01f);
    }

    // ── TextVerticalType ──────────────────────────────────────────────

    /// <summary>
    /// TextVerticalType defaults to Horizontal.
    /// </summary>
    [Fact]
    public void TextVerticalType_DefaultsToHorizontal()
    {
        var cell = CreateCell();

        cell.TextVerticalType.Should().Be(TextVerticalType.Horizontal);
    }

    /// <summary>
    /// TextVerticalType reads from XML vert attribute.
    /// </summary>
    [Theory]
    [InlineData("horz", TextVerticalType.Horizontal)]
    [InlineData("vert", TextVerticalType.Vertical)]
    [InlineData("vert270", TextVerticalType.Vertical270)]
    public void TextVerticalType_ReadsFromXml(string xmlValue, TextVerticalType expected)
    {
        var tc = MakeTcElement();
        tc.Add(new XElement(ANs + "tcPr", new XAttribute("vert", xmlValue)));
        var cell = CreateCell(tc);

        cell.TextVerticalType.Should().Be(expected);
    }

    /// <summary>
    /// TextVerticalType can be set and read back.
    /// </summary>
    [Fact]
    public void TextVerticalType_CanBeSetAndRetrieved()
    {
        var cell = CreateCell();

        cell.TextVerticalType = TextVerticalType.Vertical;

        cell.TextVerticalType.Should().Be(TextVerticalType.Vertical);
    }

    // ── TextAnchorType ────────────────────────────────────────────────

    /// <summary>
    /// TextAnchorType defaults to Top.
    /// </summary>
    [Fact]
    public void TextAnchorType_DefaultsToTop()
    {
        var cell = CreateCell();

        cell.TextAnchorType.Should().Be(TextAnchorType.Top);
    }

    /// <summary>
    /// TextAnchorType reads from XML anchor attribute.
    /// </summary>
    [Theory]
    [InlineData("t", TextAnchorType.Top)]
    [InlineData("ctr", TextAnchorType.Center)]
    [InlineData("b", TextAnchorType.Bottom)]
    public void TextAnchorType_ReadsFromXml(string xmlValue, TextAnchorType expected)
    {
        var tc = MakeTcElement();
        tc.Add(new XElement(ANs + "tcPr", new XAttribute("anchor", xmlValue)));
        var cell = CreateCell(tc);

        cell.TextAnchorType.Should().Be(expected);
    }

    /// <summary>
    /// TextAnchorType can be set and read back.
    /// </summary>
    [Fact]
    public void TextAnchorType_CanBeSetAndRetrieved()
    {
        var cell = CreateCell();

        cell.TextAnchorType = TextAnchorType.Center;

        cell.TextAnchorType.Should().Be(TextAnchorType.Center);
    }

    // ── AnchorCenter ──────────────────────────────────────────────────

    /// <summary>
    /// AnchorCenter defaults to false.
    /// </summary>
    [Fact]
    public void AnchorCenter_DefaultsToFalse()
    {
        var cell = CreateCell();

        cell.AnchorCenter.Should().BeFalse();
    }

    /// <summary>
    /// AnchorCenter reads true from XML anchorCtr attribute.
    /// </summary>
    [Theory]
    [InlineData("1", true)]
    [InlineData("true", true)]
    [InlineData("0", false)]
    public void AnchorCenter_ReadsFromXml(string xmlValue, bool expected)
    {
        var tc = MakeTcElement();
        tc.Add(new XElement(ANs + "tcPr", new XAttribute("anchorCtr", xmlValue)));
        var cell = CreateCell(tc);

        cell.AnchorCenter.Should().Be(expected);
    }

    /// <summary>
    /// AnchorCenter can be set and read back.
    /// </summary>
    [Fact]
    public void AnchorCenter_CanBeSetAndRetrieved()
    {
        var cell = CreateCell();

        cell.AnchorCenter = true;

        cell.AnchorCenter.Should().BeTrue();
    }

    /// <summary>
    /// AnchorCenter can be toggled back to false.
    /// </summary>
    [Fact]
    public void AnchorCenter_CanBeToggledBackToFalse()
    {
        var cell = CreateCell();

        cell.AnchorCenter = true;
        cell.AnchorCenter = false;

        cell.AnchorCenter.Should().BeFalse();
    }

    // ── Component hierarchy ───────────────────────────────────────────

    /// <summary>
    /// AsISlideComponent returns the cell itself.
    /// </summary>
    [Fact]
    public void AsISlideComponent_ReturnsSelf()
    {
        var cell = CreateCell();

        cell.AsISlideComponent.Should().BeSameAs(cell);
    }

    /// <summary>
    /// AsIPresentationComponent returns the cell itself.
    /// </summary>
    [Fact]
    public void AsIPresentationComponent_ReturnsSelf()
    {
        var cell = CreateCell();

        cell.AsIPresentationComponent.Should().BeSameAs(cell);
    }

    /// <summary>
    /// Slide returns the parent slide set during init.
    /// </summary>
    [Fact]
    public void Slide_ReturnsParentSlide()
    {
        var slide = new BaseSlideStub();
        var cell = CreateCell(parentSlide: slide);

        cell.Slide.Should().BeSameAs(slide);
    }

    /// <summary>
    /// Slide is null when no parent slide is set.
    /// </summary>
    [Fact]
    public void Slide_NullWhenNoParentSlide()
    {
        var cell = CreateCell();

        cell.Slide.Should().BeNull();
    }

    /// <summary>
    /// Presentation returns the presentation from the parent slide.
    /// </summary>
    [Fact]
    public void Presentation_ReturnsPresentationFromSlide()
    {
        var pres = new PresentationStub();
        var slide = new BaseSlideStub { Presentation = pres };
        var cell = CreateCell(parentSlide: slide);

        cell.Presentation.Should().BeSameAs(pres);
    }

    /// <summary>
    /// Presentation is null when no parent slide.
    /// </summary>
    [Fact]
    public void Presentation_NullWhenNoParentSlide()
    {
        var cell = CreateCell();

        cell.Presentation.Should().BeNull();
    }

    // ── Null element edge case ────────────────────────────────────────

    /// <summary>
    /// ColSpan defaults to 1 with null element.
    /// </summary>
    [Fact]
    public void ColSpan_DefaultsToOneWithNullElement()
    {
        IColumn[] columns = [new ColumnStub { Width = 100 }];
        IRow[] rows = [new RowStub { Height = 40 }];
        var table = new TableStub { Rows = new RowCollectionStub(rows), Columns = new ColumnCollectionStub(columns) };

        var cell = new Cell();
        cell.InitInternal(null, 0, 0, slidePart: null, parentSlide: null, table);

        cell.ColSpan.Should().Be(1);
    }

    /// <summary>
    /// RowSpan defaults to 1 with null element.
    /// </summary>
    [Fact]
    public void RowSpan_DefaultsToOneWithNullElement()
    {
        IColumn[] columns = [new ColumnStub { Width = 100 }];
        IRow[] rows = [new RowStub { Height = 40 }];
        var table = new TableStub { Rows = new RowCollectionStub(rows), Columns = new ColumnCollectionStub(columns) };

        var cell = new Cell();
        cell.InitInternal(null, 0, 0, slidePart: null, parentSlide: null, table);

        cell.RowSpan.Should().Be(1);
    }

    /// <summary>
    /// IsMergedCell is false with null element.
    /// </summary>
    [Fact]
    public void IsMergedCell_FalseWithNullElement()
    {
        IColumn[] columns = [new ColumnStub { Width = 100 }];
        IRow[] rows = [new RowStub { Height = 40 }];
        var table = new TableStub { Rows = new RowCollectionStub(rows), Columns = new ColumnCollectionStub(columns) };

        var cell = new Cell();
        cell.InitInternal(null, 0, 0, slidePart: null, parentSlide: null, table);

        cell.IsMergedCell.Should().BeFalse();
    }

    /// <summary>
    /// TextFrame is null with null element.
    /// </summary>
    [Fact]
    public void TextFrame_NullWithNullElement()
    {
        IColumn[] columns = [new ColumnStub { Width = 100 }];
        IRow[] rows = [new RowStub { Height = 40 }];
        var table = new TableStub { Rows = new RowCollectionStub(rows), Columns = new ColumnCollectionStub(columns) };

        var cell = new Cell();
        cell.InitInternal(null, 0, 0, slidePart: null, parentSlide: null, table);

        cell.TextFrame.Should().BeNull();
    }
}
