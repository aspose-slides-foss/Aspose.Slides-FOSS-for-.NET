using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a table column as a collection of cells (one per row).
/// </summary>
public sealed class Column : CellCollection, IColumn, IBulkTextFormattable
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const float EmuPerPoint = 12700f;

    private XElement? _gridColElement;
    private SlidePart? _slidePart;
    private XElement? _tblElement;
    private ITable? _tableRef;

    /// <summary>
    /// Gets the underlying grid column XML element.
    /// </summary>
    internal XElement? GridColElement => _gridColElement;

    /// <summary>
    /// Gets the zero-based index of this column within the table.
    /// </summary>
    internal int ColIndex { get; private set; }

    /// <summary>
    /// Initializes internal state for the column.
    /// </summary>
    internal void InitInternal(XElement gridColElement, int colIndex, List<Cell> cells, SlidePart? slidePart, IBaseSlide? parentSlide)
    {
        _gridColElement = gridColElement;
        _slidePart = slidePart;
        ColIndex = colIndex;
        base.InitInternal(cells, slidePart, parentSlide);
    }

    /// <summary>
    /// Initializes internal state for the column by parsing cells from the table XML element.
    /// </summary>
    /// <param name="gridColElement">The grid column XML element.</param>
    /// <param name="colIndex">The zero-based column index.</param>
    /// <param name="tblElement">The table XML element containing rows.</param>
    /// <param name="slidePart">The slide part for persistence.</param>
    /// <param name="parentSlide">The parent slide.</param>
    /// <param name="table">The owning table.</param>
    /// <returns>This column instance for fluent initialization.</returns>
    internal Column InitInternal(XElement gridColElement, int colIndex, XElement tblElement, SlidePart? slidePart, IBaseSlide? parentSlide, ITable table)
    {
        _gridColElement = gridColElement;
        ColIndex = colIndex;
        _tblElement = tblElement;
        _slidePart = slidePart;
        _tableRef = table;

        // Parse cells: one per row at this column index
        var cells = new List<Cell>();
        var rowIdx = 0;
        foreach (var tr in tblElement.Elements(ANs + "tr"))
        {
            var tcs = tr.Elements(ANs + "tc").ToList();
            if (colIndex < tcs.Count)
            {
                var cell = new Cell();
                cell.InitInternal(tcs[colIndex], rowIdx, colIndex, slidePart, parentSlide, table);
                cells.Add(cell);
            }
            rowIdx++;
        }

        base.InitInternal(cells, slidePart, parentSlide);
        return this;
    }

    /// <summary>
    /// Gets or sets the width of the column in points.
    /// </summary>
    /// <remarks>
    /// The value is stored as EMUs in the underlying XML. Setting the width persists the change via the slide part.
    /// Returns <c>0</c> if the backing XML element or its <c>w</c> attribute is missing.
    /// </remarks>
    public float Width
    {
        get
        {
            var w = _gridColElement?.Attribute("w")?.Value;
            return w is not null && long.TryParse(w, out var v) ? v / EmuPerPoint : 0f;
        }
        set
        {
            if (_gridColElement is null) return;
            _gridColElement.SetAttributeValue("w", (long)Math.Round(value * EmuPerPoint));
            _slidePart?.Save();
        }
    }

    /// <summary>
    /// Gets the formatting properties for this column.
    /// </summary>
    /// <remarks>
    /// Returns a new default <see cref="ColumnFormat"/> instance on every access.
    /// </remarks>
    public IColumnFormat ColumnFormat => new ColumnFormat();

    /// <summary>
    /// Gets this column as an <see cref="ICellCollection"/>.
    /// </summary>
    public ICellCollection AsICellCollection => this;

    /// <summary>
    /// Gets this column as an <see cref="IBulkTextFormattable"/>.
    /// </summary>
    public IBulkTextFormattable AsIBulkTextFormattable => this;

    /// <inheritdoc/>
    public void SetTextFormat(IBasePortionFormat source)
    {
        ArgumentNullException.ThrowIfNull(source);
        TextFormatHelper.ApplyTextFormat(this.ToList(), source, _slidePart);
    }

    /// <inheritdoc/>
    public void SetTextFormat(IParagraphFormat source)
    {
        ArgumentNullException.ThrowIfNull(source);
        TextFormatHelper.ApplyTextFormat(this.ToList(), source, _slidePart);
    }

    /// <inheritdoc/>
    public void SetTextFormat(ITextFrameFormat source)
    {
        ArgumentNullException.ThrowIfNull(source);
        TextFormatHelper.ApplyTextFormat(this.ToList(), source, _slidePart);
    }
}
