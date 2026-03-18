using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a table row as a collection of cells.
/// </summary>
public sealed class Row : CellCollection, IRow, IBulkTextFormattable
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const float EmuPerPoint = 12700f;

    private XElement? _trElement;
    private SlidePart? _slidePart;

    /// <summary>
    /// Gets the underlying table row XML element.
    /// </summary>
    internal XElement? TrElement => _trElement;

    /// <summary>
    /// Gets the zero-based index of this row within the table.
    /// </summary>
    internal int RowIndex { get; private set; }

    /// <summary>
    /// Initializes internal state for the row with pre-built cells.
    /// </summary>
    internal void InitInternal(XElement trElement, int rowIndex, List<Cell> cells, SlidePart? slidePart, IBaseSlide? parentSlide)
    {
        _trElement = trElement;
        _slidePart = slidePart;
        RowIndex = rowIndex;
        base.InitInternal(cells, slidePart, parentSlide);
    }

    /// <summary>
    /// Initializes internal state for the row by parsing cells from the row XML element.
    /// </summary>
    /// <param name="trElement">The table row XML element containing <c>&lt;a:tc&gt;</c> children.</param>
    /// <param name="rowIndex">The zero-based row index within the table.</param>
    /// <param name="slidePart">The slide part for persistence.</param>
    /// <param name="parentSlide">The parent slide.</param>
    /// <param name="table">The owning table.</param>
    /// <returns>This row instance for fluent initialization.</returns>
    internal Row InitInternal(XElement trElement, int rowIndex, SlidePart? slidePart, IBaseSlide? parentSlide, ITable table)
    {
        _trElement = trElement;
        _slidePart = slidePart;
        RowIndex = rowIndex;

        // Parse cells from <a:tc> children
        var cells = new List<Cell>();
        var colIdx = 0;
        foreach (var tc in trElement.Elements(ANs + "tc"))
        {
            var cell = new Cell();
            cell.InitInternal(tc, rowIndex, colIdx, slidePart, parentSlide, table);
            cells.Add(cell);
            colIdx++;
        }

        base.InitInternal(cells, slidePart, parentSlide);
        return this;
    }

    /// <summary>
    /// Gets the actual computed height of the row in points.
    /// </summary>
    /// <remarks>
    /// The value is stored as EMUs in the underlying XML <c>h</c> attribute.
    /// Returns <c>0</c> if the backing XML element or its <c>h</c> attribute is missing.
    /// </remarks>
    public float Height
    {
        get
        {
            var h = _trElement?.Attribute("h")?.Value;
            return h is not null && long.TryParse(h, out var v) ? v / EmuPerPoint : 0f;
        }
    }

    /// <summary>
    /// Gets or sets the minimal possible height of the row in points.
    /// The row may render taller but never shorter than this value.
    /// </summary>
    /// <remarks>
    /// The value is stored as EMUs in the underlying XML <c>h</c> attribute.
    /// Setting the height persists the change via the slide part.
    /// </remarks>
    public float MinimalHeight
    {
        get
        {
            var h = _trElement?.Attribute("h")?.Value;
            return h is not null && long.TryParse(h, out var v) ? v / EmuPerPoint : 0f;
        }
        set
        {
            if (_trElement is null) return;
            _trElement.SetAttributeValue("h", (long)Math.Round(value * EmuPerPoint));
            _slidePart?.Save();
        }
    }

    /// <summary>
    /// Gets the formatting properties for this row.
    /// </summary>
    /// <remarks>
    /// Returns a new default <see cref="RowFormat"/> instance on every access.
    /// </remarks>
    public IRowFormat RowFormat => new RowFormat();

    /// <summary>
    /// Gets this row as an <see cref="ICellCollection"/>.
    /// </summary>
    public ICellCollection AsICellCollection => this;

    /// <summary>
    /// Gets this row as an <see cref="IBulkTextFormattable"/>.
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
