namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a single cell in a table.
/// </summary>
public interface ICell
{
    /// <summary>
    /// Gets the horizontal offset of the cell from the left edge of the table, in points.
    /// </summary>
    float OffsetX { get; }

    /// <summary>
    /// Gets the vertical offset of the cell from the top edge of the table, in points.
    /// </summary>
    float OffsetY { get; }

    /// <summary>
    /// Gets the zero-based index of the first row covered by this cell.
    /// </summary>
    int FirstRowIndex { get; }

    /// <summary>
    /// Gets the zero-based index of the first column covered by this cell.
    /// </summary>
    int FirstColumnIndex { get; }

    /// <summary>
    /// Gets the total width of the cell in points, accounting for column span.
    /// </summary>
    float Width { get; }

    /// <summary>
    /// Gets the total height of the cell in points, accounting for row span.
    /// </summary>
    float Height { get; }

    /// <summary>
    /// Gets the minimal height of the cell in points.
    /// </summary>
    float MinimalHeight { get; }

    /// <summary>
    /// Gets or sets the left margin of the cell in points.
    /// </summary>
    float MarginLeft { get; set; }

    /// <summary>
    /// Gets or sets the right margin of the cell in points.
    /// </summary>
    float MarginRight { get; set; }

    /// <summary>
    /// Gets or sets the top margin of the cell in points.
    /// </summary>
    float MarginTop { get; set; }

    /// <summary>
    /// Gets or sets the bottom margin of the cell in points.
    /// </summary>
    float MarginBottom { get; set; }

    /// <summary>
    /// Gets or sets the text vertical type of the cell.
    /// </summary>
    TextVerticalType TextVerticalType { get; set; }

    /// <summary>
    /// Gets or sets the text anchor type of the cell.
    /// </summary>
    TextAnchorType TextAnchorType { get; set; }

    /// <summary>
    /// Gets or sets whether text is centered at the anchor point.
    /// </summary>
    bool AnchorCenter { get; set; }

    /// <summary>
    /// Gets the first row of this cell.
    /// </summary>
    IRow FirstRow { get; }

    /// <summary>
    /// Gets the first column of this cell.
    /// </summary>
    IColumn FirstColumn { get; }

    /// <summary>
    /// Gets the column span of this cell.
    /// </summary>
    int ColSpan { get; }

    /// <summary>
    /// Gets the row span of this cell.
    /// </summary>
    int RowSpan { get; }

    /// <summary>
    /// Gets the text frame of this cell.
    /// </summary>
    ITextFrame? TextFrame { get; }

    /// <summary>
    /// Gets the parent table.
    /// </summary>
    ITable Table { get; }

    /// <summary>
    /// Gets whether this cell is part of a merge operation.
    /// </summary>
    bool IsMergedCell { get; }

    /// <summary>
    /// Gets the cell format.
    /// </summary>
    ICellFormat CellFormat { get; }

    /// <summary>
    /// Gets the owning slide.
    /// </summary>
    IBaseSlide? Slide { get; }

    /// <summary>
    /// Gets the owning presentation.
    /// </summary>
    IPresentation? Presentation { get; }

    /// <summary>
    /// Returns this instance as an <see cref="ISlideComponent"/>.
    /// </summary>
    ISlideComponent AsISlideComponent { get; }

    /// <summary>
    /// Returns this instance as an <see cref="IPresentationComponent"/>.
    /// </summary>
    IPresentationComponent AsIPresentationComponent { get; }
}
