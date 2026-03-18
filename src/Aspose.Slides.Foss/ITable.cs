namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a table on a slide.
/// </summary>
public interface ITable : IGraphicalObject, IBulkTextFormattable
{
    /// <summary>
    /// Gets the collection of rows in the table.
    /// </summary>
    IRowCollection Rows { get; }

    /// <summary>
    /// Gets the collection of columns in the table.
    /// </summary>
    IColumnCollection Columns { get; }

    /// <summary>
    /// Gets the formatting properties object for this table.
    /// </summary>
    ITableFormat TableFormat { get; }

    /// <summary>
    /// Gets or sets the built-in table style preset.
    /// </summary>
    TableStylePreset StylePreset { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the table has right-to-left reading order.
    /// </summary>
    bool RightToLeft { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the first row should be drawn with special formatting (header row styling).
    /// </summary>
    bool FirstRow { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the first column should be drawn with special formatting.
    /// </summary>
    bool FirstCol { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the last row should be drawn with special formatting (total row styling).
    /// </summary>
    bool LastRow { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the last column should be drawn with special formatting.
    /// </summary>
    bool LastCol { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether even rows should be drawn with alternating formatting (row striping).
    /// </summary>
    bool HorizontalBanding { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether even columns should be drawn with alternating formatting (column striping).
    /// </summary>
    bool VerticalBanding { get; set; }

    /// <summary>
    /// Returns a reference to the base <see cref="IGraphicalObject"/> interface for this table.
    /// </summary>
    IGraphicalObject AsIGraphicalObject { get; }

    /// <summary>
    /// Returns a reference to the base <see cref="IBulkTextFormattable"/> interface for this table.
    /// </summary>
    IBulkTextFormattable AsIBulkTextFormattable { get; }

    /// <summary>
    /// Merges the rectangular range of cells defined by <paramref name="cell1"/> and <paramref name="cell2"/>.
    /// </summary>
    /// <param name="cell1">First cell defining one corner of the merge range.</param>
    /// <param name="cell2">Second cell defining the opposite corner of the merge range.</param>
    /// <param name="allowSplitting">When <c>true</c>, existing merged cells within the range may be split to accommodate the new merge operation.</param>
    /// <returns>The resulting merged cell.</returns>
    ICell MergeCells(ICell cell1, ICell cell2, bool allowSplitting);
}
