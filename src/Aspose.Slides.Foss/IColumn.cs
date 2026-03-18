namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a single column in a table.
/// Combines cell-collection behavior with bulk text formatting capability.
/// </summary>
public interface IColumn : ICellCollection, IBulkTextFormattable
{
    /// <summary>
    /// Gets or sets the width of the column.
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// Gets the formatting properties associated with this column.
    /// </summary>
    IColumnFormat ColumnFormat { get; }

    /// <summary>
    /// Returns this instance as an <see cref="ICellCollection"/>.
    /// </summary>
    ICellCollection AsICellCollection { get; }

    /// <summary>
    /// Returns this instance as an <see cref="IBulkTextFormattable"/>.
    /// </summary>
    IBulkTextFormattable AsIBulkTextFormattable { get; }
}
