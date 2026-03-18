namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a row in a table. Combines cell collection and bulk text formatting capabilities.
/// </summary>
public interface IRow : ICellCollection, IBulkTextFormattable
{
    /// <summary>
    /// Gets the actual computed height of the row in points.
    /// </summary>
    float Height { get; }

    /// <summary>
    /// Gets or sets the minimal possible height of the row in points.
    /// The row may render taller but never shorter than this value.
    /// </summary>
    float MinimalHeight { get; set; }

    /// <summary>
    /// Gets the formatting properties for this row.
    /// </summary>
    IRowFormat RowFormat { get; }

    /// <summary>
    /// Returns this instance as an <see cref="ICellCollection"/>.
    /// </summary>
    ICellCollection AsICellCollection { get; }

    /// <summary>
    /// Returns this instance as an <see cref="IBulkTextFormattable"/>.
    /// </summary>
    IBulkTextFormattable AsIBulkTextFormattable { get; }
}
