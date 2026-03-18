namespace Aspose.Slides.Foss;

/// <summary>
/// Represents the formatting of a table cell.
/// </summary>
public interface ICellFormat
{
    /// <summary>
    /// Gets the fill format for the cell.
    /// </summary>
    IFillFormat FillFormat { get; }

    /// <summary>
    /// Gets the left border line format.
    /// </summary>
    ILineFormat BorderLeft { get; }

    /// <summary>
    /// Gets the top border line format.
    /// </summary>
    ILineFormat BorderTop { get; }

    /// <summary>
    /// Gets the right border line format.
    /// </summary>
    ILineFormat BorderRight { get; }

    /// <summary>
    /// Gets the bottom border line format.
    /// </summary>
    ILineFormat BorderBottom { get; }

    /// <summary>
    /// Gets the diagonal-down border line format.
    /// </summary>
    ILineFormat BorderDiagonalDown { get; }

    /// <summary>
    /// Gets the diagonal-up border line format.
    /// </summary>
    ILineFormat BorderDiagonalUp { get; }
}
