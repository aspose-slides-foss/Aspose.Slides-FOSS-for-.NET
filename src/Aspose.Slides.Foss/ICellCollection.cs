namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a collection of table cells.
/// </summary>
public interface ICellCollection
{
    /// <summary>
    /// Gets the cell at the specified index.
    /// </summary>
    ICell this[int index] { get; }

    /// <summary>
    /// Gets the number of cells in the collection.
    /// </summary>
    int Count { get; }

    /// <summary>
    /// Gets the slide that owns this collection.
    /// </summary>
    IBaseSlide? Slide { get; }

    /// <summary>
    /// Gets the presentation that owns this collection.
    /// </summary>
    IPresentation? Presentation { get; }

    /// <summary>
    /// Returns a new list copy of the cells in this collection.
    /// </summary>
    IList<ICell> AsICollection { get; }

    /// <summary>
    /// Returns a fresh enumerable over the cells in this collection.
    /// </summary>
    IEnumerable<ICell> AsIEnumerable { get; }

    /// <summary>
    /// Returns this instance as an <see cref="ISlideComponent"/>.
    /// </summary>
    ISlideComponent AsISlideComponent { get; }

    /// <summary>
    /// Returns this instance as an <see cref="IPresentationComponent"/>.
    /// </summary>
    IPresentationComponent AsIPresentationComponent { get; }
}
