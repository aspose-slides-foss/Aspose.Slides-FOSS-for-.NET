namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a collection of all layout slides in presentation.
/// Extends <see cref="ILayoutSlideCollection"/> with methods for adding/cloning
/// layout slides in context of uniting of the individual collections of master's layout slides.
/// </summary>
public interface IGlobalLayoutSlideCollection : ILayoutSlideCollection
{
    /// <summary>
    /// Gets the layout slide at the specified index.
    /// </summary>
    /// <param name="index">The zero-based index of the layout slide.</param>
    /// <returns>The layout slide at the specified index.</returns>
    new ILayoutSlide this[int index] { get; }

    /// <summary>
    /// Gets the number of layout slides in the collection.
    /// </summary>
    int Count { get; }

    /// <summary>
    /// Returns <see cref="ILayoutSlideCollection"/> interface. Read-only.
    /// </summary>
    ILayoutSlideCollection AsILayoutSlideCollection { get; }
}
