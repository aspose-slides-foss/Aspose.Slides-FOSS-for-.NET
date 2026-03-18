namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a collection of master slides.
/// </summary>
public interface IMasterSlideCollection
{
    /// <summary>
    /// Gets the master slide at the specified zero-based index.
    /// </summary>
    /// <param name="index">The zero-based index of the master slide.</param>
    /// <returns>The master slide at the specified index.</returns>
    IMasterSlide this[int index] { get; }

    /// <summary>
    /// Gets the number of master slides in the collection.
    /// </summary>
    int Count { get; }

    /// <summary>
    /// Exposes the collection as a generic list interface.
    /// </summary>
    IList<IMasterSlide> AsICollection { get; }

    /// <summary>
    /// Exposes the collection as an enumerable interface.
    /// </summary>
    IEnumerable<IMasterSlide> AsIEnumerable { get; }

    /// <summary>
    /// Adds a copy of a specified master slide to the end of the collection.
    /// </summary>
    /// <param name="sourceMaster">The master slide to clone.</param>
    /// <returns>The new cloned master slide.</returns>
    IMasterSlide AddClone(IMasterSlide sourceMaster);
}
