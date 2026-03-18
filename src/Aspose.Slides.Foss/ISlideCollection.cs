namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a collection of slides in a presentation.
/// </summary>
public interface ISlideCollection : IEnumerable<ISlide>
{
    /// <summary>
    /// Gets the slide at the specified zero-based index.
    /// </summary>
    /// <param name="index">The zero-based index of the slide.</param>
    /// <returns>The slide at the specified index.</returns>
    ISlide this[int index] { get; }

    /// <summary>
    /// Gets the number of slides in the collection.
    /// </summary>
    int Count { get; }

    /// <summary>
    /// Exposes the collection as a generic list interface.
    /// </summary>
    IList<ISlide> AsICollection { get; }

    /// <summary>
    /// Exposes the collection as an enumerable interface.
    /// </summary>
    IEnumerable<ISlide> AsIEnumerable { get; }

    /// <summary>
    /// Adds a copy of a specified slide to the end of the collection.
    /// </summary>
    /// <param name="sourceSlide">The slide to clone.</param>
    /// <returns>The new cloned slide.</returns>
    ISlide AddClone(ISlide sourceSlide);

    /// <summary>
    /// Adds a copy of a specified slide to the end of the collection,
    /// applying the specified layout.
    /// </summary>
    /// <param name="sourceSlide">The slide to clone.</param>
    /// <param name="destLayout">The layout to apply to the cloned slide.</param>
    /// <returns>The new cloned slide.</returns>
    ISlide AddClone(ISlide sourceSlide, ILayoutSlide destLayout);

    /// <summary>
    /// Adds a copy of a specified slide to the end of the collection,
    /// applying the specified master slide.
    /// </summary>
    /// <param name="sourceSlide">The slide to clone.</param>
    /// <param name="destMaster">The master slide to apply.</param>
    /// <param name="allowCloneMissingLayout">
    /// If <c>true</c>, allows cloning even when the layout is not found in the destination master.
    /// </param>
    /// <returns>The new cloned slide.</returns>
    ISlide AddClone(ISlide sourceSlide, IMasterSlide destMaster, bool allowCloneMissingLayout);

    /// <summary>
    /// Inserts a copy of a specified slide at the specified position.
    /// </summary>
    /// <param name="index">The zero-based index at which to insert the cloned slide.</param>
    /// <param name="sourceSlide">The slide to clone.</param>
    /// <returns>The new cloned slide.</returns>
    ISlide InsertClone(int index, ISlide sourceSlide);

    /// <summary>
    /// Inserts a copy of a specified slide at the specified position,
    /// applying the specified layout.
    /// </summary>
    /// <param name="index">The zero-based index at which to insert the cloned slide.</param>
    /// <param name="sourceSlide">The slide to clone.</param>
    /// <param name="destLayout">The layout to apply to the cloned slide.</param>
    /// <returns>The new cloned slide.</returns>
    ISlide InsertClone(int index, ISlide sourceSlide, ILayoutSlide destLayout);

    /// <summary>
    /// Inserts a copy of a specified slide at the specified position,
    /// applying the specified master slide.
    /// </summary>
    /// <param name="index">The zero-based index at which to insert the cloned slide.</param>
    /// <param name="sourceSlide">The slide to clone.</param>
    /// <param name="destMaster">The master slide to apply.</param>
    /// <param name="allowCloneMissingLayout">
    /// If <c>true</c>, allows cloning even when the layout is not found in the destination master.
    /// </param>
    /// <returns>The new cloned slide.</returns>
    ISlide InsertClone(int index, ISlide sourceSlide, IMasterSlide destMaster, bool allowCloneMissingLayout);

    /// <summary>
    /// Returns all slides in the collection as a new array.
    /// </summary>
    /// <returns>An array containing all slides.</returns>
    ISlide[] ToArray();

    /// <summary>
    /// Returns a subset of slides starting at the specified index.
    /// </summary>
    /// <param name="startIndex">The zero-based start index.</param>
    /// <param name="count">The number of elements to return.</param>
    /// <returns>An array containing the specified range of slides.</returns>
    ISlide[] ToArray(int startIndex, int count);

    /// <summary>
    /// Adds an empty slide to the end of the collection with the specified layout.
    /// </summary>
    /// <param name="layout">The layout to apply to the new slide.</param>
    /// <returns>The newly added slide.</returns>
    ISlide AddEmptySlide(ILayoutSlide layout);

    /// <summary>
    /// Inserts an empty slide at the specified position with the specified layout.
    /// </summary>
    /// <param name="index">The zero-based index at which to insert the slide.</param>
    /// <param name="layout">The layout to apply to the new slide.</param>
    /// <returns>The newly inserted slide.</returns>
    ISlide InsertEmptySlide(int index, ILayoutSlide layout);

    /// <summary>
    /// Removes the first occurrence of the specified slide from the collection.
    /// </summary>
    /// <param name="value">The slide to remove.</param>
    void Remove(ISlide value);

    /// <summary>
    /// Removes the slide at the specified zero-based index.
    /// </summary>
    /// <param name="index">The zero-based index of the slide to remove.</param>
    void RemoveAt(int index);

    /// <summary>
    /// Returns the zero-based index of the specified slide, or -1 if not found.
    /// </summary>
    /// <param name="slide">The slide to locate.</param>
    /// <returns>The zero-based index, or -1 if the slide is not in the collection.</returns>
    int IndexOf(ISlide slide);
}
