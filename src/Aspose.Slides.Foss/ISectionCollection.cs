namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a collection of sections in a presentation.
/// </summary>
public interface ISectionCollection : IEnumerable<ISection>
{
    /// <summary>
    /// Gets the section at the specified zero-based index.
    /// </summary>
    /// <param name="index">The zero-based index of the section.</param>
    /// <returns>The section at the specified index.</returns>
    ISection this[int index] { get; }

    /// <summary>
    /// Gets the number of sections in the collection.
    /// </summary>
    int Count { get; }

    /// <summary>
    /// Adds a new section started from the specified slide.
    /// </summary>
    /// <param name="name">The name of the section.</param>
    /// <param name="startedFromSlide">The first slide of the section.</param>
    /// <returns>The newly added section.</returns>
    ISection AddSection(string name, ISlide startedFromSlide);

    /// <summary>
    /// Adds an empty section with the specified name to the end of the collection.
    /// </summary>
    /// <param name="name">The name of the section.</param>
    /// <param name="startedFromIndex">The starting index for the section.</param>
    /// <returns>The newly added section.</returns>
    ISection AppendEmptySection(string name, int startedFromIndex);

    /// <summary>
    /// Returns the index of the specified section, or -1 if not found.
    /// </summary>
    /// <param name="section">The section to locate.</param>
    /// <returns>The zero-based index, or -1 if the section is not in the collection.</returns>
    int IndexOf(ISection section);

    /// <summary>
    /// Removes the specified section from the collection.
    /// </summary>
    /// <param name="section">The section to remove.</param>
    void RemoveSection(ISection section);

    /// <summary>
    /// Removes the specified section and all its slides from the collection.
    /// </summary>
    /// <param name="section">The section to remove.</param>
    void RemoveSectionWithSlides(ISection section);

    /// <summary>
    /// Moves the section and its slides to the specified position.
    /// </summary>
    /// <param name="section">The section to move.</param>
    /// <param name="index">The target index.</param>
    void ReorderSectionWithSlides(ISection section, int index);

    /// <summary>
    /// Removes all sections from the collection.
    /// </summary>
    void Clear();
}
