namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a heading pair entry describing a content grouping in a presentation.
/// </summary>
public interface IHeadingPair
{
    /// <summary>
    /// Gets the name of the heading pair (e.g., "Slide Titles").
    /// </summary>
    string Name { get; }

    /// <summary>
    /// Gets the count of items associated with this heading.
    /// </summary>
    int Count { get; }
}
