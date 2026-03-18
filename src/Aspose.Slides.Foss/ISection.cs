namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a section of slides in a presentation.
/// </summary>
public interface ISection
{
    /// <summary>
    /// Gets or sets the name of the section.
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// Gets the unique identifier of the section.
    /// </summary>
    Guid SectionId { get; }

    /// <summary>
    /// Gets the first slide of the section.
    /// </summary>
    ISlide? StartedFromSlide { get; }

    /// <summary>
    /// Returns the list of slides belonging to this section.
    /// </summary>
    /// <returns>A list of slides in this section.</returns>
    IList<ISlide> GetSlidesListOfSection();
}
