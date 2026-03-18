namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a section of slides in a presentation.
/// </summary>
public sealed class Section : ISection
{
    /// <inheritdoc />
    public string Name { get; set; } = string.Empty;

    /// <inheritdoc />
    public Guid SectionId { get; } = Guid.NewGuid();

    /// <inheritdoc />
    public ISlide? StartedFromSlide { get; internal set; }

    /// <inheritdoc />
    public IList<ISlide> GetSlidesListOfSection() => [];
}
