namespace Aspose.Slides.Foss;

/// <summary>
/// Represents any component that belongs to a presentation.
/// </summary>
public abstract class IPresentationComponent
{
    /// <summary>
    /// Gets the owning presentation.
    /// </summary>
    public abstract IPresentation? Presentation { get; }
}
