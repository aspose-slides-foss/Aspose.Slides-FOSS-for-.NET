namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a master slide in a presentation.
/// </summary>
public interface IMasterSlide : IBaseSlide
{
    /// <summary>
    /// Returns the collection of child layout slides for this master slide. Read-only.
    /// </summary>
    IMasterLayoutSlideCollection LayoutSlides { get; }
}
