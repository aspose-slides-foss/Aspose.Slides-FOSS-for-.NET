namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a base slide.
/// </summary>
public interface IBaseSlide
{
    /// <summary>
    /// Gets the presentation that owns this slide.
    /// </summary>
    IPresentation? Presentation { get; }

    /// <summary>
    /// Gets the shapes of a slide. Read-only.
    /// </summary>
    IShapeCollection? Shapes { get; }

    /// <summary>
    /// Gets or sets the name of a slide.
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// Gets the ID of a slide. Read-only.
    /// </summary>
    int SlideId { get; }
}
