namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a picture reference within a slide.
/// </summary>
public abstract class ISlidesPicture : ISlideComponent
{
    /// <summary>
    /// Gets or sets the embedded image reference.
    /// </summary>
    public abstract IPPImage? Image { get; set; }

    /// <summary>
    /// Gets or sets the URL for a linked (external) image.
    /// </summary>
    public abstract string LinkPathLong { get; set; }

    /// <summary>
    /// Returns this instance as an <see cref="ISlideComponent"/>.
    /// </summary>
    public abstract ISlideComponent AsISlideComponent { get; }
}
