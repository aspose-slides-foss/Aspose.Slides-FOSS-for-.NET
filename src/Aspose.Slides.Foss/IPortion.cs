namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a portion (run) of text inside a paragraph.
/// </summary>
public abstract class IPortion : ISlideComponent
{
    /// <summary>
    /// Gets the formatting object for this text portion.
    /// </summary>
    public abstract IBasePortionFormat? PortionFormat { get; }

    /// <summary>
    /// Gets or sets the plain text of this portion.
    /// </summary>
    public abstract string Text { get; set; }

    /// <summary>
    /// Returns this instance cast to <see cref="ISlideComponent"/>.
    /// </summary>
    public abstract ISlideComponent AsISlideComponent { get; }
}
