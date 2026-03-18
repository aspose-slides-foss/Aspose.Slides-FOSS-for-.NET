namespace Aspose.Slides.Foss.Internal.Pptx.PresentationPart;

/// <summary>
/// Reference to a slide in the presentation, holding its unique ID and relationship ID.
/// </summary>
public sealed class SlideReference(int slideId, string relId)
{
    /// <summary>
    /// Gets or sets the unique slide ID (id attribute).
    /// </summary>
    public int SlideId { get; set; } = slideId;

    /// <summary>
    /// Gets or sets the relationship ID (r:id attribute).
    /// </summary>
    public string RelId { get; set; } = relId;
}
