namespace Aspose.Slides.Foss.Internal.Pptx.PresentationPart;

/// <summary>
/// Reference to a master slide in the presentation, holding its unique ID and relationship ID.
/// </summary>
public sealed class MasterReference(int masterId, string relId)
{
    /// <summary>
    /// Gets or sets the unique master slide ID (id attribute).
    /// </summary>
    public int MasterId { get; set; } = masterId;

    /// <summary>
    /// Gets or sets the relationship ID (r:id attribute).
    /// </summary>
    public string RelId { get; set; } = relId;
}
