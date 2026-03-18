namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a paragraph of text. Inherits from both
/// <see cref="ISlideComponent"/> and <see cref="IPresentationComponent"/>.
/// </summary>
public abstract class IParagraph : ISlideComponent
{
    /// <summary>
    /// Gets the collection of text portions (runs) that compose this paragraph.
    /// Each portion typically carries its own formatting.
    /// </summary>
    public abstract IPortionCollection Portions { get; }

    /// <summary>
    /// Gets the formatting object for this paragraph (e.g., alignment, indent, spacing).
    /// </summary>
    public abstract IParagraphFormat ParagraphFormat { get; }

    /// <summary>
    /// Gets or sets the plain text content of the paragraph.
    /// On get, returns the concatenation of all portion texts.
    /// On set, replaces the paragraph's content with the provided string value.
    /// </summary>
    public abstract string Text { get; set; }

    /// <summary>
    /// Returns this instance cast to the <see cref="ISlideComponent"/> base interface.
    /// </summary>
    public abstract ISlideComponent AsISlideComponent { get; }
}
