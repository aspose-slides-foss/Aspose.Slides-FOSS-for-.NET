namespace Aspose.Slides.Foss;

/// <summary>
/// Represents the text frame of a shape or cell.
/// </summary>
public interface ITextFrame
{
    /// <summary>
    /// Gets the collection of paragraphs in this text frame. Read-only.
    /// </summary>
    IParagraphCollection Paragraphs { get; }

    /// <summary>
    /// Gets or sets the plain text for this text frame.
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// Gets the formatting object for this text frame. Read-only.
    /// </summary>
    ITextFrameFormat TextFrameFormat { get; }

    /// <summary>
    /// Gets the parent shape, or <c>null</c> if the parent does not implement <see cref="IShape"/>.
    /// </summary>
    IShape? ParentShape { get; }

    /// <summary>
    /// Gets the parent cell, or <c>null</c> if the parent does not implement <see cref="ICell"/>.
    /// </summary>
    ICell? ParentCell { get; }

    /// <summary>
    /// Returns this instance viewed as an <see cref="ISlideComponent"/>. Read-only.
    /// </summary>
    ISlideComponent AsISlideComponent { get; }
}
