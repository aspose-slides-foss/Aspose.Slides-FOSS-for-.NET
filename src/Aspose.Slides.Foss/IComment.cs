using Aspose.Slides.Foss.Drawing;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a comment on a slide.
/// </summary>
public interface IComment
{
    /// <summary>
    /// Returns or sets the plain text of a slide comment.
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// Returns or sets the time of a comment creation.
    /// </summary>
    DateTime? CreatedTime { get; set; }

    /// <summary>
    /// Returns the parent slide of a comment. Read-only.
    /// </summary>
    ISlide Slide { get; }

    /// <summary>
    /// Returns the author of a comment. Read-only.
    /// </summary>
    ICommentAuthor Author { get; }

    /// <summary>
    /// Returns or sets the position of a comment on a slide.
    /// </summary>
    PointF Position { get; set; }

    /// <summary>
    /// Gets or sets parent comment.
    /// </summary>
    IComment? ParentComment { get; set; }

    /// <summary>
    /// Removes comment and all its replies from the parent collection.
    /// </summary>
    void Remove();
}
