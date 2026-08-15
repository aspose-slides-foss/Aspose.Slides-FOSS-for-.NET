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
    /// Gets or sets the comment this one replies to.
    /// </summary>
    /// <remarks>
    /// The reply relationship is held in memory and is not written to the file. PowerPoint expresses
    /// threads in a <c>ppt/threadedComments/</c> part, which this library does not write yet; the
    /// classic <c>&lt;p:cm&gt;</c> element has no attribute for a parent, so a saved deck carries
    /// flat comments. Setting this still governs <see cref="Remove"/>, which removes a comment
    /// together with its replies.
    /// </remarks>
    IComment? ParentComment { get; set; }

    /// <summary>
    /// Removes comment and all its replies from the parent collection.
    /// </summary>
    void Remove();
}
