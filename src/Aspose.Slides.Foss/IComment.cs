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
    /// Returns or sets the position of a comment on a slide, in centimetres from the top-left
    /// corner of the slide.
    /// </summary>
    /// <remarks>
    /// The file stores the position in eighths of a point, the unit PowerPoint writes and reads, so
    /// a value is kept to the nearest eighth of a point — within about 0.002 cm.
    /// </remarks>
    PointF Position { get; set; }

    /// <summary>
    /// Gets or sets the comment this one replies to, or <c>null</c> for a top-level comment.
    /// </summary>
    /// <remarks>
    /// A reply is written as PowerPoint writes one: the reply's classic <c>&lt;p:cm&gt;</c> carries a
    /// <c>p15:threadingInfo</c> extension whose <c>p15:parentCm</c> names the parent by author and
    /// index, and the parent carries an empty <c>p15:threadingInfo</c>. No
    /// <c>ppt/threadedComments/</c> part is written. A reply and its parent must be on the same
    /// slide. A reply whose parent is removed on its own is saved as a top-level comment;
    /// <see cref="Remove"/> removes a comment together with its replies.
    /// </remarks>
    /// <exception cref="ArgumentException">
    /// The value is this comment, a comment on another slide, or not a comment of this library.
    /// </exception>
    IComment? ParentComment { get; set; }

    /// <summary>
    /// Removes comment and all its replies from the parent collection.
    /// </summary>
    void Remove();
}
