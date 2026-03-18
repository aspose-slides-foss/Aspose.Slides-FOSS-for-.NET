namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a slide in a presentation.
/// </summary>
public interface ISlide : IBaseSlide
{
    /// <summary>
    /// Returns a number of slide.
    /// Index of slide in collection is always equal to SlideNumber - 1.
    /// Read/write.
    /// </summary>
    int SlideNumber { get; set; }

    /// <summary>
    /// Determines whether the specified slide is hidden during a slide show. Read/write.
    /// </summary>
    bool Hidden { get; set; }

    /// <summary>
    /// Returns or sets the layout slide for the current slide. Read/write.
    /// </summary>
    ILayoutSlide? LayoutSlide { get; set; }

    /// <summary>
    /// Allow to access notes slide, add and remove it. Read-only.
    /// </summary>
    INotesSlideManager NotesSlideManager { get; }

    /// <summary>
    /// Returns the comments for the specified author, or all comments if author is <c>null</c>.
    /// </summary>
    /// <param name="author">The author to filter by, or <c>null</c> for all comments.</param>
    /// <returns>A list of comments on this slide.</returns>
    List<IComment> GetSlideComments(ICommentAuthor? author);

    /// <summary>
    /// Removes the slide from the presentation.
    /// </summary>
    void Remove();
}
