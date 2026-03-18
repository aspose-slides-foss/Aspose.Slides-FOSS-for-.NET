namespace Aspose.Slides.Foss;

/// <summary>
/// Manages notes slide operations for a slide.
/// </summary>
public interface INotesSlideManager
{
    /// <summary>
    /// Returns the notes slide for the associated slide, or <c>null</c> if none exists. Read-only.
    /// </summary>
    INotesSlide? NotesSlide { get; }

    /// <summary>
    /// Creates and adds a new notes slide for the associated slide.
    /// </summary>
    /// <returns>The newly created notes slide.</returns>
    INotesSlide AddNotesSlide();

    /// <summary>
    /// Removes the notes slide from the associated slide.
    /// </summary>
    void RemoveNotesSlide();
}
