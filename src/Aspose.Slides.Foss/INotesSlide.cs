namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a notes slide in a presentation.
/// </summary>
public interface INotesSlide : IBaseSlide
{
    /// <summary>
    /// Returns the HeaderFooter manager of the notes slide. Read-only.
    /// </summary>
    INotesSlideHeaderFooterManager HeaderFooterManager { get; }

    /// <summary>
    /// Returns a TextFrame with notes' text if there is one. Read-only.
    /// </summary>
    ITextFrame NotesTextFrame { get; }

    /// <summary>
    /// Returns the parent slide. Read-only.
    /// </summary>
    ISlide ParentSlide { get; }

    /// <summary>
    /// Allows to get the base IBaseSlide interface. Read-only.
    /// </summary>
    IBaseSlide AsIBaseSlide { get; }
}
