using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Manages visibility and text content of header, footer, date-time, and slide number
/// placeholders on a notes slide.
/// </summary>
public sealed class NotesSlideHeaderFooterManager : BaseHandoutNotesSlideHeaderFooterManager, INotesSlideHeaderFooterManager
{
    private NotesSlidePart _notesPart = null!;

    /// <summary>
    /// Initializes the manager with the specified notes slide part.
    /// </summary>
    /// <param name="notesPart">The notes slide part that holds placeholder state.</param>
    internal void InitInternal(NotesSlidePart notesPart)
    {
        _notesPart = notesPart;
    }

    /// <summary>
    /// Gets a value indicating whether the footer placeholder is visible on the notes slide.
    /// </summary>
    public bool IsFooterVisible => _notesPart.HasPlaceholder("ftr");

    /// <summary>
    /// Gets a value indicating whether the date-time placeholder is visible on the notes slide.
    /// </summary>
    public bool IsDateTimeVisible => _notesPart.HasPlaceholder("dt");

    /// <summary>
    /// Gets a value indicating whether the header placeholder is visible on the notes slide.
    /// </summary>
    public bool IsHeaderVisible => _notesPart.HasPlaceholder("hdr");

    /// <summary>
    /// Gets a value indicating whether the slide number placeholder is visible on the notes slide.
    /// </summary>
    public bool IsSlideNumberVisible => _notesPart.HasPlaceholder("sldNum");

    /// <summary>
    /// Gets this instance cast as <see cref="IBaseHandoutNotesSlideHeaderFooterManager"/>.
    /// </summary>
    public IBaseHandoutNotesSlideHeaderFooterManager AsIBaseHandoutNotesSlideHeaderFooterManag => this;

    /// <summary>
    /// Gets this instance cast as <see cref="IBaseSlideHeaderFooterManager"/>.
    /// </summary>
    public IBaseSlideHeaderFooterManager AsIBaseSlideHeaderFooterManager => this;

    /// <summary>
    /// Gets this instance cast as <see cref="IBaseHeaderFooterManager"/>.
    /// </summary>
    public IBaseHeaderFooterManager AsIBaseHeaderFooterManager => this;

    /// <summary>
    /// Sets the visibility of the footer placeholder on the notes slide.
    /// </summary>
    /// <param name="isVisible"><c>true</c> to make the footer visible; <c>false</c> to remove it.</param>
    public void SetFooterVisibility(bool isVisible)
    {
        if (isVisible)
            _notesPart.AddPlaceholder("ftr");
        else
            _notesPart.RemovePlaceholder("ftr");
    }

    /// <summary>
    /// Sets the visibility of the date-time placeholder on the notes slide.
    /// </summary>
    /// <param name="isVisible"><c>true</c> to make the date-time visible; <c>false</c> to remove it.</param>
    public void SetDateTimeVisibility(bool isVisible)
    {
        if (isVisible)
            _notesPart.AddPlaceholder("dt");
        else
            _notesPart.RemovePlaceholder("dt");
    }

    /// <summary>
    /// Sets the visibility of the header placeholder on the notes slide.
    /// </summary>
    /// <param name="isVisible"><c>true</c> to make the header visible; <c>false</c> to remove it.</param>
    public void SetHeaderVisibility(bool isVisible)
    {
        if (isVisible)
            _notesPart.AddPlaceholder("hdr");
        else
            _notesPart.RemovePlaceholder("hdr");
    }

    /// <summary>
    /// Sets the visibility of the slide number placeholder on the notes slide.
    /// </summary>
    /// <param name="isVisible"><c>true</c> to make the slide number visible; <c>false</c> to remove it.</param>
    public void SetSlideNumberVisibility(bool isVisible)
    {
        if (isVisible)
            _notesPart.AddPlaceholder("sldNum");
        else
            _notesPart.RemovePlaceholder("sldNum");
    }

    /// <summary>
    /// Sets the text content of the footer placeholder.
    /// </summary>
    /// <param name="text">The text to assign to the footer placeholder.</param>
    public void SetFooterText(string text)
    {
        _notesPart.SetPlaceholderText("ftr", text);
    }

    /// <summary>
    /// Sets the text content of the date-time placeholder.
    /// </summary>
    /// <param name="text">The text to assign to the date-time placeholder.</param>
    public void SetDateTimeText(string text)
    {
        _notesPart.SetPlaceholderText("dt", text);
    }

    /// <summary>
    /// Sets the text content of the header placeholder.
    /// </summary>
    /// <param name="text">The text to assign to the header placeholder.</param>
    public void SetHeaderText(string text)
    {
        _notesPart.SetPlaceholderText("hdr", text);
    }
}
