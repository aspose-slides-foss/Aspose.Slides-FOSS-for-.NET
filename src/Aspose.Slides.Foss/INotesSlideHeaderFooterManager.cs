namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a manager for notes slide header and footer placeholders.
/// </summary>
public interface INotesSlideHeaderFooterManager : IBaseHandoutNotesSlideHeaderFooterManager
{
    /// <summary>
    /// Gets a value indicating whether the footer placeholder is visible on the notes slide.
    /// </summary>
    bool IsFooterVisible { get; }

    /// <summary>
    /// Gets a value indicating whether the slide number placeholder is visible on the notes slide.
    /// </summary>
    bool IsSlideNumberVisible { get; }

    /// <summary>
    /// Sets the visibility of the footer placeholder on the notes slide.
    /// </summary>
    /// <param name="isVisible"><c>true</c> to make the footer visible; <c>false</c> to remove it.</param>
    void SetFooterVisibility(bool isVisible);

    /// <summary>
    /// Sets the text content of the footer placeholder.
    /// </summary>
    /// <param name="text">The text to assign to the footer placeholder.</param>
    void SetFooterText(string text);

    /// <summary>
    /// Sets the visibility of the slide number placeholder on the notes slide.
    /// </summary>
    /// <param name="isVisible"><c>true</c> to make the slide number visible; <c>false</c> to remove it.</param>
    void SetSlideNumberVisibility(bool isVisible);
}
