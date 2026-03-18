namespace Aspose.Slides.Foss;

/// <summary>
/// Represents the metadata properties of a presentation document.
/// </summary>
public interface IDocumentProperties
{
    // ── Core properties ───────────────────────────────────────

    /// <summary>Gets or sets the title of the presentation.</summary>
    string Title { get; set; }

    /// <summary>Gets or sets the subject of the presentation.</summary>
    string Subject { get; set; }

    /// <summary>Gets or sets the author (dc:creator) of the presentation.</summary>
    string Author { get; set; }

    /// <summary>Gets or sets the keywords associated with the presentation.</summary>
    string Keywords { get; set; }

    /// <summary>Gets or sets the comments (dc:description) for the presentation.</summary>
    string Comments { get; set; }

    /// <summary>Gets or sets the category of the presentation.</summary>
    string Category { get; set; }

    /// <summary>Gets or sets the content status of the presentation.</summary>
    string ContentStatus { get; set; }

    /// <summary>Gets or sets the content type of the presentation.</summary>
    string ContentType { get; set; }

    /// <summary>Gets or sets the name of the last person who saved the presentation.</summary>
    string LastSavedBy { get; set; }

    /// <summary>Gets or sets the revision number of the presentation.</summary>
    int RevisionNumber { get; set; }

    /// <summary>Gets or sets the creation date and time in UTC.</summary>
    DateTime? CreatedTime { get; set; }

    /// <summary>Gets or sets the last saved date and time in UTC.</summary>
    DateTime? LastSavedTime { get; set; }

    /// <summary>Gets or sets the last printed date and time in UTC.</summary>
    DateTime? LastPrinted { get; set; }

    // ── App properties ────────────────────────────────────────

    /// <summary>Gets the application version string.</summary>
    string AppVersion { get; }

    /// <summary>Gets or sets the name of the application that created the presentation.</summary>
    string NameOfApplication { get; set; }

    /// <summary>Gets or sets the company name.</summary>
    string Company { get; set; }

    /// <summary>Gets or sets the manager name.</summary>
    string Manager { get; set; }

    /// <summary>Gets or sets the presentation format description.</summary>
    string PresentationFormat { get; set; }

    /// <summary>Gets or sets the application template name.</summary>
    string ApplicationTemplate { get; set; }

    /// <summary>Gets or sets the hyperlink base URI.</summary>
    string HyperlinkBase { get; set; }

    /// <summary>Gets or sets the total editing time.</summary>
    TimeSpan TotalEditingTime { get; set; }

    /// <summary>Gets or sets a value indicating whether the document is shared.</summary>
    bool SharedDoc { get; set; }

    /// <summary>Gets or sets a value indicating whether the thumbnail should be cropped/scaled.</summary>
    bool ScaleCrop { get; set; }

    /// <summary>Gets or sets a value indicating whether links are up to date.</summary>
    bool LinksUpToDate { get; set; }

    /// <summary>Gets or sets a value indicating whether hyperlinks have changed.</summary>
    bool HyperlinksChanged { get; set; }

    /// <summary>Gets the number of slides in the presentation.</summary>
    int Slides { get; }

    /// <summary>Gets the number of hidden slides in the presentation.</summary>
    int HiddenSlides { get; }

    /// <summary>Gets the number of notes pages in the presentation.</summary>
    int Notes { get; }

    /// <summary>Gets the paragraph count statistic.</summary>
    int Paragraphs { get; }

    /// <summary>Gets the word count statistic.</summary>
    int Words { get; }

    /// <summary>Gets the multimedia clip count statistic.</summary>
    int MultimediaClips { get; }

    /// <summary>Gets the heading pairs describing content groupings.</summary>
    IReadOnlyList<IHeadingPair> HeadingPairs { get; }

    /// <summary>Gets the titles of parts corresponding to heading pairs.</summary>
    IReadOnlyList<string> TitlesOfParts { get; }

    // ── Custom properties ─────────────────────────────────────

    /// <summary>Gets the number of custom properties.</summary>
    int CountOfCustomProperties { get; }

    /// <summary>Gets the value of a custom property by name.</summary>
    /// <param name="name">The property name.</param>
    /// <returns>The property value, or <c>null</c> if not found.</returns>
    object? GetCustomPropertyValue(string name);

    /// <summary>Sets or creates a custom property.</summary>
    /// <param name="name">The property name.</param>
    /// <param name="value">The property value. Supported types: <see cref="string"/>, <see cref="int"/>, <see cref="double"/>, <see cref="bool"/>, <see cref="DateTime"/>.</param>
    void SetCustomPropertyValue(string name, object value);

    /// <summary>Gets the name of a custom property by its zero-based index.</summary>
    /// <param name="index">The zero-based index.</param>
    /// <returns>The property name.</returns>
    string GetCustomPropertyName(int index);

    /// <summary>Removes a custom property by name.</summary>
    /// <param name="name">The property name.</param>
    /// <returns><c>true</c> if the property existed and was removed; otherwise, <c>false</c>.</returns>
    bool RemoveCustomProperty(string name);

    /// <summary>Determines whether a custom property with the specified name exists.</summary>
    /// <param name="name">The property name.</param>
    /// <returns><c>true</c> if the property exists; otherwise, <c>false</c>.</returns>
    bool ContainsCustomProperty(string name);

    /// <summary>Removes all custom properties.</summary>
    void ClearCustomProperties();

    /// <summary>Resets all built-in (core and app) properties to their default values.</summary>
    void ClearBuiltInProperties();
}
