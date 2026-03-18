namespace Aspose.Slides.Foss.Internal.Pptx.AppPropertiesPart;

/// <summary>
/// Internal representation of a heading pair entry in docProps/app.xml.
/// Each pair associates a category name (e.g. "Slide Titles") with a count of entries.
/// </summary>
public sealed class HeadingPairData(string name, int count)
{
    /// <summary>
    /// Gets or sets the heading pair name.
    /// </summary>
    public string Name { get; set; } = name;

    /// <summary>
    /// Gets or sets the count of entries for this heading pair.
    /// </summary>
    public int Count { get; set; } = count;
}
