namespace Aspose.Slides.Foss.Export;

/// <summary>
/// Represents options that control how a presentation is saved.
/// </summary>
public interface ISaveOptions
{
    /// <summary>
    /// Gets or sets the default regular font used when a source font is not found.
    /// </summary>
    string? DefaultRegularFont { get; set; }
}
