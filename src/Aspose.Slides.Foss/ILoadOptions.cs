namespace Aspose.Slides.Foss;

/// <summary>
/// Represents options that can be used to control how a presentation is loaded.
/// </summary>
public interface ILoadOptions
{
    /// <summary>
    /// Gets or sets the default regular font used when a source font is not found.
    /// </summary>
    string? DefaultRegularFont { get; set; }

    /// <summary>
    /// Gets or sets the default Asian font used when a source Asian font is not found.
    /// </summary>
    string? DefaultAsianFont { get; set; }

    /// <summary>
    /// Gets or sets the default symbol font used when a source symbol font is not found.
    /// </summary>
    string? DefaultSymbolFont { get; set; }

    /// <summary>
    /// Gets or sets the password for opening a protected presentation.
    /// </summary>
    string? Password { get; set; }
}
