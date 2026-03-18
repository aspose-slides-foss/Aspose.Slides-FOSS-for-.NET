namespace Aspose.Slides.Foss;

/// <summary>
/// Represents options that can be used to control how a presentation is loaded.
/// </summary>
public sealed class LoadOptions : ILoadOptions
{
    /// <inheritdoc />
    public string? DefaultRegularFont { get; set; }

    /// <inheritdoc />
    public string? DefaultAsianFont { get; set; }

    /// <inheritdoc />
    public string? DefaultSymbolFont { get; set; }

    /// <inheritdoc />
    public string? Password { get; set; }
}
