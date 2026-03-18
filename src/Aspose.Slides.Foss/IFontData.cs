namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a font definition.
/// </summary>
public interface IFontData
{
    /// <summary>
    /// Gets the font name (typeface).
    /// </summary>
    string FontName { get; }

    /// <summary>
    /// Returns the font name, optionally resolving against a theme.
    /// </summary>
    /// <param name="theme">The theme to resolve against.</param>
    /// <returns>The resolved font name.</returns>
    string GetFontName(object? theme);
}
