namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a font definition with a typeface name.
/// </summary>
public sealed class FontData : IFontData
{
    /// <summary>
    /// Initializes a new instance of <see cref="FontData"/> with the specified typeface name.
    /// </summary>
    /// <param name="fontName">The font typeface name.</param>
    public FontData(string fontName)
    {
        FontName = fontName;
    }

    /// <inheritdoc/>
    public string FontName { get; }

    /// <summary>
    /// Returns the font name, optionally resolving against a theme.
    /// </summary>
    /// <param name="theme">The theme to resolve against (currently unused).</param>
    /// <returns>The font name.</returns>
    public string GetFontName(object? theme) => FontName;
}
