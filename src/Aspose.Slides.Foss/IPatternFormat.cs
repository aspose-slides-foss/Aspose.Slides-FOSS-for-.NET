namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a pattern fill format.
/// </summary>
public interface IPatternFormat
{
    /// <summary>
    /// Gets or sets the pattern style.
    /// </summary>
    PatternStyle PatternStyle { get; set; }

    /// <summary>
    /// Gets the foreground pattern color.
    /// </summary>
    IColorFormat ForeColor { get; }

    /// <summary>
    /// Gets the background pattern color.
    /// </summary>
    IColorFormat BackColor { get; }
}
