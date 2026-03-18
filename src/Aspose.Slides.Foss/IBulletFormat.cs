namespace Aspose.Slides.Foss;

/// <summary>
/// Represents paragraph bullet formatting properties.
/// </summary>
public interface IBulletFormat
{
    /// <summary>
    /// Gets or sets the bullet type.
    /// </summary>
    BulletType Type { get; set; }

    /// <summary>
    /// Gets or sets the bullet character (for symbol bullets).
    /// </summary>
    string Char { get; set; }

    /// <summary>
    /// Gets or sets the bullet font.
    /// </summary>
    IFontData? Font { get; set; }

    /// <summary>
    /// Gets or sets the bullet height as a percentage. NaN means inherited.
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// Gets the bullet color format.
    /// </summary>
    IColorFormat Color { get; }

    /// <summary>
    /// Gets or sets the starting number for numbered bullets.
    /// </summary>
    int NumberedBulletStartWith { get; set; }

    /// <summary>
    /// Gets or sets the numbered bullet style.
    /// </summary>
    NumberedBulletStyle NumberedBulletStyle { get; set; }

    /// <summary>
    /// Gets or sets whether the bullet uses a hard-coded color.
    /// </summary>
    NullableBool IsBulletHardColor { get; set; }

    /// <summary>
    /// Gets or sets whether the bullet uses a hard-coded font.
    /// </summary>
    NullableBool IsBulletHardFont { get; set; }

    /// <summary>
    /// Gets the picture used for picture bullets.
    /// </summary>
    ISlidesPicture? Picture { get; }
}
