namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a picture fill style.
/// </summary>
public interface IPictureFillFormat
{
    /// <summary>
    /// Returns or sets the dpi which is used to fill a picture. Read/write.
    /// </summary>
    int Dpi { get; set; }

    /// <summary>
    /// Returns or sets the picture fill mode. Read/write.
    /// </summary>
    PictureFillMode PictureFillMode { get; set; }

    /// <summary>
    /// Returns the picture. Read-only.
    /// </summary>
    ISlidesPicture Picture { get; }

    /// <summary>
    /// Returns or sets the number of percents of real image width that are cropped off
    /// the left of the picture. Read/write.
    /// </summary>
    float CropLeft { get; set; }

    /// <summary>
    /// Returns or sets the number of percents of real image height that are cropped off
    /// the top of the picture. Read/write.
    /// </summary>
    float CropTop { get; set; }

    /// <summary>
    /// Returns or sets the number of percents of real image width that are cropped off
    /// the right of the picture. Read/write.
    /// </summary>
    float CropRight { get; set; }

    /// <summary>
    /// Returns or sets the number of percents of real image height that are cropped off
    /// the bottom of the picture. Read/write.
    /// </summary>
    float CropBottom { get; set; }

    /// <summary>
    /// Returns or sets left edge of the fill rectangle defined by a percentage offset
    /// from the left edge of the shape's bounding box. Read/write.
    /// </summary>
    float StretchOffsetLeft { get; set; }

    /// <summary>
    /// Returns or sets top edge of the fill rectangle defined by a percentage offset
    /// from the top edge of the shape's bounding box. Read/write.
    /// </summary>
    float StretchOffsetTop { get; set; }

    /// <summary>
    /// Returns or sets right edge of the fill rectangle defined by a percentage offset
    /// from the right edge of the shape's bounding box. Read/write.
    /// </summary>
    float StretchOffsetRight { get; set; }

    /// <summary>
    /// Returns or sets bottom edge of the fill rectangle defined by a percentage offset
    /// from the bottom edge of the shape's bounding box. Read/write.
    /// </summary>
    float StretchOffsetBottom { get; set; }

    /// <summary>
    /// Returns or sets the horizontal offset of the texture from the shape's origin in points.
    /// Read/write.
    /// </summary>
    float TileOffsetX { get; set; }

    /// <summary>
    /// Returns or sets the vertical offset of the texture from the shape's origin in points.
    /// Read/write.
    /// </summary>
    float TileOffsetY { get; set; }

    /// <summary>
    /// Returns or sets the horizontal scale for the texture fill as a percentage. Read/write.
    /// </summary>
    float TileScaleX { get; set; }

    /// <summary>
    /// Returns or sets the vertical scale for the texture fill as a percentage. Read/write.
    /// </summary>
    float TileScaleY { get; set; }

    /// <summary>
    /// Returns or sets how the texture is aligned within the shape. Read/write.
    /// </summary>
    RectangleAlignment TileAlignment { get; set; }

    /// <summary>
    /// Flips the texture tile around its horizontal, vertical or both axis. Read/write.
    /// </summary>
    TileFlip TileFlip { get; set; }
}
