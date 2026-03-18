namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a picture frame shape.
/// </summary>
public interface IPictureFrame : IGeometryShape
{
    /// <summary>
    /// Gets the picture frame lock settings.
    /// </summary>
    IPictureFrameLock? PictureFrameLock { get; }

    /// <summary>
    /// Gets the picture fill format.
    /// </summary>
    IPictureFillFormat? PictureFormat { get; }

    /// <summary>
    /// Gets or sets the relative scale height. <c>1.0</c> means 100%.
    /// </summary>
    float RelativeScaleHeight { get; set; }

    /// <summary>
    /// Gets or sets the relative scale width. <c>1.0</c> means 100%.
    /// </summary>
    float RelativeScaleWidth { get; set; }
}
