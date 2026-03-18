namespace Aspose.Slides.Foss;

/// <summary>
/// Represents the 3-D camera properties for a shape.
/// </summary>
public interface ICamera
{
    /// <summary>
    /// Gets or sets the camera preset type.
    /// </summary>
    CameraPresetType CameraType { get; set; }

    /// <summary>
    /// Gets or sets the field of view angle in degrees.
    /// </summary>
    float FieldOfViewAngle { get; set; }

    /// <summary>
    /// Gets or sets the zoom percentage.
    /// </summary>
    float Zoom { get; set; }

    /// <summary>
    /// Sets the camera rotation angles.
    /// </summary>
    void SetRotation(float latitude, float longitude, float revolution);

    /// <summary>
    /// Gets the camera rotation as [latitude, longitude, revolution] in degrees.
    /// </summary>
    float[] GetRotation();
}
