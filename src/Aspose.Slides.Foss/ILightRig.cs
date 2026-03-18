namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a light rig.
/// </summary>
public interface ILightRig
{
    /// <summary>
    /// Light direction. Read/write.
    /// </summary>
    LightingDirection Direction { get; set; }

    /// <summary>
    /// Represents a preset light right that can be applied to a shape. Read/write.
    /// </summary>
    LightRigPresetType LightType { get; set; }

    /// <summary>
    /// Sets the rotation of the light rig.
    /// </summary>
    void SetRotation(float latitude, float longitude, float revolution);

    /// <summary>
    /// Gets the rotation of the light rig as [latitude, longitude, revolution].
    /// </summary>
    float[] GetRotation();
}
