namespace Aspose.Slides.Foss;

/// <summary>
/// Represents 3-D properties.
/// </summary>
public interface IThreeDFormat
{
    /// <summary>
    /// Returns or sets the width of a 3D contour. Read/write.
    /// </summary>
    float ContourWidth { get; set; }

    /// <summary>
    /// Returns or sets the height of an extrusion effect. Read/write.
    /// </summary>
    float ExtrusionHeight { get; set; }

    /// <summary>
    /// Returns or sets the depth of a 3D shape. Read/write.
    /// </summary>
    float Depth { get; set; }

    /// <summary>
    /// Returns the type of a top 3D bevel. Read-only.
    /// </summary>
    IShapeBevel BevelTop { get; }

    /// <summary>
    /// Returns the type of a bottom 3D bevel. Read-only.
    /// </summary>
    IShapeBevel BevelBottom { get; }

    /// <summary>
    /// Returns the color of a contour. Read-only.
    /// </summary>
    IColorFormat ContourColor { get; }

    /// <summary>
    /// Returns the color of an extrusion. Read-only.
    /// </summary>
    IColorFormat ExtrusionColor { get; }

    /// <summary>
    /// Returns the settings of a camera. Read-only.
    /// </summary>
    ICamera Camera { get; }

    /// <summary>
    /// Returns the type of a light. Read-only.
    /// </summary>
    ILightRig LightRig { get; }

    /// <summary>
    /// Returns or sets the type of a material. Read/write.
    /// </summary>
    MaterialPresetType Material { get; set; }
}
