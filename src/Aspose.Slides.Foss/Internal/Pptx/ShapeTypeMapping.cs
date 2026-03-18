namespace Aspose.Slides.Foss.Internal.Pptx;

/// <summary>
/// Mapping between OOXML preset geometry names (ST_ShapeType) and <see cref="ShapeType"/> enum member names.
/// </summary>
internal static class ShapeTypeMapping
{
    /// <summary>
    /// Converts an OOXML prstGeom <c>prst</c> attribute value to a <see cref="ShapeType"/> enum member name.
    /// </summary>
    /// <param name="prst">The OOXML preset geometry string (e.g. <c>"rect"</c>).</param>
    /// <returns>The enum member name (e.g. <c>"Rectangle"</c>), or <c>null</c> if not recognized.</returns>
    internal static string? OoxmlPrstToShapeTypeName(string prst)
    {
        var shapeType = OoxmlPresetMapping.FromPreset(prst);
        return shapeType == ShapeType.NotDefined ? null : shapeType.ToString();
    }

    /// <summary>
    /// Converts a <see cref="ShapeType"/> enum member name to an OOXML prstGeom <c>prst</c> attribute value.
    /// </summary>
    /// <param name="name">The enum member name (e.g. <c>"Rectangle"</c>).</param>
    /// <returns>The OOXML preset string (e.g. <c>"rect"</c>), or <c>null</c> if not recognized.</returns>
    internal static string? ShapeTypeNameToOoxmlPrst(string name)
    {
        if (!Enum.TryParse<ShapeType>(name, out var shapeType))
            return null;

        return OoxmlPresetMapping.ToPreset(shapeType);
    }
}
