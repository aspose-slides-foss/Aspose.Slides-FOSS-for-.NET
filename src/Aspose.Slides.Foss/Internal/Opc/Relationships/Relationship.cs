namespace Aspose.Slides.Foss.Internal.Opc.Relationships;

/// <summary>
/// Represents a single relationship in a .rels file.
/// </summary>
/// <param name="Id">The relationship identifier (e.g., "rId1").</param>
/// <param name="Type">The relationship type URI.</param>
/// <param name="Target">The target part path (relative to the source part).</param>
/// <param name="TargetMode">Optional target mode ("External" for external targets).</param>
public sealed record Relationship(
    string Id,
    string Type,
    string Target,
    string? TargetMode = null);
