namespace Aspose.Slides.Foss.Internal;

/// <summary>
/// Manages default content type registrations for file extensions in [Content_Types].xml.
/// </summary>
internal sealed class ContentTypesManager
{
    private readonly Dictionary<string, string> _defaults = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Registers a default content type for the given file extension if not already present.
    /// </summary>
    internal void RegisterDefaultContentType(string extension, string contentType)
    {
        _defaults.TryAdd(extension, contentType);
    }

    /// <summary>
    /// Gets the registered content type for the given extension, or <c>null</c> if not registered.
    /// </summary>
    internal string? GetContentType(string extension)
    {
        return _defaults.TryGetValue(extension, out var ct) ? ct : null;
    }
}
