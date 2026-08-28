namespace Aspose.Slides.Foss.Internal;

/// <summary>
/// Part-name arithmetic for OPC packages.
/// </summary>
/// <remarks>
/// Part names are absolute and slash-separated ("ppt/slides/slide1.xml"); relationship targets are
/// normally relative to the directory of the part that owns the relationship ("../media/image1.png").
/// These helpers convert between the two, which is the step that decides whether an <c>r:embed</c>
/// resolves to the part the caller meant.
/// </remarks>
internal static class OpcPaths
{
    /// <summary>
    /// Gets the directory of a part name, without a trailing slash.
    /// Returns an empty string for a part at the package root.
    /// </summary>
    internal static string DirectoryOf(string partName)
    {
        var lastSlash = partName.LastIndexOf('/');
        return lastSlash >= 0 ? partName[..lastSlash] : string.Empty;
    }

    /// <summary>
    /// Gets the <c>_rels</c> part name holding the relationships owned by <paramref name="partName"/>.
    /// </summary>
    internal static string RelsPartNameFor(string partName)
    {
        var dir = DirectoryOf(partName);
        var file = partName[(partName.LastIndexOf('/') + 1)..];
        return dir.Length == 0 ? $"_rels/{file}.rels" : $"{dir}/_rels/{file}.rels";
    }

    /// <summary>
    /// Expresses an absolute part name as a relationship target relative to <paramref name="fromDir"/>.
    /// </summary>
    internal static string Relative(string fromDir, string toPartName)
    {
        var fromParts = fromDir.Split('/', StringSplitOptions.RemoveEmptyEntries);
        var toParts = toPartName.Split('/', StringSplitOptions.RemoveEmptyEntries);

        int common = 0;
        int max = Math.Min(fromParts.Length, toParts.Length);
        while (common < max &&
               string.Equals(fromParts[common], toParts[common], StringComparison.OrdinalIgnoreCase))
        {
            common++;
        }

        var segments = new List<string>();
        for (int i = common; i < fromParts.Length; i++)
            segments.Add("..");
        for (int i = common; i < toParts.Length; i++)
            segments.Add(toParts[i]);

        return string.Join("/", segments);
    }

    /// <summary>
    /// Resolves a relationship target against the directory of the part that owns it,
    /// producing an absolute part name.
    /// </summary>
    internal static string Resolve(string baseDir, string target)
    {
        var parts = baseDir.Split('/', StringSplitOptions.RemoveEmptyEntries).ToList();

        if (target.StartsWith('/'))
            parts.Clear();

        foreach (var segment in target.Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            if (segment == "..")
            {
                if (parts.Count > 0)
                    parts.RemoveAt(parts.Count - 1);
            }
            else if (segment != ".")
            {
                parts.Add(segment);
            }
        }

        return string.Join("/", parts);
    }
}
