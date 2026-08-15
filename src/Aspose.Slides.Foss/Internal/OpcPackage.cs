using System.IO.Compression;

namespace Aspose.Slides.Foss.Internal;

/// <summary>
/// Represents an Open Packaging Convention (OPC) package containing parts.
/// </summary>
internal sealed class OpcPackage
{
    private readonly Dictionary<string, byte[]> _parts = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Creates a new empty OPC package.
    /// </summary>
    internal static OpcPackage CreateNew()
    {
        return new OpcPackage();
    }

    /// <summary>
    /// Opens an OPC package from a file path by reading it as a ZIP archive.
    /// </summary>
    internal static OpcPackage Open(string path)
    {
        using var stream = File.OpenRead(path);
        return Open(stream);
    }

    /// <summary>
    /// Opens an OPC package from a stream by reading it as a ZIP archive.
    /// </summary>
    internal static OpcPackage Open(Stream stream)
    {
        var package = new OpcPackage();
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        foreach (var entry in archive.Entries)
        {
            if (string.IsNullOrEmpty(entry.Name))
                continue;

            using var entryStream = entry.Open();
            using var ms = new MemoryStream();
            entryStream.CopyTo(ms);
            package._parts[entry.FullName] = ms.ToArray();
        }
        return package;
    }

    /// <summary>
    /// Creates an independent copy of this package.
    /// </summary>
    /// <remarks>
    /// Part contents are immutable byte arrays that are replaced rather than edited in place, so the
    /// copy shares them safely. This exists so a save that writes something other than the whole
    /// presentation — a subset of slides — can build it without changing the presentation the caller
    /// still holds.
    /// </remarks>
    internal OpcPackage Clone()
    {
        var copy = new OpcPackage();
        foreach (var (partName, data) in _parts)
            copy._parts[partName] = data;
        return copy;
    }

    /// <summary>
    /// Saves the OPC package to a stream as a ZIP archive.
    /// </summary>
    internal void SaveToStream(Stream stream)
    {
        using var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);
        foreach (var (partName, data) in _parts)
        {
            var entry = archive.CreateEntry(partName, CompressionLevel.Optimal);
            using var entryStream = entry.Open();
            entryStream.Write(data, 0, data.Length);
        }
    }

    /// <summary>
    /// Gets the binary data for the specified part name.
    /// </summary>
    internal byte[]? GetPart(string partName)
    {
        return _parts.TryGetValue(partName, out var data) ? data : null;
    }

    /// <summary>
    /// Stores binary data for the specified part name.
    /// </summary>
    internal void SetPart(string partName, byte[] data)
    {
        _parts[partName] = data;
    }

    /// <summary>
    /// Removes a part from the package by name.
    /// </summary>
    internal void RemovePart(string partName)
    {
        _parts.Remove(partName);
    }

    /// <summary>
    /// Gets all part names in the package.
    /// </summary>
    internal IEnumerable<string> GetPartNames()
    {
        return _parts.Keys;
    }

    /// <summary>
    /// Gets all part names sorted alphabetically.
    /// </summary>
    internal IReadOnlyList<string> GetSortedPartNames()
    {
        var names = _parts.Keys.ToList();
        names.Sort(StringComparer.OrdinalIgnoreCase);
        return names;
    }
}
