using System.IO.Compression;

namespace Aspose.Slides.Foss.Internal.Opc;

/// <summary>
/// Manages an Open Packaging Conventions (OPC) package.
/// An OPC package is a ZIP archive containing parts (files) organized
/// according to Office Open XML conventions. Supports loading/saving
/// from file paths or streams and preserves unknown parts for round-trip fidelity.
/// </summary>
public sealed class OpcPackage
{
    private readonly Dictionary<string, byte[]> _parts = [];
    private string? _sourcePath;

    /// <summary>
    /// Gets the original file path if the package was loaded from a file,
    /// or <c>null</c> if loaded from a stream or created new.
    /// </summary>
    public string? SourcePath => _sourcePath;

    /// <summary>
    /// Opens an OPC package from a file path.
    /// </summary>
    /// <param name="path">Path to the package file.</param>
    /// <returns>A loaded <see cref="OpcPackage"/> instance.</returns>
    /// <exception cref="FileNotFoundException">If the file does not exist.</exception>
    public static OpcPackage Open(string path)
    {
        if (!File.Exists(path))
        {
            throw new FileNotFoundException($"Package file not found: {path}", path);
        }

        var package = new OpcPackage { _sourcePath = path };
        package.LoadFromPath(path);
        return package;
    }

    /// <summary>
    /// Opens an OPC package from a binary stream.
    /// </summary>
    /// <param name="stream">A readable stream containing a ZIP archive.</param>
    /// <returns>A loaded <see cref="OpcPackage"/> instance.</returns>
    public static OpcPackage Open(Stream stream)
    {
        var package = new OpcPackage();
        package.LoadFromStream(stream);
        return package;
    }

    /// <summary>
    /// Creates a new empty OPC package.
    /// </summary>
    /// <returns>A new empty <see cref="OpcPackage"/> instance.</returns>
    public static OpcPackage CreateNew() => new();

    /// <summary>
    /// Loads all parts from a ZIP file at the given path.
    /// </summary>
    /// <param name="path">Path to the ZIP file.</param>
    internal void LoadFromPath(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        LoadFromZipfile(archive);
    }

    /// <summary>
    /// Loads all parts from a ZIP stream.
    /// </summary>
    /// <param name="stream">A readable stream containing a ZIP archive.</param>
    internal void LoadFromStream(Stream stream)
    {
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        LoadFromZipfile(archive);
    }

    /// <summary>
    /// Loads all parts from an open <see cref="ZipArchive"/>.
    /// </summary>
    /// <param name="archive">The ZIP archive to read parts from.</param>
    internal void LoadFromZipfile(ZipArchive archive)
    {
        foreach (var entry in archive.Entries)
        {
            using var entryStream = entry.Open();
            using var ms = new MemoryStream();
            entryStream.CopyTo(ms);
            _parts[entry.FullName] = ms.ToArray();
        }
    }

    /// <summary>
    /// Saves the OPC package to a file path.
    /// </summary>
    /// <param name="path">Destination file path.</param>
    public void Save(string path) => SaveToPath(path);

    /// <summary>
    /// Saves the OPC package to a stream.
    /// </summary>
    /// <param name="stream">Destination writable stream.</param>
    public void Save(Stream stream) => SaveToStream(stream);

    /// <summary>
    /// Saves all parts to a ZIP file at the given path.
    /// </summary>
    /// <param name="path">Destination file path.</param>
    internal void SaveToPath(string path)
    {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        SaveToZipfile(archive);
    }

    /// <summary>
    /// Saves all parts to a ZIP stream.
    /// </summary>
    /// <param name="stream">Destination writable stream.</param>
    internal void SaveToStream(Stream stream)
    {
        using var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);
        SaveToZipfile(archive);
    }

    /// <summary>
    /// Saves all parts to an open <see cref="ZipArchive"/>.
    /// </summary>
    /// <param name="archive">The ZIP archive to write parts to.</param>
    internal void SaveToZipfile(ZipArchive archive)
    {
        foreach (var (name, content) in _parts)
        {
            var entry = archive.CreateEntry(name, CompressionLevel.Optimal);
            using var entryStream = entry.Open();
            entryStream.Write(content);
        }
    }

    /// <summary>
    /// Gets the content of a part by name.
    /// </summary>
    /// <param name="partName">The part path within the package (e.g., "ppt/presentation.xml").</param>
    /// <returns>Part content as a byte array, or <c>null</c> if the part does not exist.</returns>
    public byte[]? GetPart(string partName)
    {
        return _parts.GetValueOrDefault(partName);
    }

    /// <summary>
    /// Sets or updates the content of a part.
    /// </summary>
    /// <param name="partName">The part path within the package.</param>
    /// <param name="content">Part content as bytes.</param>
    public void SetPart(string partName, byte[] content)
    {
        _parts[partName] = content;
    }

    /// <summary>
    /// Sets or updates the content of a part from a string (encoded as UTF-8).
    /// </summary>
    /// <param name="partName">The part path within the package.</param>
    /// <param name="content">Part content as a string.</param>
    public void SetPart(string partName, string content)
    {
        _parts[partName] = System.Text.Encoding.UTF8.GetBytes(content);
    }

    /// <summary>
    /// Checks if a part exists in the package.
    /// </summary>
    /// <param name="partName">The part path to check.</param>
    /// <returns><c>true</c> if the part exists; otherwise, <c>false</c>.</returns>
    public bool HasPart(string partName)
    {
        return _parts.ContainsKey(partName);
    }

    /// <summary>
    /// Deletes a part from the package.
    /// </summary>
    /// <param name="partName">The part path to delete.</param>
    /// <returns><c>true</c> if the part was deleted; <c>false</c> if it did not exist.</returns>
    public bool DeletePart(string partName)
    {
        return _parts.Remove(partName);
    }

    /// <summary>
    /// Gets a list of all part names in the package.
    /// </summary>
    /// <returns>A list of part paths.</returns>
    public List<string> GetPartNames()
    {
        return [.. _parts.Keys];
    }

    /// <summary>
    /// Closes the package and releases resources.
    /// Clears all in-memory part data. The package should not be used after closing.
    /// </summary>
    public void Close()
    {
        _parts.Clear();
        _sourcePath = null;
    }
}
