using System.IO.Compression;
using System.Xml.Linq;

namespace Aspose.Slides.Foss.ConformanceTests.Harness;

/// <summary>
/// A produced <c>.pptx</c> opened as a plain ZIP archive.
/// <para>
/// Nothing here goes through the library that wrote the file. Every accessor reads the bytes in the
/// package, so an assertion made through this type says what a PowerPoint, an Open XML SDK or a
/// python-pptx user would see — not what the writer believes it wrote.
/// </para>
/// </summary>
internal sealed class PptxPackage : IDisposable
{
    private readonly ZipArchive _zip;

    private PptxPackage(string path, ZipArchive zip)
    {
        Path = path;
        _zip = zip;
        PartNames = zip.Entries.Select(e => e.FullName).OrderBy(n => n, StringComparer.Ordinal).ToArray();
    }

    /// <summary>Gets the path the package was opened from.</summary>
    internal string Path { get; }

    /// <summary>Gets every ZIP entry name in the package, ordered.</summary>
    internal IReadOnlyList<string> PartNames { get; }

    /// <summary>
    /// Gets the part names that carry package content, excluding the content-types part and every
    /// relationships part.
    /// </summary>
    internal IEnumerable<string> ContentPartNames =>
        PartNames.Where(n => n != "[Content_Types].xml" && !IsRelationshipsPart(n));

    /// <summary>Gets the part names that are XML markup this harness can parse.</summary>
    internal IEnumerable<string> XmlPartNames =>
        ContentPartNames.Where(n => n.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));

    /// <summary>
    /// Opens a produced package. Fails the test with the file size when the package cannot be read
    /// as a ZIP archive at all — a 0-byte or truncated output is a conformance failure, not an
    /// infrastructure error.
    /// </summary>
    internal static PptxPackage Open(string path)
    {
        Assert.True(File.Exists(path), $"No package was written to '{path}'.");
        var length = new FileInfo(path).Length;
        try
        {
            return new PptxPackage(path, ZipFile.OpenRead(path));
        }
        catch (InvalidDataException ex)
        {
            throw new Xunit.Sdk.XunitException(
                $"'{path}' is {length} bytes and is not a readable ZIP package: {ex.Message}");
        }
    }

    /// <summary>Returns whether the package contains a part with this exact name.</summary>
    internal bool Contains(string partName) => _zip.GetEntry(partName) is not null;

    /// <summary>Reads a part as text, failing the test when it is absent.</summary>
    internal string Text(string partName)
    {
        var entry = _zip.GetEntry(partName);
        Assert.True(entry is not null,
            $"Part '{partName}' is missing. The package contains: {string.Join(", ", PartNames)}");
        using var reader = new StreamReader(entry!.Open());
        return reader.ReadToEnd();
    }

    /// <summary>Reads a part as XML, failing the test when it is absent.</summary>
    internal XDocument Xml(string partName) => XDocument.Parse(Text(partName));

    /// <summary>
    /// Reads the relationships part belonging to <paramref name="partName"/>, or <c>null</c> when the
    /// part has none. Relationships for <c>ppt/slides/slide1.xml</c> live in
    /// <c>ppt/slides/_rels/slide1.xml.rels</c>.
    /// </summary>
    internal XDocument? RelationshipsFor(string partName)
    {
        var relsName = RelationshipsPartNameFor(partName);
        return Contains(relsName) ? Xml(relsName) : null;
    }

    /// <summary>Computes the relationships part name for a part.</summary>
    internal static string RelationshipsPartNameFor(string partName)
    {
        var slash = partName.LastIndexOf('/');
        return slash < 0
            ? $"_rels/{partName}.rels"
            : $"{partName[..slash]}/_rels/{partName[(slash + 1)..]}.rels";
    }

    /// <summary>Returns whether a part name is a relationships part.</summary>
    internal static bool IsRelationshipsPart(string partName) =>
        partName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Resolves a relationship target against the directory of the part that declares it, producing an
    /// absolute part name such as <c>ppt/media/image1.png</c>.
    /// </summary>
    internal static string ResolveTarget(string owningPartName, string target)
    {
        if (target.StartsWith('/'))
            return target[1..];

        var slash = owningPartName.LastIndexOf('/');
        var directory = slash < 0 ? string.Empty : owningPartName[..slash];
        var segments = new List<string>(directory.Length == 0 ? [] : directory.Split('/'));
        foreach (var segment in target.Split('/'))
        {
            if (segment is "." or "")
                continue;
            if (segment == "..")
            {
                if (segments.Count > 0)
                    segments.RemoveAt(segments.Count - 1);
            }
            else
            {
                segments.Add(segment);
            }
        }

        return string.Join('/', segments);
    }

    /// <summary>Enumerates the relationships declared for a part.</summary>
    internal IEnumerable<Relationship> Relationships(string partName)
    {
        var rels = RelationshipsFor(partName);
        if (rels?.Root is null)
            yield break;

        foreach (var element in rels.Root.Elements(Ns.Rel + "Relationship"))
        {
            yield return new Relationship(
                (string?)element.Attribute("Id") ?? string.Empty,
                (string?)element.Attribute("Type") ?? string.Empty,
                (string?)element.Attribute("Target") ?? string.Empty,
                (string?)element.Attribute("TargetMode") ?? "Internal");
        }
    }

    /// <summary>
    /// Resolves the content type declared for a part: an <c>Override</c> if one names it, otherwise the
    /// <c>Default</c> registered for its extension, otherwise <c>null</c>.
    /// </summary>
    internal string? ContentTypeOf(string partName)
    {
        var types = Xml("[Content_Types].xml").Root;
        if (types is null)
            return null;

        foreach (var over in types.Elements(Ns.Ct + "Override"))
        {
            var name = (string?)over.Attribute("PartName");
            if (string.Equals(name?.TrimStart('/'), partName, StringComparison.OrdinalIgnoreCase))
                return (string?)over.Attribute("ContentType");
        }

        var extension = System.IO.Path.GetExtension(partName).TrimStart('.');
        foreach (var def in types.Elements(Ns.Ct + "Default"))
        {
            if (string.Equals((string?)def.Attribute("Extension"), extension, StringComparison.OrdinalIgnoreCase))
                return (string?)def.Attribute("ContentType");
        }

        return null;
    }

    /// <summary>Enumerates every part name an <c>Override</c> claims to describe.</summary>
    internal IEnumerable<string> OverriddenPartNames =>
        Xml("[Content_Types].xml").Root?.Elements(Ns.Ct + "Override")
            .Select(o => ((string?)o.Attribute("PartName") ?? string.Empty).TrimStart('/'))
            .Where(n => n.Length > 0)
        ?? [];

    public void Dispose() => _zip.Dispose();

    /// <summary>One row of a <c>.rels</c> part.</summary>
    internal readonly record struct Relationship(string Id, string Type, string Target, string TargetMode);
}
