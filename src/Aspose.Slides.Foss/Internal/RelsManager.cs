using System.Xml.Linq;

namespace Aspose.Slides.Foss.Internal;

/// <summary>
/// Manages OPC relationships for a part (e.g., slide → image, slide → external link).
/// </summary>
internal sealed class RelsManager
{
    private static readonly XNamespace RelsNs = "http://schemas.openxmlformats.org/package/2006/relationships";

    private const string ImageRelType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

    private readonly List<Relationship> _relationships = new();
    private int _nextId = 1;

    /// <summary>
    /// Gets or sets the package these relationships are written to.
    /// </summary>
    internal OpcPackage? Package { get; set; }

    /// <summary>
    /// Gets or sets the name of the part that owns these relationships
    /// (e.g., "ppt/slides/slide1.xml"). Determines where <see cref="Save"/> writes.
    /// </summary>
    internal string OwnerPartName { get; set; } = string.Empty;

    /// <summary>
    /// Gets a relationship by its ID.
    /// </summary>
    internal Relationship? GetById(string id)
    {
        return _relationships.Find(r => r.Id == id);
    }

    /// <summary>
    /// Gets every relationship this part declares.
    /// </summary>
    internal IReadOnlyList<Relationship> All => _relationships;

    /// <summary>
    /// Finds relationships matching the given type.
    /// </summary>
    internal IEnumerable<Relationship> FindByType(string relationshipType)
    {
        return _relationships.Where(r => r.Type == relationshipType);
    }

    /// <summary>
    /// Adds a new relationship and returns its ID.
    /// </summary>
    internal string Add(string type, string target, bool isExternal = false)
    {
        var id = $"rId{_nextId++}";
        _relationships.Add(new Relationship(id, type, target, isExternal));
        return id;
    }

    /// <summary>
    /// Removes a relationship by ID.
    /// </summary>
    internal void Remove(string id)
    {
        _relationships.RemoveAll(r => r.Id == id);
    }

    /// <summary>
    /// Loads relationships from a .rels XML byte array.
    /// </summary>
    internal void Load(byte[] data)
    {
        _relationships.Clear();

        using var ms = new MemoryStream(data);
        var doc = XDocument.Load(ms);
        var root = doc.Root;
        if (root is null)
            return;

        int maxId = 0;
        foreach (var el in root.Elements(RelsNs + "Relationship"))
        {
            var id = el.Attribute("Id")?.Value ?? "";
            var type = el.Attribute("Type")?.Value ?? "";
            var target = el.Attribute("Target")?.Value ?? "";
            var targetMode = el.Attribute("TargetMode")?.Value;
            var isExternal = string.Equals(targetMode, "External", StringComparison.OrdinalIgnoreCase);

            _relationships.Add(new Relationship(id, type, target, isExternal));

            if (id.StartsWith("rId", StringComparison.OrdinalIgnoreCase) &&
                int.TryParse(id.AsSpan(3), out var num) && num > maxId)
            {
                maxId = num;
            }
        }

        _nextId = maxId + 1;
    }

    /// <summary>
    /// Serializes the relationships to a .rels XML byte array.
    /// </summary>
    internal byte[] ToBytes()
    {
        var doc = new XDocument(
            new XDeclaration("1.0", "UTF-8", "yes"),
            new XElement(RelsNs + "Relationships",
                _relationships.Select(r =>
                {
                    var el = new XElement(RelsNs + "Relationship",
                        new XAttribute("Id", r.Id),
                        new XAttribute("Type", r.Type),
                        new XAttribute("Target", r.Target));
                    if (r.IsExternal)
                        el.Add(new XAttribute("TargetMode", "External"));
                    return el;
                })));

        using var ms = new MemoryStream();
        doc.Save(ms);
        return ms.ToArray();
    }

    /// <summary>
    /// Returns the id of a relationship from the owning part to the given media part, adding one
    /// and persisting it if it does not already exist.
    /// </summary>
    /// <param name="imagePartName">The absolute part name of the image (e.g., "ppt/media/image1.png").</param>
    /// <returns>The relationship id to write as <c>r:embed</c>.</returns>
    internal string EnsureImageRelationship(string imagePartName)
    {
        var fromDir = OpcPaths.DirectoryOf(OwnerPartName);

        foreach (var rel in FindByType(ImageRelType))
        {
            if (!rel.IsExternal &&
                string.Equals(OpcPaths.Resolve(fromDir, rel.Target), imagePartName,
                    StringComparison.OrdinalIgnoreCase))
            {
                return rel.Id;
            }
        }

        var id = Add(ImageRelType, OpcPaths.Relative(fromDir, imagePartName));
        Save();
        return id;
    }

    /// <summary>
    /// Writes the relationships to the owning part's <c>.rels</c> part in the package.
    /// </summary>
    /// <remarks>
    /// Does nothing when <see cref="Package"/> or <see cref="OwnerPartName"/> has not been set,
    /// which is the case for the short-lived instances used to read an existing <c>.rels</c> part.
    /// Anything that adds a relationship meant to reach the file must be wired to both.
    /// </remarks>
    internal void Save()
    {
        if (Package is null || string.IsNullOrEmpty(OwnerPartName))
            return;

        Package.SetPart(OpcPaths.RelsPartNameFor(OwnerPartName), ToBytes());
    }
}

/// <summary>
/// Represents a single OPC relationship.
/// </summary>
internal sealed class Relationship
{
    internal string Id { get; }
    internal string Type { get; }
    internal string Target { get; }
    internal bool IsExternal { get; }

    internal Relationship(string id, string type, string target, bool isExternal)
    {
        Id = id;
        Type = type;
        Target = target;
        IsExternal = isExternal;
    }
}
