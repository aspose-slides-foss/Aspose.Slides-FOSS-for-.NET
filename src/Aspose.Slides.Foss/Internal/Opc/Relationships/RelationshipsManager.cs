using System.Xml.Linq;

namespace Aspose.Slides.Foss.Internal.Opc.Relationships;

/// <summary>
/// Manages relationships (.rels) files in an OPC package.
/// Each part can have an associated .rels file that defines its relationships
/// to other parts. This class provides methods to load, add, remove, and
/// serialize relationships back to XML.
/// </summary>
public sealed class RelationshipsManager
{
    private static readonly XNamespace Ns = RelConstants.RelsNamespace;

    private readonly OpcPackage _package;
    private readonly string _sourcePart;
    private readonly string _relsPartName;
    private XElement _root;
    private readonly Dictionary<string, Relationship> _relationships = [];

    /// <summary>
    /// Initializes the relationships manager for a specific part.
    /// </summary>
    /// <param name="package">The OPC package.</param>
    /// <param name="sourcePart">
    /// The part path whose relationships to manage.
    /// Empty string for root relationships (<c>_rels/.rels</c>).
    /// </param>
    public RelationshipsManager(OpcPackage package, string sourcePart = "")
    {
        _package = package;
        _sourcePart = sourcePart;
        _relsPartName = GetRelsPartName(sourcePart);
        _root = new XElement(Ns + "Relationships");
        Load();
    }

    /// <summary>
    /// Gets the .rels file path for a given source part.
    /// </summary>
    /// <param name="sourcePart">The source part path.</param>
    /// <returns>The corresponding .rels part path.</returns>
    /// <example>
    /// <c>""</c> → <c>"_rels/.rels"</c><br/>
    /// <c>"ppt/presentation.xml"</c> → <c>"ppt/_rels/presentation.xml.rels"</c><br/>
    /// <c>"ppt/slides/slide1.xml"</c> → <c>"ppt/slides/_rels/slide1.xml.rels"</c>
    /// </example>
    public static string GetRelsPartName(string sourcePart)
    {
        if (string.IsNullOrEmpty(sourcePart))
            return "_rels/.rels";

        var slashIndex = sourcePart.LastIndexOf('/');
        if (slashIndex >= 0)
        {
            var directory = sourcePart[..slashIndex];
            var filename = sourcePart[(slashIndex + 1)..];
            return $"{directory}/_rels/{filename}.rels";
        }

        return $"_rels/{sourcePart}.rels";
    }

    /// <summary>
    /// Loads and parses the .rels file from the package.
    /// </summary>
    public void Load()
    {
        var content = _package.GetPart(_relsPartName);
        if (content is not null)
        {
            using var ms = new MemoryStream(content);
            _root = XElement.Load(ms);
            ParseRelationships();
        }
        else
        {
            _root = new XElement(Ns + "Relationships");
        }
    }

    /// <summary>
    /// Parses relationships from the loaded XML element.
    /// </summary>
    public void ParseRelationships()
    {
        _relationships.Clear();
        foreach (var relElem in _root.Elements(Ns + "Relationship"))
        {
            var rel = new Relationship(
                Id: (string?)relElem.Attribute("Id") ?? "",
                Type: (string?)relElem.Attribute("Type") ?? "",
                Target: (string?)relElem.Attribute("Target") ?? "",
                TargetMode: (string?)relElem.Attribute("TargetMode")
            );
            _relationships[rel.Id] = rel;
        }
    }

    /// <summary>
    /// Gets a relationship by ID.
    /// </summary>
    /// <param name="relId">The relationship ID (e.g., "rId1").</param>
    /// <returns>The <see cref="Relationship"/> or <c>null</c> if not found.</returns>
    public Relationship? GetRelationship(string relId)
    {
        return _relationships.GetValueOrDefault(relId);
    }

    /// <summary>
    /// Gets all relationships of a specific type.
    /// </summary>
    /// <param name="relType">The relationship type URI.</param>
    /// <returns>A list of matching relationships.</returns>
    public List<Relationship> GetRelationshipsByType(string relType)
    {
        return [.. _relationships.Values.Where(r => r.Type == relType)];
    }

    /// <summary>
    /// Gets all relationships.
    /// </summary>
    /// <returns>A list of all relationships.</returns>
    public List<Relationship> GetAllRelationships()
    {
        return [.. _relationships.Values];
    }

    /// <summary>
    /// Generates a unique relationship ID of the form "rIdN".
    /// </summary>
    /// <returns>A relationship ID not already in use.</returns>
    public string GenerateRelId()
    {
        var existing = _relationships.Keys.ToHashSet();
        var counter = 1;
        while (true)
        {
            var relId = $"rId{counter}";
            if (!existing.Contains(relId))
                return relId;
            counter++;
        }
    }

    /// <summary>
    /// Adds a new relationship.
    /// </summary>
    /// <param name="relType">The relationship type URI.</param>
    /// <param name="target">The target part path (relative to the source).</param>
    /// <param name="relId">Optional specific ID. If <c>null</c>, auto-generated.</param>
    /// <param name="targetMode">Optional target mode ("External" for external links).</param>
    /// <returns>The relationship ID.</returns>
    public string AddRelationship(string relType, string target, string? relId = null, string? targetMode = null)
    {
        relId ??= GenerateRelId();

        var rel = new Relationship(
            Id: relId,
            Type: relType,
            Target: target,
            TargetMode: targetMode
        );
        _relationships[relId] = rel;

        var relElem = new XElement(Ns + "Relationship",
            new XAttribute("Id", relId),
            new XAttribute("Type", relType),
            new XAttribute("Target", target));

        if (targetMode is not null)
            relElem.Add(new XAttribute("TargetMode", targetMode));

        _root.Add(relElem);
        return relId;
    }

    /// <summary>
    /// Removes a relationship by ID.
    /// </summary>
    /// <param name="relId">The relationship ID to remove.</param>
    /// <returns><c>true</c> if removed; <c>false</c> if not found.</returns>
    public bool RemoveRelationship(string relId)
    {
        if (!_relationships.Remove(relId))
            return false;

        var elem = _root.Elements(Ns + "Relationship")
            .FirstOrDefault(e => (string?)e.Attribute("Id") == relId);
        elem?.Remove();
        return true;
    }

    /// <summary>
    /// Saves the relationships back to the package.
    /// If there are no relationships, the .rels part is deleted.
    /// </summary>
    public void Save()
    {
        if (_relationships.Count == 0)
        {
            _package.DeletePart(_relsPartName);
            return;
        }

        var doc = new XDocument(
            new XDeclaration("1.0", "UTF-8", "yes"),
            _root);

        using var ms = new MemoryStream();
        doc.Save(ms);
        _package.SetPart(_relsPartName, ms.ToArray());
    }

    /// <summary>
    /// Gets the .rels file part name for this manager.
    /// </summary>
    public string PartName => _relsPartName;
}
