using System.Xml.Linq;

namespace Aspose.Slides.Foss.Internal.Opc.ContentTypes;

/// <summary>
/// Manages the [Content_Types].xml file in an OPC package.
/// Provides methods to parse existing content types, add overrides and default extensions,
/// query content types for parts, and serialize back to XML.
/// </summary>
public sealed class ContentTypesManager
{
    private static readonly XNamespace Ns = ContentTypeConstants.CtNamespace;

    /// <summary>
    /// The well-known part name for the content types stream in an OPC package.
    /// </summary>
    public const string PartName = "[Content_Types].xml";

    private readonly OpcPackage _package;
    private XElement _root = null!;

    /// <summary>
    /// Initializes the content types manager for the given OPC package.
    /// Loads existing [Content_Types].xml or creates a minimal default structure.
    /// </summary>
    /// <param name="package">The OPC package to manage content types for.</param>
    internal ContentTypesManager(OpcPackage package)
    {
        _package = package;
        Load();
    }

    /// <summary>
    /// Loads and parses the [Content_Types].xml from the package.
    /// If the part does not exist, creates a minimal structure with default
    /// extensions for <c>.rels</c> and <c>.xml</c>.
    /// </summary>
    public void Load()
    {
        var content = _package.GetPart(PartName);
        if (content is not null)
        {
            using var ms = new MemoryStream(content);
            _root = XElement.Load(ms);
        }
        else
        {
            _root = new XElement(Ns + "Types");
            AddDefaultExtension("rels", "application/vnd.openxmlformats-package.relationships+xml");
            AddDefaultExtension("xml", "application/xml");
        }
    }

    /// <summary>
    /// Adds a default content type mapping for a file extension.
    /// </summary>
    /// <param name="extension">The file extension without a leading dot (e.g., <c>"xml"</c>).</param>
    /// <param name="contentType">The MIME content type to associate with the extension.</param>
    public void AddDefaultExtension(string extension, string contentType)
    {
        var element = new XElement(Ns + "Default");
        element.SetAttributeValue("Extension", extension);
        element.SetAttributeValue("ContentType", contentType);
        _root.Add(element);
    }

    /// <summary>
    /// Adds or updates a content type override for a specific part.
    /// If an override for the given part already exists, its content type is updated.
    /// </summary>
    /// <param name="partName">The part path (e.g., <c>/ppt/slides/slide1.xml</c>).</param>
    /// <param name="contentType">The MIME content type.</param>
    public void AddOverride(string partName, string contentType)
    {
        partName = EnsureLeadingSlash(partName);

        foreach (var existing in _root.Elements(Ns + "Override"))
        {
            if (existing.Attribute("PartName")?.Value == partName)
            {
                existing.SetAttributeValue("ContentType", contentType);
                return;
            }
        }

        var element = new XElement(Ns + "Override");
        element.SetAttributeValue("PartName", partName);
        element.SetAttributeValue("ContentType", contentType);
        _root.Add(element);
    }

    /// <summary>
    /// Removes a content type override for a specific part.
    /// </summary>
    /// <param name="partName">The part path.</param>
    /// <returns><c>true</c> if the override was removed; <c>false</c> if it did not exist.</returns>
    public bool RemoveOverride(string partName)
    {
        partName = EnsureLeadingSlash(partName);

        foreach (var existing in _root.Elements(Ns + "Override"))
        {
            if (existing.Attribute("PartName")?.Value == partName)
            {
                existing.Remove();
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Gets the content type for a specific part by checking overrides first,
    /// then falling back to default extension mappings.
    /// </summary>
    /// <param name="partName">The part path.</param>
    /// <returns>The content type string, or <c>null</c> if not found.</returns>
    public string? GetContentType(string partName)
    {
        partName = EnsureLeadingSlash(partName);

        // Check overrides first
        foreach (var existing in _root.Elements(Ns + "Override"))
        {
            if (existing.Attribute("PartName")?.Value == partName)
            {
                return existing.Attribute("ContentType")?.Value;
            }
        }

        // Fall back to defaults based on extension
        var ext = partName.Contains('.')
            ? partName[(partName.LastIndexOf('.') + 1)..]
            : "";

        foreach (var def in _root.Elements(Ns + "Default"))
        {
            if (def.Attribute("Extension")?.Value == ext)
            {
                return def.Attribute("ContentType")?.Value;
            }
        }

        return null;
    }

    /// <summary>
    /// Saves the content types back to the package as a well-formed XML document.
    /// </summary>
    public void Save()
    {
        var doc = new XDocument(
            new XDeclaration("1.0", "UTF-8", "yes"),
            _root);

        using var ms = new MemoryStream();
        doc.Save(ms);
        _package.SetPart(PartName, ms.ToArray());
    }

    private static string EnsureLeadingSlash(string partName)
    {
        return partName.StartsWith('/') ? partName : "/" + partName;
    }
}
