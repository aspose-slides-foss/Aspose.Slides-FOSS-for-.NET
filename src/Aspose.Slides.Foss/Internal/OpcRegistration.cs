using System.Xml.Linq;

namespace Aspose.Slides.Foss.Internal;

/// <summary>
/// Helpers to register content types and relationships in an OPC package.
/// </summary>
internal static class OpcRegistration
{
    private static readonly XNamespace CtNs =
        "http://schemas.openxmlformats.org/package/2006/content-types";

    /// <summary>
    /// Adds a content type override for the given part name if one does not already exist.
    /// </summary>
    internal static void AddContentTypeOverride(OpcPackage package, string partName, string contentType)
    {
        var ctData = package.GetPart("[Content_Types].xml");
        XDocument doc;
        if (ctData is not null)
        {
            using var ms = new MemoryStream(ctData);
            doc = XDocument.Load(ms);
        }
        else
        {
            doc = new XDocument(new XElement(CtNs + "Types"));
        }

        var root = doc.Root!;
        var existing = root.Elements(CtNs + "Override")
            .FirstOrDefault(e => string.Equals(
                e.Attribute("PartName")?.Value?.TrimStart('/'),
                partName.TrimStart('/'),
                StringComparison.OrdinalIgnoreCase));

        if (existing is null)
        {
            root.Add(new XElement(CtNs + "Override",
                new XAttribute("PartName", "/" + partName.TrimStart('/')),
                new XAttribute("ContentType", contentType)));
        }
        else if (existing.Attribute("ContentType")?.Value != contentType)
        {
            existing.SetAttributeValue("ContentType", contentType);
        }

        using var outMs = new MemoryStream();
        doc.Save(outMs);
        package.SetPart("[Content_Types].xml", outMs.ToArray());
    }

    /// <summary>
    /// Removes the content type override for the given part name, if one is declared.
    /// </summary>
    /// <remarks>
    /// An <c>Override</c> naming a part that is no longer in the package is invalid by
    /// ISO/IEC 29500-2 §10.1.2.3, so deleting a part has to delete its declaration with it.
    /// </remarks>
    internal static void RemoveContentTypeOverride(OpcPackage package, string partName)
    {
        var ctData = package.GetPart("[Content_Types].xml");
        if (ctData is null)
            return;

        XDocument doc;
        using (var ms = new MemoryStream(ctData))
            doc = XDocument.Load(ms);

        var root = doc.Root;
        if (root is null)
            return;

        var existing = root.Elements(CtNs + "Override")
            .Where(e => string.Equals(
                e.Attribute("PartName")?.Value?.TrimStart('/'),
                partName.TrimStart('/'),
                StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (existing.Count == 0)
            return;

        foreach (var element in existing)
            element.Remove();

        using var outMs = new MemoryStream();
        doc.Save(outMs);
        package.SetPart("[Content_Types].xml", outMs.ToArray());
    }

    /// <summary>
    /// Declares a default content type for a file extension, adding the <c>Default</c> element if
    /// it is absent and correcting it if it names a different content type.
    /// </summary>
    /// <remarks>
    /// A part whose extension has no <c>Default</c> and no <c>Override</c> resolves no content type
    /// at all. Per ISO/IEC 29500-2 §10.1.2 the content type is the part's identity, so such a part
    /// is not the image (or anything else) it was meant to be, and consumers reject the package.
    /// §10.1.2.2 allows at most one <c>Default</c> per extension, hence the update-in-place.
    /// </remarks>
    internal static void AddContentTypeDefault(OpcPackage package, string extension, string contentType)
    {
        extension = extension.TrimStart('.');

        var ctData = package.GetPart("[Content_Types].xml");
        XDocument doc;
        if (ctData is not null)
        {
            using var ms = new MemoryStream(ctData);
            doc = XDocument.Load(ms);
        }
        else
        {
            doc = new XDocument(new XElement(CtNs + "Types"));
        }

        var root = doc.Root!;
        var existing = root.Elements(CtNs + "Default")
            .FirstOrDefault(e => string.Equals(
                e.Attribute("Extension")?.Value,
                extension,
                StringComparison.OrdinalIgnoreCase));

        if (existing is null)
        {
            // Default elements precede Override elements in the documents PowerPoint writes; keep
            // that shape rather than appending after the overrides.
            var newDefault = new XElement(CtNs + "Default",
                new XAttribute("Extension", extension),
                new XAttribute("ContentType", contentType));

            var lastDefault = root.Elements(CtNs + "Default").LastOrDefault();
            if (lastDefault is not null)
                lastDefault.AddAfterSelf(newDefault);
            else
                root.AddFirst(newDefault);
        }
        else if (existing.Attribute("ContentType")?.Value != contentType)
        {
            existing.SetAttributeValue("ContentType", contentType);
        }

        using var outMs = new MemoryStream();
        doc.Save(outMs);
        package.SetPart("[Content_Types].xml", outMs.ToArray());
    }

    /// <summary>
    /// Ensures a relationship of the given type exists from <paramref name="fromPartName"/>
    /// to <paramref name="target"/>. Does nothing if the relationship already exists.
    /// </summary>
    internal static void EnsureRelationship(OpcPackage package, string fromPartName,
        string relType, string target)
    {
        var relsPath = GetRelsPartName(fromPartName);
        var rels = new RelsManager();
        var data = package.GetPart(relsPath);
        if (data is not null)
            rels.Load(data);

        if (rels.FindByType(relType).Any())
            return;

        rels.Add(relType, target);
        package.SetPart(relsPath, rels.ToBytes());
    }

    private static string GetRelsPartName(string partName)
    {
        if (partName.Contains('/'))
        {
            var dir = partName[..partName.LastIndexOf('/')];
            var file = partName[(partName.LastIndexOf('/') + 1)..];
            return $"{dir}/_rels/{file}.rels";
        }
        return $"_rels/{partName}.rels";
    }
}
