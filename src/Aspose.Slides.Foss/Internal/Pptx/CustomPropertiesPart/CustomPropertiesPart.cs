using System.Globalization;
using System.Xml.Linq;

namespace Aspose.Slides.Foss.Internal.Pptx.CustomPropertiesPart;

/// <summary>
/// Parses and serializes docProps/custom.xml.
/// Each custom property is stored with a typed value (string, int, double, bool, or DateTime).
/// PIDs start at 2. The file is created on demand only when custom properties are added.
/// </summary>
public sealed class CustomPropertiesPart
{
    private static readonly XNamespace CustomNs = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
    private static readonly XNamespace VtNs = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

    private const string Fmtid = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

    /// <summary>
    /// The part path within the OPC package.
    /// </summary>
    public const string PartName = "docProps/custom.xml";

    private readonly Internal.OpcPackage _package;
    private readonly Dictionary<string, object> _properties = [];
    private bool _dirty;

    /// <summary>
    /// Initializes the custom properties part by parsing docProps/custom.xml from the package.
    /// </summary>
    /// <param name="package">The OPC package containing the presentation.</param>
    internal CustomPropertiesPart(Internal.OpcPackage package)
    {
        _package = package;
        Parse();
    }

    /// <summary>
    /// Gets the number of custom properties.
    /// </summary>
    public int Count => _properties.Count;

    /// <summary>
    /// Determines whether a custom property with the specified name exists.
    /// </summary>
    /// <param name="name">The property name.</param>
    /// <returns><c>true</c> if the property exists; otherwise, <c>false</c>.</returns>
    public bool Contains(string name) => _properties.ContainsKey(name);

    /// <summary>
    /// Parses the docProps/custom.xml part from the package.
    /// </summary>
    public void Parse()
    {
        var data = _package.GetPart(PartName);
        if (data is null)
            return;

        using var stream = new MemoryStream(data);
        var doc = XDocument.Load(stream);
        var root = doc.Root;
        if (root is null)
            return;

        foreach (var propEl in root.Elements(CustomNs + "property"))
        {
            var name = propEl.Attribute("name")?.Value;
            if (name is null)
                continue;
            var value = ReadValue(propEl);
            if (value is not null)
                _properties[name] = value;
        }
    }

    /// <summary>
    /// Reads a typed value from a property element.
    /// </summary>
    /// <param name="propEl">The property XML element.</param>
    /// <returns>The typed value, or <c>null</c> if no recognized value type is found.</returns>
    public object? ReadValue(XElement propEl)
    {
        foreach (var child in propEl.Elements())
        {
            var tag = child.Name;
            var text = child.Value ?? "";

            if (tag == VtNs + "lpwstr")
                return text;

            if (tag == VtNs + "i4")
                return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var i) ? i : 0;

            if (tag == VtNs + "r8")
                return double.TryParse(text, NumberStyles.Float | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var d) ? d : 0.0;

            if (tag == VtNs + "bool")
                return text.Equals("true", StringComparison.OrdinalIgnoreCase) || text == "1";

            if (tag == VtNs + "filetime")
                return ParseFiletime(text);
        }

        return null;
    }

    /// <summary>
    /// Parses a filetime string (ISO 8601 format in OOXML) to a <see cref="DateTime"/> value.
    /// </summary>
    /// <param name="text">The filetime string to parse.</param>
    /// <returns>The parsed <see cref="DateTime"/>, or <c>null</c> if the input is null, empty, or invalid.</returns>
    public static DateTime? ParseFiletime(string? text)
    {
        if (string.IsNullOrEmpty(text))
            return null;

        try
        {
            if (text.EndsWith('Z'))
                text = text[..^1] + "+00:00";

            return DateTime.Parse(text, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);
        }
        catch (FormatException)
        {
            return null;
        }
    }

    /// <summary>
    /// Gets the name of a custom property by its index.
    /// </summary>
    /// <param name="index">The zero-based index.</param>
    /// <returns>The property name.</returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when the index is out of range.</exception>
    public string GetName(int index)
    {
        var names = _properties.Keys.ToList();
        if (index < 0 || index >= names.Count)
            throw new ArgumentOutOfRangeException(nameof(index), $"Index {index} out of range");
        return names[index];
    }

    /// <summary>
    /// Gets the value of a custom property by name.
    /// </summary>
    /// <param name="name">The property name.</param>
    /// <returns>The property value, or <c>null</c> if not found.</returns>
    public object? GetValue(string name)
    {
        return _properties.TryGetValue(name, out var value) ? value : null;
    }

    /// <summary>
    /// Sets a custom property value. Creates the property if it does not exist.
    /// </summary>
    /// <param name="name">The property name.</param>
    /// <param name="value">The property value (string, int, double, bool, or DateTime).</param>
    public void SetValue(string name, object value)
    {
        _properties[name] = value;
        _dirty = true;
    }

    /// <summary>
    /// Removes a custom property by name.
    /// </summary>
    /// <param name="name">The property name.</param>
    /// <returns><c>true</c> if the property was removed; <c>false</c> if it did not exist.</returns>
    public bool Remove(string name)
    {
        if (_properties.Remove(name))
        {
            _dirty = true;
            return true;
        }
        return false;
    }

    /// <summary>
    /// Removes all custom properties.
    /// </summary>
    public void Clear()
    {
        if (_properties.Count > 0)
        {
            _properties.Clear();
            _dirty = true;
        }
    }

    /// <summary>
    /// Serializes the custom properties back to the OPC package.
    /// Only writes if properties have been modified. Removes the part if no properties remain.
    /// </summary>
    public void Save()
    {
        if (!_dirty)
            return;

        if (_properties.Count == 0)
        {
            _package.RemovePart(PartName);
            _dirty = false;
            return;
        }

        var root = new XElement(CustomNs + "Properties",
            new XAttribute("xmlns", CustomNs.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "vt", VtNs));

        var pid = 2;
        foreach (var (name, value) in _properties)
        {
            var propEl = new XElement(CustomNs + "property",
                new XAttribute("fmtid", Fmtid),
                new XAttribute("pid", pid.ToString(CultureInfo.InvariantCulture)),
                new XAttribute("name", name));

            WriteValue(propEl, value);
            root.Add(propEl);
            pid++;
        }

        var doc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
        using var ms = new MemoryStream();
        doc.Save(ms);
        _package.SetPart(PartName, ms.ToArray());

        EnsureContentType();
        EnsureRelationship();
        _dirty = false;
    }

    /// <summary>
    /// Writes a typed value element to a property element.
    /// </summary>
    /// <param name="propEl">The property element to write to.</param>
    /// <param name="value">The typed value to write.</param>
    public void WriteValue(XElement propEl, object value)
    {
        switch (value)
        {
            case bool b:
                propEl.Add(new XElement(VtNs + "bool", b ? "true" : "false"));
                break;
            case int i:
                propEl.Add(new XElement(VtNs + "i4", i.ToString(CultureInfo.InvariantCulture)));
                break;
            case double d:
                propEl.Add(new XElement(VtNs + "r8", d.ToString(CultureInfo.InvariantCulture)));
                break;
            case DateTime dt:
                var utcDt = dt.Kind == DateTimeKind.Unspecified
                    ? DateTime.SpecifyKind(dt, DateTimeKind.Utc)
                    : dt;
                propEl.Add(new XElement(VtNs + "filetime",
                    utcDt.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture)));
                break;
            default:
                propEl.Add(new XElement(VtNs + "lpwstr", value.ToString()));
                break;
        }
    }

    /// <summary>
    /// Ensures [Content_Types].xml has an override entry for docProps/custom.xml.
    /// </summary>
    public void EnsureContentType()
    {
        var ctData = _package.GetPart("[Content_Types].xml");
        if (ctData is null)
            return;

        using var msCt = new MemoryStream(ctData);
        var ctRoot = XElement.Load(msCt);
        XNamespace ctNs = "http://schemas.openxmlformats.org/package/2006/content-types";

        foreach (var overrideEl in ctRoot.Elements(ctNs + "Override"))
        {
            if (overrideEl.Attribute("PartName")?.Value == "/docProps/custom.xml")
                return;
        }

        ctRoot.Add(new XElement(ctNs + "Override",
            new XAttribute("PartName", "/docProps/custom.xml"),
            new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.custom-properties+xml")));

        var doc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), ctRoot);
        using var ms = new MemoryStream();
        doc.Save(ms);
        _package.SetPart("[Content_Types].xml", ms.ToArray());
    }

    /// <summary>
    /// Ensures _rels/.rels has a relationship entry for docProps/custom.xml.
    /// </summary>
    public void EnsureRelationship()
    {
        var relsData = _package.GetPart("_rels/.rels");
        if (relsData is null)
            return;

        using var msRels = new MemoryStream(relsData);
        var relsRoot = XElement.Load(msRels);
        const string relType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";

        var maxId = 0;
        foreach (var rel in relsRoot.Elements())
        {
            if (rel.Attribute("Type")?.Value == relType)
                return;

            var rid = rel.Attribute("Id")?.Value ?? "";
            if (rid.StartsWith("rId", StringComparison.Ordinal) && int.TryParse(rid[3..], out var num))
                maxId = Math.Max(maxId, num);
        }

        relsRoot.Add(new XElement("Relationship",
            new XAttribute("Id", $"rId{maxId + 1}"),
            new XAttribute("Type", relType),
            new XAttribute("Target", "docProps/custom.xml")));

        var doc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), relsRoot);
        using var ms = new MemoryStream();
        doc.Save(ms);
        _package.SetPart("_rels/.rels", ms.ToArray());
    }
}
