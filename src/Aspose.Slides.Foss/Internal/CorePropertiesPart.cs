using System.Globalization;
using System.Xml.Linq;

namespace Aspose.Slides.Foss.Internal;

/// <summary>
/// Manages the OPC core properties part (docProps/core.xml).
/// </summary>
internal sealed class CorePropertiesPart
{
    private static readonly XNamespace CpNs = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private static readonly XNamespace DcNs = "http://purl.org/dc/elements/1.1/";
    private static readonly XNamespace DcTermsNs = "http://purl.org/dc/terms/";
    private static readonly XNamespace XsiNs = "http://www.w3.org/2001/XMLSchema-instance";

    private static readonly string PartPath = "docProps/core.xml";

    private readonly XElement _root;
    private bool _isDirty;

    internal bool IsDirty => _isDirty;

    internal CorePropertiesPart(XElement root)
    {
        _root = root;
    }

    internal static CorePropertiesPart CreateFromPackage(OpcPackage package)
    {
        var data = package.GetPart(PartPath);
        if (data is not null)
        {
            using var stream = new MemoryStream(data);
            var doc = XDocument.Load(stream);
            return new CorePropertiesPart(doc.Root ?? throw new InvalidOperationException("Core properties XML has no root element"));
        }

        var root = new XElement(CpNs + "coreProperties",
            new XAttribute(XNamespace.Xmlns + "cp", CpNs),
            new XAttribute(XNamespace.Xmlns + "dc", DcNs),
            new XAttribute(XNamespace.Xmlns + "dcterms", DcTermsNs),
            new XAttribute(XNamespace.Xmlns + "xsi", XsiNs));
        return new CorePropertiesPart(root);
    }

    internal void MarkDirty() => _isDirty = true;

    internal string? Title
    {
        get => GetDcElement("title");
        set => SetDcElement("title", value);
    }

    internal string? Subject
    {
        get => GetDcElement("subject");
        set => SetDcElement("subject", value);
    }

    internal string? Creator
    {
        get => GetDcElement("creator");
        set => SetDcElement("creator", value);
    }

    internal string? Keywords
    {
        get => GetCpElement("keywords");
        set => SetCpElement("keywords", value);
    }

    internal string? Description
    {
        get => GetDcElement("description");
        set => SetDcElement("description", value);
    }

    internal string? Category
    {
        get => GetCpElement("category");
        set => SetCpElement("category", value);
    }

    internal string? ContentStatus
    {
        get => GetCpElement("contentStatus");
        set => SetCpElement("contentStatus", value);
    }

    internal string? ContentType
    {
        get => GetCpElement("contentType");
        set => SetCpElement("contentType", value);
    }

    internal string? LastModifiedBy
    {
        get => GetCpElement("lastModifiedBy");
        set => SetCpElement("lastModifiedBy", value);
    }

    internal string? Revision
    {
        get => GetCpElement("revision");
        set => SetCpElement("revision", value);
    }

    internal DateTime? Created
    {
        get => GetDcTermsDateTime("created");
        set => SetDcTermsDateTime("created", value);
    }

    internal DateTime? Modified
    {
        get => GetDcTermsDateTime("modified");
        set => SetDcTermsDateTime("modified", value);
    }

    internal DateTime? LastPrinted
    {
        get => GetCpDateTime("lastPrinted");
        set => SetCpDateTime("lastPrinted", value);
    }

    internal XElement Root => _root;

    private string? GetDcElement(string name) => _root.Element(DcNs + name)?.Value;

    private void SetDcElement(string name, string? value)
    {
        var elem = _root.Element(DcNs + name);
        if (elem is null)
        {
            elem = new XElement(DcNs + name);
            _root.Add(elem);
        }
        elem.Value = value ?? string.Empty;
    }

    private string? GetCpElement(string name) => _root.Element(CpNs + name)?.Value;

    private void SetCpElement(string name, string? value)
    {
        var elem = _root.Element(CpNs + name);
        if (elem is null)
        {
            elem = new XElement(CpNs + name);
            _root.Add(elem);
        }
        elem.Value = value ?? string.Empty;
    }

    private DateTime? GetDcTermsDateTime(string name)
    {
        var text = _root.Element(DcTermsNs + name)?.Value;
        if (string.IsNullOrEmpty(text))
            return null;
        return DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out var dt)
            ? dt
            : null;
    }

    private void SetDcTermsDateTime(string name, DateTime? value)
    {
        var elem = _root.Element(DcTermsNs + name);
        if (value is null)
        {
            elem?.Remove();
            return;
        }

        if (elem is null)
        {
            elem = new XElement(DcTermsNs + name,
                new XAttribute(XsiNs + "type", "dcterms:W3CDTF"));
            _root.Add(elem);
        }
        elem.Value = value.Value.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);
    }

    private DateTime? GetCpDateTime(string name)
    {
        var text = _root.Element(CpNs + name)?.Value;
        if (string.IsNullOrEmpty(text))
            return null;
        return DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out var dt)
            ? dt
            : null;
    }

    private void SetCpDateTime(string name, DateTime? value)
    {
        var elem = _root.Element(CpNs + name);
        if (value is null)
        {
            elem?.Remove();
            return;
        }

        if (elem is null)
        {
            elem = new XElement(CpNs + name);
            _root.Add(elem);
        }
        elem.Value = value.Value.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);
    }
}
