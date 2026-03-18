using System.Globalization;
using System.Xml.Linq;

namespace Aspose.Slides.Foss.Internal;

/// <summary>
/// Manages the OPC custom properties part (docProps/custom.xml).
/// </summary>
internal sealed class CustomPropertiesPart
{
    private static readonly XNamespace CustNs = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
    private static readonly XNamespace VtNs = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

    private static readonly string PartPath = "docProps/custom.xml";

    private readonly XElement _root;
    private bool _isDirty;

    internal bool IsDirty => _isDirty;

    internal CustomPropertiesPart(XElement root)
    {
        _root = root;
    }

    internal static CustomPropertiesPart CreateFromPackage(OpcPackage package)
    {
        var data = package.GetPart(PartPath);
        if (data is not null)
        {
            using var stream = new MemoryStream(data);
            var doc = XDocument.Load(stream);
            return new CustomPropertiesPart(doc.Root ?? throw new InvalidOperationException("Custom properties XML has no root element"));
        }

        var root = new XElement(CustNs + "Properties",
            new XAttribute(XNamespace.Xmlns + "cust", CustNs),
            new XAttribute(XNamespace.Xmlns + "vt", VtNs));
        return new CustomPropertiesPart(root);
    }

    internal void MarkDirty() => _isDirty = true;

    internal int Count => _root.Elements(CustNs + "property").Count();

    internal string GetNameAt(int index)
    {
        var props = _root.Elements(CustNs + "property").ToList();
        if (index < 0 || index >= props.Count)
            throw new ArgumentOutOfRangeException(nameof(index));
        return props[index].Attribute("name")?.Value ?? string.Empty;
    }

    internal bool Contains(string name)
    {
        return FindProperty(name) is not null;
    }

    internal object? GetValue(string name)
    {
        var prop = FindProperty(name);
        if (prop is null)
            return null;

        var vtChild = prop.Elements().FirstOrDefault(e => e.Name.Namespace == VtNs);
        if (vtChild is null)
            return null;

        return ParseVtValue(vtChild);
    }

    internal void SetValue(string name, object value)
    {
        var prop = FindProperty(name);
        if (prop is null)
        {
            prop = new XElement(CustNs + "property",
                new XAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"),
                new XAttribute("pid", GetNextPid()),
                new XAttribute("name", name));
            _root.Add(prop);
        }
        else
        {
            prop.Elements().Where(e => e.Name.Namespace == VtNs).Remove();
        }

        prop.Add(CreateVtElement(value));
        _isDirty = true;
    }

    internal bool Remove(string name)
    {
        var prop = FindProperty(name);
        if (prop is null)
            return false;

        prop.Remove();
        _isDirty = true;
        return true;
    }

    internal void Clear()
    {
        _root.Elements(CustNs + "property").Remove();
        _isDirty = true;
    }

    internal XElement Root => _root;

    private XElement? FindProperty(string name)
    {
        return _root.Elements(CustNs + "property")
            .FirstOrDefault(e => e.Attribute("name")?.Value == name);
    }

    private int GetNextPid()
    {
        int maxPid = 1;
        foreach (var prop in _root.Elements(CustNs + "property"))
        {
            if (int.TryParse(prop.Attribute("pid")?.Value, CultureInfo.InvariantCulture, out var pid) && pid > maxPid)
                maxPid = pid;
        }
        return maxPid + 1;
    }

    private static object? ParseVtValue(XElement vtElem)
    {
        var localName = vtElem.Name.LocalName;
        var text = vtElem.Value;

        return localName switch
        {
            "lpwstr" or "lpstr" => text,
            "i4" or "int" => int.TryParse(text, CultureInfo.InvariantCulture, out var i) ? i : 0,
            "r8" => double.TryParse(text, CultureInfo.InvariantCulture, out var d) ? d : 0.0,
            "bool" => text.Equals("true", StringComparison.OrdinalIgnoreCase) || text == "1",
            "filetime" => DateTime.TryParse(text, CultureInfo.InvariantCulture,
                DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out var dt) ? dt : null,
            _ => text
        };
    }

    private static XElement CreateVtElement(object value)
    {
        return value switch
        {
            string s => new XElement(VtNs + "lpwstr", s),
            int i => new XElement(VtNs + "i4", i.ToString(CultureInfo.InvariantCulture)),
            float f => new XElement(VtNs + "r8", f.ToString(CultureInfo.InvariantCulture)),
            double d => new XElement(VtNs + "r8", d.ToString(CultureInfo.InvariantCulture)),
            bool b => new XElement(VtNs + "bool", b ? "true" : "false"),
            DateTime dt => new XElement(VtNs + "filetime",
                dt.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture)),
            _ => new XElement(VtNs + "lpwstr", value.ToString() ?? string.Empty)
        };
    }
}
