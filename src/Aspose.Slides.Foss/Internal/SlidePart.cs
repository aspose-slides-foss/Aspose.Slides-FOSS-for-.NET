using System.Xml.Linq;

namespace Aspose.Slides.Foss.Internal;

/// <summary>
/// Manages a slide XML part within the presentation package.
/// </summary>
internal sealed class SlidePart
{
    private const string SlideLayoutRelType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";

    private readonly RelsManager _relsManager = new();
    private XElement? _element;
    private string _partName = string.Empty;

    /// <summary>
    /// Gets the relationships manager for this slide part.
    /// </summary>
    internal RelsManager RelsManager => _relsManager;

    /// <summary>
    /// Gets the part name (e.g., "ppt/slides/slide1.xml").
    /// </summary>
    internal string PartName => _partName;

    /// <summary>
    /// Gets or sets the XML element for this slide part.
    /// </summary>
    internal XElement? Element
    {
        get => _element;
        set => _element = value;
    }

    /// <summary>
    /// Gets or sets the OPC package reference for this slide part.
    /// </summary>
    internal OpcPackage? Package { get; set; }

    /// <summary>
    /// Gets the layout slide part name resolved from this slide's relationships.
    /// Returns <c>null</c> if no slide layout relationship is found.
    /// </summary>
    internal string? LayoutPartName
    {
        get
        {
            var layoutRel = _relsManager.FindByType(SlideLayoutRelType).FirstOrDefault();
            if (layoutRel is null)
                return null;

            return ResolvePartName(GetDirectoryPath(_partName), layoutRel.Target);
        }
    }

    /// <summary>
    /// Initializes the slide part with its part name.
    /// </summary>
    internal void InitInternal(string partName)
    {
        _partName = partName;
    }

    /// <summary>
    /// Persists changes to the underlying package.
    /// </summary>
    internal void Save()
    {
        if (_element is null || Package is null || string.IsNullOrEmpty(_partName))
            return;

        using var ms = new MemoryStream();
        using (var writer = new System.IO.StreamWriter(ms, new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false), leaveOpen: true))
        {
            var doc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), new XElement(_element));
            doc.Save(writer);
        }
        Package.SetPart(_partName, ms.ToArray());
    }

    private static string ResolvePartName(string basePath, string target)
    {
        var combined = basePath + target;
        var segments = combined.Split('/');
        var resolved = new Stack<string>();

        foreach (var segment in segments)
        {
            if (segment == "..")
            {
                if (resolved.Count > 0)
                    resolved.Pop();
            }
            else if (segment is not ("" or "."))
            {
                resolved.Push(segment);
            }
        }

        return string.Join("/", resolved.Reverse());
    }

    private static string GetDirectoryPath(string partName)
    {
        var lastSlash = partName.LastIndexOf('/');
        return lastSlash >= 0 ? partName[..(lastSlash + 1)] : "";
    }
}
