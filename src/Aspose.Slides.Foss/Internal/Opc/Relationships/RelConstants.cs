using System.Collections.Frozen;

namespace Aspose.Slides.Foss.Internal.Opc.Relationships;

/// <summary>
/// Constants for OPC relationship namespaces and common relationship types.
/// </summary>
public static class RelConstants
{
    /// <summary>
    /// The OPC relationships XML namespace URI.
    /// </summary>
    public const string RelsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";

    /// <summary>
    /// The OPC relationships namespace formatted for use in <see cref="System.Xml.Linq.XName"/> lookups.
    /// </summary>
    public const string RelsNs = "{" + RelsNamespace + "}";

    /// <summary>
    /// Common OPC/OOXML relationship type URIs keyed by a short friendly name.
    /// </summary>
    public static readonly FrozenDictionary<string, string> RelTypes = new Dictionary<string, string>
    {
        ["office_document"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
        ["slide"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
        ["slide_layout"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
        ["slide_master"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
        ["notes_slide"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide",
        ["notes_master"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster",
        ["handout_master"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster",
        ["theme"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
        ["core_properties"] = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
        ["extended_properties"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
        ["thumbnail"] = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail",
        ["image"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        ["hyperlink"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        ["chart"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
        ["oleObject"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject",
        ["package"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
        ["audio"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio",
        ["video"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video",
        ["comments"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
        ["commentAuthors"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors",
    }.ToFrozenDictionary();
}
