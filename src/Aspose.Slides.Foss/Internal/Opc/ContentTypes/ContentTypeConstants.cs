using System.Collections.Frozen;

namespace Aspose.Slides.Foss.Internal.Opc.ContentTypes;

/// <summary>
/// Constants for OPC Content Types XML namespace and common PPTX content types.
/// </summary>
internal static class ContentTypeConstants
{
    /// <summary>
    /// The XML namespace URI for [Content_Types].xml.
    /// </summary>
    public const string CtNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";

    /// <summary>
    /// The XML namespace prefix string in <c>{uri}</c> format for element construction.
    /// </summary>
    public const string CtNs = "{" + CtNamespace + "}";

    /// <summary>
    /// Common PPTX content type mappings keyed by logical name.
    /// </summary>
    public static readonly FrozenDictionary<string, string> ContentTypes = new Dictionary<string, string>
    {
        ["presentation"] = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
        ["presentation_macro"] = "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml",
        ["slide"] = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
        ["slide_layout"] = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
        ["slide_master"] = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
        ["notes_slide"] = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
        ["notes_master"] = "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml",
        ["handout_master"] = "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml",
        ["theme"] = "application/vnd.openxmlformats-officedocument.theme+xml",
        ["core_properties"] = "application/vnd.openxmlformats-package.core-properties+xml",
        ["extended_properties"] = "application/vnd.openxmlformats-officedocument.extended-properties+xml",
        ["chart"] = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
        ["chartsheet"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml",
        ["comments"] = "application/vnd.openxmlformats-officedocument.presentationml.comments+xml",
        ["commentAuthors"] = "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml",
    }.ToFrozenDictionary();
}
