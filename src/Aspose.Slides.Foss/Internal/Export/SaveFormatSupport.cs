using Aspose.Slides.Foss.Export;

namespace Aspose.Slides.Foss.Internal.Export;

/// <summary>
/// Decides what a requested <see cref="SaveFormat"/> means for the written package.
/// </summary>
/// <remarks>
/// <para>
/// Every format this library can write is an Office Open XML package; the formats differ only in the
/// content type declared for <c>/ppt/presentation.xml</c> in <c>[Content_Types].xml</c>. Per
/// ISO/IEC 29500-2 §10.1.2 that content type is the part's identity, and PowerPoint cross-checks it
/// against the file extension: a package whose main part claims to be a presentation, saved under a
/// <c>.potx</c> name, is refused with <i>"PowerPoint can't open this file because its file extension
/// has changed"</i>.
/// </para>
/// <para>
/// Anything that is not an Office Open XML package — PDF, HTML, images, the legacy binary formats —
/// would need a renderer or a different container, and this library has neither. Those formats are
/// refused rather than silently answered with a presentation package under a misleading name.
/// </para>
/// </remarks>
internal static class SaveFormatSupport
{
    /// <summary>The part whose content type identifies the package format.</summary>
    internal const string MainPartName = "ppt/presentation.xml";

    private static readonly Dictionary<SaveFormat, string> MainPartContentTypes = new()
    {
        [SaveFormat.Pptx] = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
        [SaveFormat.Ppsx] = "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml",
        [SaveFormat.Potx] = "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
    };

    /// <summary>
    /// Macro-enabled formats. These are Office Open XML packages too, but a package that declares a
    /// macro-enabled content type without carrying a <c>ppt/vbaProject.bin</c> part misrepresents
    /// itself, so they are refused until VBA parts are supported.
    /// </summary>
    private static readonly SaveFormat[] MacroEnabledFormats =
        [SaveFormat.Pptm, SaveFormat.Ppsm, SaveFormat.Potm];

    /// <summary>
    /// Gets the content type to declare for the main presentation part.
    /// </summary>
    /// <param name="format">The requested save format.</param>
    /// <returns>The content type for <c>/ppt/presentation.xml</c>.</returns>
    /// <exception cref="NotSupportedException">
    /// Thrown when <paramref name="format"/> cannot be written by this library.
    /// </exception>
    internal static string MainPartContentTypeFor(SaveFormat format)
    {
        if (MainPartContentTypes.TryGetValue(format, out var contentType))
            return contentType;

        throw new NotSupportedException(BuildRefusal(format));
    }

    /// <summary>
    /// Declares the main presentation part with the content type the requested format calls for.
    /// </summary>
    /// <param name="package">The package about to be written.</param>
    /// <param name="format">The requested save format.</param>
    /// <exception cref="NotSupportedException">
    /// Thrown when <paramref name="format"/> cannot be written by this library.
    /// </exception>
    internal static void ApplyTo(OpcPackage package, SaveFormat format)
    {
        OpcRegistration.AddContentTypeOverride(package, MainPartName, MainPartContentTypeFor(format));
    }

    private static string BuildRefusal(SaveFormat format)
    {
        var supported = string.Join(", ", MainPartContentTypes.Keys);

        if (Array.IndexOf(MacroEnabledFormats, format) >= 0)
        {
            return $"Save format '{format}' is not supported: a macro-enabled package requires a VBA " +
                   $"project part, which this library does not write. Supported formats: {supported}.";
        }

        return $"Save format '{format}' is not supported: it is not an Office Open XML presentation " +
               $"package and this library does not render or convert. Supported formats: {supported}.";
    }
}
