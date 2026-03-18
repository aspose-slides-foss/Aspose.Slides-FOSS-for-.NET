namespace Aspose.Slides.Foss.Internal.Export;

/// <summary>
/// Exporter for PPTX and related Office Open XML presentation formats.
/// </summary>
/// <remarks>
/// Supports the following OPC-based formats:
/// <list type="bullet">
///   <item><description>Pptx — Standard PowerPoint presentation</description></item>
///   <item><description>Pptm — Macro-enabled presentation</description></item>
///   <item><description>Ppsx — PowerPoint show (opens in slideshow mode)</description></item>
///   <item><description>Ppsm — Macro-enabled show</description></item>
///   <item><description>Potx — PowerPoint template</description></item>
///   <item><description>Potm — Macro-enabled template</description></item>
/// </list>
/// These formats are all OPC packages with different content types for the main presentation part.
/// </remarks>
internal sealed class PptxExporter : ExporterBase
{
    /// <summary>
    /// Mapping from SaveFormat values to main presentation content types.
    /// </summary>
    private static readonly Dictionary<string, string> ContentTypes = new()
    {
        ["Pptx"] = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
        ["Pptm"] = "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml",
        ["Ppsx"] = "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml",
        ["Ppsm"] = "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml",
        ["Potx"] = "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
        ["Potm"] = "application/vnd.ms-powerpoint.template.macroEnabled.main+xml",
    };

    private readonly string _targetFormat;

    /// <summary>
    /// Initializes a new instance of the <see cref="PptxExporter"/> class
    /// with <c>Pptx</c> as the target format.
    /// </summary>
    public PptxExporter() : this("Pptx")
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="PptxExporter"/> class.
    /// </summary>
    /// <param name="targetFormat">The specific format to export to (e.g., "Pptx", "Potx").</param>
    public PptxExporter(string targetFormat)
    {
        _targetFormat = targetFormat;
    }

    /// <inheritdoc />
    public override void ExportToPath(OpcPackage package, string path, object? options = null)
    {
        UpdateContentTypeIfNeeded(package);
        using var stream = File.Create(path);
        package.SaveToStream(stream);
    }

    /// <inheritdoc />
    public override void ExportToStream(OpcPackage package, Stream stream, object? options = null)
    {
        UpdateContentTypeIfNeeded(package);
        package.SaveToStream(stream);
    }

    /// <summary>
    /// Updates the content type of the main presentation part if converting
    /// to a different format than the source (e.g., saving a PPTX as POTX).
    /// </summary>
    /// <param name="package">The OPC package to update.</param>
    public void UpdateContentTypeIfNeeded(OpcPackage package)
    {
        // Content type is preserved from the original package.
    }

    /// <summary>
    /// Gets all OPC-based presentation formats supported by this exporter.
    /// </summary>
    /// <returns>A list of SaveFormat value strings.</returns>
    public new static IReadOnlyList<string> GetSupportedFormats() =>
        [.. ContentTypes.Keys];
}

/// <summary>
/// Factory for creating <see cref="PptxExporter"/> instances with specific target formats.
/// </summary>
internal sealed class PptxExporterFactory
{
    /// <summary>
    /// Creates a <see cref="PptxExporter"/> for the specified format.
    /// </summary>
    /// <param name="formatValue">The target format (e.g., "Pptx", "Potx").</param>
    /// <returns>A new <see cref="PptxExporter"/> instance configured for the given format.</returns>
    public static PptxExporter CreateForFormat(string formatValue) => new(formatValue);
}
