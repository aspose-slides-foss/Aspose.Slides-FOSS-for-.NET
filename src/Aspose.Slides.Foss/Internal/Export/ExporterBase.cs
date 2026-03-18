namespace Aspose.Slides.Foss.Internal.Export;

/// <summary>
/// Abstract base class for presentation format exporters.
/// Each exporter handles conversion from the internal presentation
/// representation to a specific output format (PPTX, PDF, HTML, etc.).
/// </summary>
internal abstract class ExporterBase
{
    /// <summary>
    /// Exports the presentation to a file path.
    /// </summary>
    /// <param name="package">The OPC package containing the presentation data.</param>
    /// <param name="path">The output file path.</param>
    /// <param name="options">Optional export options specific to the format.</param>
    /// <exception cref="IOException">If the file cannot be written.</exception>
    /// <exception cref="ArgumentException">If the options are invalid.</exception>
    public abstract void ExportToPath(OpcPackage package, string path, object? options = null);

    /// <summary>
    /// Exports the presentation to a binary stream.
    /// </summary>
    /// <param name="package">The OPC package containing the presentation data.</param>
    /// <param name="stream">The output stream with write capability.</param>
    /// <param name="options">Optional export options specific to the format.</param>
    /// <exception cref="IOException">If the stream cannot be written to.</exception>
    /// <exception cref="ArgumentException">If the options are invalid.</exception>
    public abstract void ExportToStream(OpcPackage package, Stream stream, object? options = null);

    /// <summary>
    /// Gets the list of SaveFormat values this exporter supports.
    /// </summary>
    /// <returns>List of SaveFormat enum value strings (e.g., ["Pptx", "Pptm"]).</returns>
    public static IReadOnlyList<string> GetSupportedFormats() => [];

    /// <summary>
    /// Exports to either a file path or stream.
    /// </summary>
    /// <param name="package">The OPC package containing the presentation data.</param>
    /// <param name="destination">File path (string) or binary stream.</param>
    /// <param name="options">Optional export options specific to the format.</param>
    public void Export(OpcPackage package, object destination, object? options = null)
    {
        if (destination is string path)
        {
            ExportToPath(package, path, options);
        }
        else if (destination is Stream stream)
        {
            ExportToStream(package, stream, options);
        }
        else
        {
            throw new ArgumentException("Destination must be a file path (string) or a Stream.", nameof(destination));
        }
    }
}
