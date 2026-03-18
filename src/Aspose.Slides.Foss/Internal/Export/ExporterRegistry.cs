using System.Reflection;

namespace Aspose.Slides.Foss.Internal.Export;

/// <summary>
/// Central registry for format exporters.
/// Maintains a mapping from SaveFormat value strings to their corresponding
/// exporter classes. New exporters can be registered dynamically.
/// </summary>
internal static class ExporterRegistry
{
    private static readonly Dictionary<string, Type> Exporters = new();

    /// <summary>
    /// Registers an exporter class for all formats it supports.
    /// The exporter type must define a static <c>GetSupportedFormats()</c> method.
    /// </summary>
    /// <param name="exporterType">The exporter type to register.</param>
    /// <exception cref="ArgumentException">
    /// Thrown when <paramref name="exporterType"/> does not derive from
    /// <see cref="ExporterBase"/> or lacks a static <c>GetSupportedFormats</c> method.
    /// </exception>
    public static void Register(Type exporterType)
    {
        if (!typeof(ExporterBase).IsAssignableFrom(exporterType))
        {
            throw new ArgumentException(
                $"Type {exporterType.Name} must derive from ExporterBase.",
                nameof(exporterType));
        }

        var formats = GetSupportedFormatsFromType(exporterType);
        foreach (var format in formats)
        {
            Exporters[format] = exporterType;
        }
    }

    /// <summary>
    /// Unregisters the exporter for a specific format.
    /// </summary>
    /// <param name="formatValue">The SaveFormat value string to unregister.</param>
    /// <returns><c>true</c> if the format was unregistered; <c>false</c> if not found.</returns>
    public static bool Unregister(string formatValue) => Exporters.Remove(formatValue);

    /// <summary>
    /// Gets a new exporter instance for the specified format.
    /// </summary>
    /// <param name="formatValue">The SaveFormat value string (e.g., "Pptx", "Pdf").</param>
    /// <returns>An exporter instance, or <c>null</c> if no exporter is registered.</returns>
    public static ExporterBase? GetExporter(string formatValue)
    {
        var exporterType = GetExporterClass(formatValue);
        return exporterType is not null
            ? (ExporterBase)Activator.CreateInstance(exporterType)!
            : null;
    }

    /// <summary>
    /// Gets the exporter class (type) registered for a specific format.
    /// </summary>
    /// <param name="formatValue">The SaveFormat value string.</param>
    /// <returns>The exporter type, or <c>null</c> if not registered.</returns>
    public static Type? GetExporterClass(string formatValue) =>
        Exporters.GetValueOrDefault(formatValue);

    /// <summary>
    /// Checks whether a format has a registered exporter.
    /// </summary>
    /// <param name="formatValue">The SaveFormat value string.</param>
    /// <returns><c>true</c> if an exporter is registered for this format.</returns>
    public static bool IsFormatSupported(string formatValue) =>
        Exporters.ContainsKey(formatValue);

    /// <summary>
    /// Gets all format strings that have registered exporters.
    /// </summary>
    /// <returns>A list of registered SaveFormat value strings.</returns>
    public static List<string> GetSupportedFormats() => [.. Exporters.Keys];

    /// <summary>
    /// Clears all registered exporters. Intended for testing.
    /// </summary>
    public static void Clear() => Exporters.Clear();

    private static IReadOnlyList<string> GetSupportedFormatsFromType(Type type)
    {
        var method = type.GetMethod(
            nameof(ExporterBase.GetSupportedFormats),
            BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy);

        if (method is null)
        {
            throw new ArgumentException(
                $"Type {type.Name} does not have a static GetSupportedFormats method.",
                nameof(type));
        }

        return (IReadOnlyList<string>)method.Invoke(null, null)!;
    }
}
