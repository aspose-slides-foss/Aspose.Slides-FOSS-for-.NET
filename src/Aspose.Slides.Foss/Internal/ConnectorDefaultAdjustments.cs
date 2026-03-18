namespace Aspose.Slides.Foss.Internal;

/// <summary>
/// Default adjustment values for connector preset geometries.
/// Values are in OOXML 1/100,000th percentage units (e.g., 50000 = 50%).
/// </summary>
internal static class ConnectorDefaultAdjustments
{
    /// <summary>
    /// Maps connector preset type names to their default adjustment values.
    /// Each entry is an array of (name, value) tuples.
    /// </summary>
    internal static readonly Dictionary<string, (string Name, long Value)[]> Defaults = new(StringComparer.Ordinal)
    {
        ["straightConnector1"] = [],
        ["bentConnector2"] = [],
        ["bentConnector3"] = [("adj1", 50000)],
        ["bentConnector4"] = [("adj1", 50000), ("adj2", 50000)],
        ["bentConnector5"] = [("adj1", 50000), ("adj2", 50000), ("adj3", 50000)],
        ["curvedConnector2"] = [],
        ["curvedConnector3"] = [("adj1", 50000)],
        ["curvedConnector4"] = [("adj1", 50000), ("adj2", 50000)],
        ["curvedConnector5"] = [("adj1", 50000), ("adj2", 50000), ("adj3", 50000)],
    };
}
