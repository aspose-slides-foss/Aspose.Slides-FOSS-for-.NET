using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a geometry shape adjustment value backed by an XML guide definition element.
/// </summary>
public sealed class AdjustValue : IAdjustValue
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private const float AngleFactor = 60000f;

    private XElement? _gdElement;
    private SlidePart? _slidePart;

    /// <summary>
    /// Initializes internal state. Called after default construction to inject dependencies.
    /// </summary>
    internal void InitInternal(XElement gdElement, SlidePart? slidePart)
    {
        _gdElement = gdElement;
        _slidePart = slidePart;
    }

    /// <inheritdoc/>
    public string Name => _gdElement?.Attribute("name")?.Value ?? string.Empty;

    /// <inheritdoc/>
    public int RawValue
    {
        get
        {
            if (_gdElement is null)
                return 0;

            var fmla = _gdElement.Attribute("fmla")?.Value;
            if (fmla is null || !fmla.StartsWith("val ", StringComparison.Ordinal))
                return 0;

            return int.TryParse(fmla.AsSpan(4), out var v) ? v : 0;
        }
        set
        {
            if (_gdElement is null)
                return;

            _gdElement.SetAttributeValue("fmla", $"val {value}");
            _slidePart?.Save();
        }
    }

    /// <inheritdoc/>
    public float AngleValue
    {
        get => RawValue / AngleFactor;
        set => RawValue = (int)Math.Round(value * AngleFactor, MidpointRounding.ToEven);
    }
}
