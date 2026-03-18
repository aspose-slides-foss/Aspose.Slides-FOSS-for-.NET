using System.Globalization;
using System.Xml.Linq;
using Aspose.Slides.Foss.Drawing;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents the size of a notes slide.
/// </summary>
public sealed class NotesSize : INotesSize
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private const float EmuPerPoint = 12700f;

    private PresentationPart? _presentationPart;

    /// <summary>
    /// Initializes internal state. Called after default construction to inject dependencies.
    /// </summary>
    internal void InitInternal(PresentationPart presentationPart)
    {
        _presentationPart = presentationPart;
    }

    /// <inheritdoc />
    public SizeF Size
    {
        get
        {
            var el = _presentationPart?.Document.Root?.Element(PNs + "notesSz");
            if (el is null)
                return new SizeF();

            var cx = float.Parse(el.Attribute("cx")?.Value ?? "0", CultureInfo.InvariantCulture);
            var cy = float.Parse(el.Attribute("cy")?.Value ?? "0", CultureInfo.InvariantCulture);
            return new SizeF(cx / EmuPerPoint, cy / EmuPerPoint);
        }
        set
        {
            // Late-init: _presentationPart is set via InitInternal before use
            var root = _presentationPart!.Document.Root
                ?? throw new InvalidOperationException("Presentation XML has no root element");
            var el = root.Element(PNs + "notesSz");
            var cx = ((long)(value.Width * EmuPerPoint)).ToString(CultureInfo.InvariantCulture);
            var cy = ((long)(value.Height * EmuPerPoint)).ToString(CultureInfo.InvariantCulture);

            if (el is null)
            {
                el = new XElement(PNs + "notesSz", new XAttribute("cx", cx), new XAttribute("cy", cy));
                root.Add(el);
            }
            else
            {
                el.SetAttributeValue("cx", cx);
                el.SetAttributeValue("cy", cy);
            }
        }
    }
}
