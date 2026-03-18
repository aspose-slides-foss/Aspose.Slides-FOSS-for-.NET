using System.Text.RegularExpressions;
using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents an AutoShape — a preset or custom geometric shape that may contain text.
/// </summary>
public sealed class AutoShape : GeometryShape, IAutoShape
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// <inheritdoc/>
    public override ShapeType ShapeType
    {
        get
        {
            if (_element is null)
                return ShapeType.NotDefined;

            var spPr = _element.Element(PNs + "spPr");
            if (spPr is null)
                return ShapeType.NotDefined;

            var custGeom = spPr.Element(ANs + "custGeom");
            if (custGeom is not null)
                return ShapeType.Custom;

            var prstGeom = spPr.Element(ANs + "prstGeom");
            if (prstGeom is null)
                return ShapeType.NotDefined;

            var prst = prstGeom.Attribute("prst")?.Value;
            return OoxmlPresetMapping.FromPreset(prst);
        }
        set
        {
            if (_element is null)
                return;
            if (value == ShapeType.NotDefined || value == ShapeType.Custom)
                return;

            var preset = OoxmlPresetMapping.ToPreset(value);
            if (preset is null)
                return;

            var spPr = _element.Element(PNs + "spPr");
            if (spPr is null)
            {
                spPr = new XElement(PNs + "spPr");
                _element.Add(spPr);
            }

            var custGeom = spPr.Element(ANs + "custGeom");
            custGeom?.Remove();

            var prstGeom = spPr.Element(ANs + "prstGeom");
            if (prstGeom is null)
            {
                prstGeom = new XElement(ANs + "prstGeom");
                spPr.Add(prstGeom);
            }
            else
            {
                prstGeom.RemoveNodes();
            }

            prstGeom.SetAttributeValue("prst", preset);

            _slidePart?.Save();
        }
    }

    /// <inheritdoc/>
    public ITextFrame? TextFrame
    {
        get
        {
            if (_element is null)
                return null;

            var txBody = _element.Element(PNs + "txBody");
            if (txBody is null)
                return null;

            var tf = new TextFrame();
            tf.InitInternal(txBody, _slidePart, _parentSlide, this);
            return tf;
        }
    }

    /// <inheritdoc/>
    public bool IsTextBox
    {
        get
        {
            if (_element is null)
                return false;

            return _element.Element(PNs + "txBody") is not null;
        }
    }

    /// <inheritdoc/>
    public IGeometryShape AsIGeometryShape => this;

    /// <inheritdoc/>
    public ITextFrame? AddTextFrame(string? text)
    {
        if (_element is null)
            return null;

        _element.Element(PNs + "txBody")?.Remove();

        // Mark as text box
        var nvSpPr = _element.Element(PNs + "nvSpPr");
        if (nvSpPr is not null)
        {
            var cNvSpPr = nvSpPr.Element(PNs + "cNvSpPr");
            if (cNvSpPr is not null)
            {
                cNvSpPr.SetAttributeValue("txBox", "1");
            }
        }

        var txBody = new XElement(PNs + "txBody",
            new XElement(ANs + "bodyPr",
                new XAttribute("rtlCol", "0"),
                new XAttribute("anchor", "ctr")),
            new XElement(ANs + "lstStyle"));

        var lines = string.IsNullOrEmpty(text)
            ? [""]
            : Regex.Split(text, @"\r\n|\r|\n");

        foreach (var line in lines)
        {
            txBody.Add(new XElement(ANs + "p",
                new XElement(ANs + "pPr", new XAttribute("algn", "ctr")),
                new XElement(ANs + "r",
                    new XElement(ANs + "t", line))));
        }

        _element.Add(txBody);

        _slidePart?.Save();

        var tf = new TextFrame();
        tf.InitInternal(txBody, _slidePart, _parentSlide, this);
        return tf;
    }
}
