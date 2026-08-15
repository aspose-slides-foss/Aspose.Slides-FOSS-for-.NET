using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a picture frame shape.
/// </summary>
public sealed class PictureFrame : GeometryShape, IPictureFrame
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private float _relativeScaleHeight = 1.0f;
    private float _relativeScaleWidth = 1.0f;

    /// <inheritdoc/>
    public override ShapeType ShapeType
    {
        get
        {
            if (_element is null)
                return ShapeType.Rectangle;

            var spPr = _element.Element(PNs + "spPr");
            if (spPr is null)
                return ShapeType.Rectangle;

            var prstGeom = spPr.Element(ANs + "prstGeom");
            if (prstGeom is null)
                return ShapeType.Rectangle;

            var prst = prstGeom.Attribute("prst")?.Value;
            if (prst is null)
            {
                prstGeom.SetAttributeValue("prst", "rect");
                return ShapeType.Rectangle;
            }

            var result = OoxmlPresetMapping.FromPreset(prst);
            return result == ShapeType.NotDefined ? ShapeType.Rectangle : result;
        }
        set
        {
            if (_element is null)
                return;

            var spPr = _element.Element(PNs + "spPr");
            if (spPr is null)
                return;

            var preset = OoxmlPresetMapping.ToPreset(value) ?? "rect";

            var prstGeom = spPr.Element(ANs + "prstGeom");
            if (prstGeom is null)
            {
                prstGeom = new XElement(ANs + "prstGeom");
                spPr.Add(prstGeom);
            }

            prstGeom.SetAttributeValue("prst", preset);
        }
    }

    /// <inheritdoc/>
    public IPictureFrameLock? PictureFrameLock
    {
        get
        {
            if (_element is null)
                return null;

            var nvPicPr = _element.Element(PNs + "nvPicPr");
            if (nvPicPr is null)
                return null;

            var cNvPicPr = nvPicPr.Element(PNs + "cNvPicPr");
            if (cNvPicPr is null)
                return null;

            var picLocks = cNvPicPr.Element(ANs + "picLocks");
            if (picLocks is null)
            {
                picLocks = new XElement(ANs + "picLocks");
                cNvPicPr.Add(picLocks);
            }

            var lockObj = new PictureFrameLock();
            lockObj.InitInternal(picLocks, _slidePart);
            return lockObj;
        }
    }

    /// <inheritdoc/>
    public IPictureFillFormat? PictureFormat
    {
        get
        {
            if (_element is null)
                return null;

            var blipFill = _element.Element(PNs + "blipFill");
            if (blipFill is null)
                return null;

            var format = new PictureFillFormat();
            format.InitInternal(blipFill, _slidePart, _parentSlide);
            return format;
        }
    }

    /// <inheritdoc/>
    public float RelativeScaleHeight
    {
        get => _relativeScaleHeight;
        set => _relativeScaleHeight = value;
    }

    /// <inheritdoc/>
    public float RelativeScaleWidth
    {
        get => _relativeScaleWidth;
        set => _relativeScaleWidth = value;
    }

    /// <summary>
    /// Gets a value indicating whether this picture frame is a cameo.
    /// </summary>
    public bool IsCameo => false;
}
