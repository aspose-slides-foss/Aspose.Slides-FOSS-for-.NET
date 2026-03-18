using System.Xml.Linq;

namespace Aspose.Slides.Foss.Internal.Pptx.CommentsPart;

/// <summary>
/// Raw data for a comment parsed from XML.
/// Wraps the underlying <c>&lt;p:cm&gt;</c> element and provides typed access to its attributes.
/// </summary>
public sealed class CommentData
{
    private const int CmToEmu = 360000;

    private static readonly XNamespace PNs =
        "http://schemas.openxmlformats.org/presentationml/2006/main";

    internal readonly XElement Elem;

    /// <summary>
    /// Initializes a new <see cref="CommentData"/> wrapping the given XML element.
    /// </summary>
    public CommentData(XElement elem)
    {
        Elem = elem;
    }

    /// <summary>
    /// Gets the author ID for this comment.
    /// </summary>
    public int AuthorId =>
        int.Parse(Elem.Attribute("authorId")?.Value ?? "0");

    /// <summary>
    /// Gets the comment index.
    /// </summary>
    public int Idx =>
        int.Parse(Elem.Attribute("idx")?.Value ?? "0");

    /// <summary>
    /// Gets or sets the datetime string in OOXML format.
    /// </summary>
    public string DtStr
    {
        get => Elem.Attribute("dt")?.Value ?? "";
        set => Elem.SetAttributeValue("dt", value);
    }

    /// <summary>
    /// Gets or sets the parent comment ID, or <c>null</c> if this is a top-level comment.
    /// </summary>
    public int? ParentCmId
    {
        get
        {
            var val = Elem.Attribute("parentCmId")?.Value;
            return val is not null ? int.Parse(val) : null;
        }
        set
        {
            if (value is null)
                Elem.Attribute("parentCmId")?.Remove();
            else
                Elem.SetAttributeValue("parentCmId", value.Value);
        }
    }

    /// <summary>
    /// Gets or sets the comment text content.
    /// </summary>
    public string Text
    {
        get => Elem.Element(PNs + "text")?.Value ?? "";
        set
        {
            var textElem = Elem.Element(PNs + "text");
            if (textElem is null)
            {
                textElem = new XElement(PNs + "text");
                Elem.Add(textElem);
            }
            textElem.Value = value;
        }
    }

    /// <summary>
    /// Gets or sets the X position in centimeters.
    /// </summary>
    public float PosX
    {
        get
        {
            var pos = Elem.Element(PNs + "pos");
            return pos is not null
                ? int.Parse(pos.Attribute("x")?.Value ?? "0") / (float)CmToEmu
                : 0f;
        }
        set
        {
            var pos = Elem.Element(PNs + "pos");
            if (pos is null)
            {
                pos = new XElement(PNs + "pos");
                Elem.Add(pos);
            }
            pos.SetAttributeValue("x", (int)Math.Round(value * CmToEmu));
        }
    }

    /// <summary>
    /// Gets or sets the Y position in centimeters.
    /// </summary>
    public float PosY
    {
        get
        {
            var pos = Elem.Element(PNs + "pos");
            return pos is not null
                ? int.Parse(pos.Attribute("y")?.Value ?? "0") / (float)CmToEmu
                : 0f;
        }
        set
        {
            var pos = Elem.Element(PNs + "pos");
            if (pos is null)
            {
                pos = new XElement(PNs + "pos");
                Elem.Add(pos);
            }
            pos.SetAttributeValue("y", (int)Math.Round(value * CmToEmu));
        }
    }
}
