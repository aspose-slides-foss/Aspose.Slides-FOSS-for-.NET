using System.Xml.Linq;

namespace Aspose.Slides.Foss.Internal;

/// <summary>
/// Raw data for a comment parsed from XML.
/// </summary>
internal sealed class CommentData
{
    private const int CmToEmu = 360000;

    internal readonly XElement Elem;

    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";

    internal CommentData(XElement elem)
    {
        Elem = elem;
    }

    internal int AuthorId =>
        int.Parse(Elem.Attribute("authorId")?.Value ?? "0");

    internal int Idx =>
        int.Parse(Elem.Attribute("idx")?.Value ?? "0");

    internal string DtStr
    {
        get => Elem.Attribute("dt")?.Value ?? "";
        set => Elem.SetAttributeValue("dt", value);
    }

    /// <summary>
    /// The <c>idx</c> of the comment this one replies to, held in memory only.
    /// </summary>
    /// <remarks>
    /// See <see cref="CommentsPart.ParentMarkerAttribute"/>: this is not part of CT_Comment and is
    /// removed when the comments part is written.
    /// </remarks>
    internal int? ParentCmId
    {
        get
        {
            var val = Elem.Attribute(CommentsPart.ParentMarkerAttribute)?.Value;
            return val is not null ? int.Parse(val) : null;
        }
        set
        {
            if (value is null)
                Elem.Attribute(CommentsPart.ParentMarkerAttribute)?.Remove();
            else
                Elem.SetAttributeValue(CommentsPart.ParentMarkerAttribute, value.Value);
        }
    }

    internal string Text
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

    internal float PosX
    {
        get
        {
            var pos = Elem.Element(PNs + "pos");
            if (pos is not null)
                return int.Parse(pos.Attribute("x")?.Value ?? "0") / (float)CmToEmu;
            return 0f;
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

    internal float PosY
    {
        get
        {
            var pos = Elem.Element(PNs + "pos");
            if (pos is not null)
                return int.Parse(pos.Attribute("y")?.Value ?? "0") / (float)CmToEmu;
            return 0f;
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
