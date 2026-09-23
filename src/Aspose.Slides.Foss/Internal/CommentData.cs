using System.Xml.Linq;

namespace Aspose.Slides.Foss.Internal;

/// <summary>
/// Raw data for a comment parsed from XML.
/// </summary>
internal sealed class CommentData
{
    /// <summary>
    /// PowerPoint's unit for a comment's <c>p:pos</c>, per centimetre.
    /// </summary>
    /// <remarks>
    /// The schema types <c>p:pos</c> as <c>a:CT_Point2D</c>, in EMU, but PowerPoint writes and reads
    /// it in eighths of a point, 576 to the inch: a comment it places 10 pt from the corner of the
    /// slide is written <c>x="80" y="80"</c>. Written in EMU, a position is 1587.5 times too large
    /// and lands far outside the slide.
    /// </remarks>
    internal const double PositionUnitsPerCm = 576 / 2.54;

    /// <summary>
    /// The <c>p:ext/@uri</c> of the extension that carries <c>p15:threadingInfo</c>.
    /// </summary>
    /// <remarks>
    /// <c>CT_Comment</c> has no attribute for a parent. PowerPoint ties a reply to the comment it
    /// answers with this extension: the reply's <c>p15:threadingInfo</c> holds a
    /// <c>p15:parentCm</c> naming the parent by author and index, and the parent's is empty. No
    /// other part is needed for PowerPoint to show the reply as a reply.
    /// </remarks>
    internal const string ThreadingExtensionUri = "{C676402C-5697-4E1C-873F-D02D1690AC5C}";

    internal readonly XElement Elem;

    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";

    private static readonly XNamespace P15Ns = "http://schemas.microsoft.com/office/powerpoint/2012/main";

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
    /// The author id and <c>idx</c> of the comment this one replies to, or <c>null</c> for a
    /// top-level comment. Read from and written to <c>p15:threadingInfo/p15:parentCm</c>.
    /// </summary>
    /// <remarks>
    /// Clearing it removes <c>p15:parentCm</c> and keeps the empty <c>p15:threadingInfo</c>, which is
    /// what PowerPoint writes on a top-level comment.
    /// </remarks>
    internal (int AuthorId, int Idx)? ParentRef
    {
        get
        {
            var parentCm = ThreadingInfo()?.Element(P15Ns + "parentCm");
            if (parentCm is null ||
                !int.TryParse(parentCm.Attribute("authorId")?.Value, out var authorId) ||
                !int.TryParse(parentCm.Attribute("idx")?.Value, out var idx))
            {
                return null;
            }
            return (authorId, idx);
        }
        set
        {
            if (value is null)
            {
                ThreadingInfo()?.Elements(P15Ns + "parentCm").Remove();
                return;
            }

            var info = EnsureThreadingInfo();
            info.Elements(P15Ns + "parentCm").Remove();
            info.Add(new XElement(P15Ns + "parentCm",
                new XAttribute("authorId", value.Value.AuthorId),
                new XAttribute("idx", value.Value.Idx)));
        }
    }

    /// <summary>
    /// Returns this comment's <c>p15:threadingInfo</c>, adding an empty one if it has none.
    /// </summary>
    /// <remarks>
    /// PowerPoint writes <c>dt</c> in local time and records the offset in <c>timeZoneBias</c>, the
    /// minutes to add to reach UTC. This library writes <c>dt</c> as the caller's
    /// <see cref="DateTime"/> reads, so the bias is this machine's offset at that time — what
    /// PowerPoint itself writes for a comment made here.
    /// </remarks>
    internal XElement EnsureThreadingInfo()
    {
        var info = ThreadingInfo();
        if (info is not null)
            return info;

        var extLst = Elem.Element(PNs + "extLst");
        if (extLst is null)
        {
            // p:extLst is the last child of CT_Comment, so appending keeps the schema order.
            extLst = new XElement(PNs + "extLst");
            Elem.Add(extLst);
        }

        var created = DateTimeHelpers.StrToDt(DtStr) ?? DateTime.Now;
        info = new XElement(P15Ns + "threadingInfo",
            new XAttribute(XNamespace.Xmlns + "p15", P15Ns.NamespaceName),
            new XAttribute("timeZoneBias", -(int)TimeZoneInfo.Local.GetUtcOffset(created).TotalMinutes));
        extLst.Add(new XElement(PNs + "ext", new XAttribute("uri", ThreadingExtensionUri), info));
        return info;
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

    /// <summary>The X position, in centimetres from the left edge of the slide.</summary>
    internal float PosX
    {
        get
        {
            var pos = Elem.Element(PNs + "pos");
            if (pos is not null)
                return FromPositionUnits(int.Parse(pos.Attribute("x")?.Value ?? "0"));
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
            pos.SetAttributeValue("x", ToPositionUnits(value));
        }
    }

    /// <summary>The Y position, in centimetres from the top edge of the slide.</summary>
    internal float PosY
    {
        get
        {
            var pos = Elem.Element(PNs + "pos");
            if (pos is not null)
                return FromPositionUnits(int.Parse(pos.Attribute("y")?.Value ?? "0"));
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
            pos.SetAttributeValue("y", ToPositionUnits(value));
        }
    }

    /// <summary>
    /// Converts a coordinate in centimetres to the value written in <c>p:pos</c>; see
    /// <see cref="PositionUnitsPerCm"/>.
    /// </summary>
    internal static int ToPositionUnits(float cm) => (int)Math.Round(cm * PositionUnitsPerCm);

    /// <summary>
    /// Converts a <c>p:pos</c> coordinate to centimetres; see <see cref="PositionUnitsPerCm"/>.
    /// </summary>
    internal static float FromPositionUnits(int units) => (float)(units / PositionUnitsPerCm);

    private XElement? ThreadingInfo() =>
        Elem.Elements(PNs + "extLst").Elements(PNs + "ext")
            .Where(ext => (string?)ext.Attribute("uri") == ThreadingExtensionUri)
            .Elements(P15Ns + "threadingInfo")
            .FirstOrDefault();
}
