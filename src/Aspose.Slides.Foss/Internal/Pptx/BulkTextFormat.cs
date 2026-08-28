using System.Xml.Linq;

namespace Aspose.Slides.Foss.Internal.Pptx;

/// <summary>
/// Helper for applying bulk text formatting to collections of cells.
/// Implements the logic behind the <see cref="IBulkTextFormattable"/> overloads of
/// <c>SetTextFormat</c>
/// for Table, Row, and Column classes.
/// </summary>
internal static class BulkTextFormat
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static readonly XName ATxBody = ANs + "txBody";
    private static readonly XName AP = ANs + "p";
    private static readonly XName AR = ANs + "r";
    private static readonly XName ARPr = ANs + "rPr";
    private static readonly XName AEndParaRPr = ANs + "endParaRPr";
    private static readonly XName APPr = ANs + "pPr";
    private static readonly XName ABodyPr = ANs + "bodyPr";
    private static readonly XName ATcPr = ANs + "tcPr";

    /// <summary>
    /// Copies all attributes from <paramref name="src"/> to <paramref name="dst"/>,
    /// overwriting existing ones.
    /// </summary>
    internal static void CopyXmlAttrs(XElement src, XElement dst)
    {
        foreach (var attr in src.Attributes())
        {
            dst.SetAttributeValue(attr.Name, attr.Value);
        }
    }

    /// <summary>
    /// Replaces the first child element with the same name in <paramref name="parent"/>,
    /// or appends a deep copy if no match exists.
    /// </summary>
    internal static void ReplaceOrAddChild(XElement parent, XElement srcChild)
    {
        var clone = new XElement(srcChild);
        var existing = parent.Element(srcChild.Name);
        if (existing is not null)
        {
            existing.ReplaceWith(clone);
        }
        else
        {
            parent.Add(clone);
        }
    }

    /// <summary>
    /// Copies attributes and child elements from a source rPr to a target rPr-like element.
    /// </summary>
    internal static void ApplyRprToElement(XElement srcRpr, XElement target)
    {
        CopyXmlAttrs(srcRpr, target);
        foreach (var child in srcRpr.Elements())
        {
            ReplaceOrAddChild(target, child);
        }
    }

    /// <summary>
    /// Applies a <see cref="BasePortionFormat"/> to every run and endParaRPr in every cell.
    /// </summary>
    internal static void ApplyPortionFormat(IEnumerable<Cell> cells, BasePortionFormat source)
    {
        var srcRpr = source.RprElement;
        if (srcRpr is null)
            return;

        foreach (var cell in cells)
        {
            var txBody = cell.TcElement?.Element(ATxBody);
            if (txBody is null)
                continue;

            foreach (var p in txBody.Elements(AP))
            {
                // Apply to runs
                foreach (var r in p.Elements(AR))
                {
                    var rpr = r.Element(ARPr);
                    if (rpr is null)
                    {
                        rpr = new XElement(ARPr);
                        r.AddFirst(rpr);
                    }
                    ApplyRprToElement(srcRpr, rpr);
                }

                // Apply to endParaRPr
                var endRpr = p.Element(AEndParaRPr);
                if (endRpr is not null)
                {
                    ApplyRprToElement(srcRpr, endRpr);
                }
            }
        }
    }

    /// <summary>
    /// Applies a <see cref="ParagraphFormat"/> to every paragraph in every cell.
    /// </summary>
    internal static void ApplyParagraphFormat(IEnumerable<Cell> cells, ParagraphFormat source)
    {
        var srcPpr = source.PprElement;
        if (srcPpr is null)
            return;

        foreach (var cell in cells)
        {
            var txBody = cell.TcElement?.Element(ATxBody);
            if (txBody is null)
                continue;

            foreach (var p in txBody.Elements(AP))
            {
                var ppr = p.Element(APPr);
                if (ppr is null)
                {
                    ppr = new XElement(APPr);
                    p.AddFirst(ppr);
                }
                CopyXmlAttrs(srcPpr, ppr);
                foreach (var child in srcPpr.Elements())
                {
                    ReplaceOrAddChild(ppr, child);
                }
            }
        }
    }

    /// <summary>
    /// Applies a <see cref="TextFrameFormat"/> to every text body in every cell.
    /// Also propagates the <c>vert</c> attribute to <c>&lt;a:tcPr&gt;</c> to match
    /// Aspose.Slides behaviour where vertical text type is mirrored on the cell properties element.
    /// </summary>
    internal static void ApplyTextFrameFormat(IEnumerable<Cell> cells, TextFrameFormat source)
    {
        var srcBodyPr = source.TxBodyElement?.Element(ABodyPr);
        if (srcBodyPr is null)
            return;

        var vertVal = (string?)srcBodyPr.Attribute("vert");

        foreach (var cell in cells)
        {
            var txBody = cell.TcElement?.Element(ATxBody);
            if (txBody is null)
                continue;

            var bodyPr = txBody.Element(ABodyPr);
            if (bodyPr is null)
            {
                bodyPr = new XElement(ABodyPr);
                txBody.AddFirst(bodyPr);
            }

            CopyXmlAttrs(srcBodyPr, bodyPr);
            foreach (var child in srcBodyPr.Elements())
            {
                ReplaceOrAddChild(bodyPr, child);
            }

            // Mirror vert on <a:tcPr>
            if (vertVal is not null)
            {
                var tcPr = cell.TcElement?.Element(ATcPr);
                if (tcPr is not null)
                {
                    tcPr.SetAttributeValue("vert", vertVal);
                }
            }
        }
    }

    /// <summary>
    /// Dispatches to the correct applier based on the runtime type of <paramref name="source"/>,
    /// then saves the slide part.
    /// </summary>
    /// <param name="cells">Cells to format.</param>
    /// <param name="source">A <see cref="BasePortionFormat"/>, <see cref="ParagraphFormat"/>,
    /// or <see cref="TextFrameFormat"/> instance.</param>
    /// <param name="slidePart">The slide part to save afterwards.</param>
    /// <exception cref="ArgumentException">
    /// Thrown when <paramref name="source"/> is not a recognized format type.
    /// </exception>
    internal static void ApplyTextFormat(IEnumerable<Cell> cells, object source, SlidePart? slidePart)
    {
        switch (source)
        {
            case BasePortionFormat portionFormat:
                ApplyPortionFormat(cells, portionFormat);
                break;
            case ParagraphFormat paragraphFormat:
                ApplyParagraphFormat(cells, paragraphFormat);
                break;
            case TextFrameFormat textFrameFormat:
                ApplyTextFrameFormat(cells, textFrameFormat);
                break;
            default:
                throw new ArgumentException(
                    $"SetTextFormat expects PortionFormat, ParagraphFormat, or TextFrameFormat, got {source.GetType().Name}",
                    nameof(source));
        }

        slidePart?.Save();
    }
}
