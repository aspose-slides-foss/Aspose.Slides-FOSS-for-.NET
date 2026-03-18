using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a single cell within a table in a PowerPoint presentation.
/// </summary>
public sealed class Cell : ISlideComponent, ICell
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const float EmuPerPoint = 12700f;
    private const int DefaultMarginLeftRight = 91440;
    private const int DefaultMarginTopBottom = 45720;

    private static readonly Dictionary<string, TextVerticalType> TextVerticalTypeFromXml = new()
    {
        ["horz"] = TextVerticalType.Horizontal,
        ["vert"] = TextVerticalType.Vertical,
        ["vert270"] = TextVerticalType.Vertical270,
        ["wordArtVert"] = TextVerticalType.WordArtVertical,
        ["eaVert"] = TextVerticalType.EastAsianVertical,
        ["mongolianVert"] = TextVerticalType.MongolianVertical,
        ["wordArtVertRtl"] = TextVerticalType.WordArtVerticalRightToLeft,
    };

    private static readonly Dictionary<TextVerticalType, string> TextVerticalTypeToXml = new()
    {
        [TextVerticalType.Horizontal] = "horz",
        [TextVerticalType.Vertical] = "vert",
        [TextVerticalType.Vertical270] = "vert270",
        [TextVerticalType.WordArtVertical] = "wordArtVert",
        [TextVerticalType.EastAsianVertical] = "eaVert",
        [TextVerticalType.MongolianVertical] = "mongolianVert",
        [TextVerticalType.WordArtVerticalRightToLeft] = "wordArtVertRtl",
    };

    private static readonly Dictionary<string, TextAnchorType> TextAnchorTypeFromXml = new()
    {
        ["t"] = TextAnchorType.Top,
        ["ctr"] = TextAnchorType.Center,
        ["b"] = TextAnchorType.Bottom,
        ["just"] = TextAnchorType.Distributed,
        ["dist"] = TextAnchorType.Distributed,
    };

    private static readonly Dictionary<TextAnchorType, string> TextAnchorTypeToXml = new()
    {
        [TextAnchorType.Top] = "t",
        [TextAnchorType.Center] = "ctr",
        [TextAnchorType.Bottom] = "b",
        [TextAnchorType.Distributed] = "dist",
    };

    private XElement? _tcElement;

    /// <summary>
    /// Gets the backing &lt;a:tc&gt; element for internal use.
    /// </summary>
    internal XElement? TcElement => _tcElement;

    private int _rowIndex;
    private int _colIndex;
    private SlidePart? _slidePart;
    private IBaseSlide? _parentSlide;
    private ITable _table = null!;

    /// <summary>
    /// Initializes internal state. Called after default construction to inject dependencies.
    /// </summary>
    internal Cell InitInternal(XElement? tcElement, int rowIndex, int colIndex, SlidePart? slidePart, IBaseSlide? parentSlide, ITable table)
    {
        _tcElement = tcElement;
        _rowIndex = rowIndex;
        _colIndex = colIndex;
        _slidePart = slidePart;
        _parentSlide = parentSlide;
        _table = table;
        return this;
    }

    /// <inheritdoc />
    public float OffsetX
    {
        get
        {
            float offset = 0;
            var columns = _table.Columns;
            for (int i = 0; i < _colIndex; i++)
                offset += columns[i].Width;
            return offset;
        }
    }

    /// <inheritdoc />
    public float OffsetY
    {
        get
        {
            float offset = 0;
            var rows = _table.Rows;
            for (int i = 0; i < _rowIndex; i++)
                offset += rows[i].Height;
            return offset;
        }
    }

    /// <inheritdoc />
    public int FirstRowIndex => _rowIndex;

    /// <inheritdoc />
    public int FirstColumnIndex => _colIndex;

    /// <inheritdoc />
    public float Width
    {
        get
        {
            var columns = _table.Columns;
            int end = Math.Min(_colIndex + ColSpan, columns.Count);
            float total = 0;
            for (int i = _colIndex; i < end; i++)
                total += columns[i].Width;
            return total;
        }
    }

    /// <inheritdoc />
    public float Height
    {
        get
        {
            var rows = _table.Rows;
            int end = Math.Min(_rowIndex + RowSpan, rows.Count);
            float total = 0;
            for (int i = _rowIndex; i < end; i++)
                total += rows[i].Height;
            return total;
        }
    }

    /// <inheritdoc />
    public float MinimalHeight
    {
        get
        {
            var rows = _table.Rows;
            int end = Math.Min(_rowIndex + RowSpan, rows.Count);
            float total = 0;
            for (int i = _rowIndex; i < end; i++)
                total += rows[i].MinimalHeight;
            return total;
        }
    }

    /// <inheritdoc />
    public float MarginLeft
    {
        get => GetMargin("marL", DefaultMarginLeftRight);
        set => SetMargin("marL", value);
    }

    /// <inheritdoc />
    public float MarginRight
    {
        get => GetMargin("marR", DefaultMarginLeftRight);
        set => SetMargin("marR", value);
    }

    /// <inheritdoc />
    public float MarginTop
    {
        get => GetMargin("marT", DefaultMarginTopBottom);
        set => SetMargin("marT", value);
    }

    /// <inheritdoc />
    public float MarginBottom
    {
        get => GetMargin("marB", DefaultMarginTopBottom);
        set => SetMargin("marB", value);
    }

    /// <inheritdoc />
    public TextVerticalType TextVerticalType
    {
        get
        {
            var tcPr = GetTcPr();
            var vert = tcPr?.Attribute("vert")?.Value;
            if (vert is not null && TextVerticalTypeFromXml.TryGetValue(vert, out var result))
                return result;
            return TextVerticalType.Horizontal;
        }
        set
        {
            var tcPr = EnsureTcPr();
            if (TextVerticalTypeToXml.TryGetValue(value, out var xmlValue))
                tcPr.SetAttributeValue("vert", xmlValue);
            _slidePart?.Save();
        }
    }

    /// <inheritdoc />
    public TextAnchorType TextAnchorType
    {
        get
        {
            var tcPr = GetTcPr();
            var anchor = tcPr?.Attribute("anchor")?.Value;
            if (anchor is not null && TextAnchorTypeFromXml.TryGetValue(anchor, out var result))
                return result;
            return TextAnchorType.Top;
        }
        set
        {
            var tcPr = EnsureTcPr();
            if (TextAnchorTypeToXml.TryGetValue(value, out var xmlValue))
                tcPr.SetAttributeValue("anchor", xmlValue);
            _slidePart?.Save();
        }
    }

    /// <inheritdoc />
    public bool AnchorCenter
    {
        get
        {
            var tcPr = GetTcPr();
            var val = tcPr?.Attribute("anchorCtr")?.Value;
            return val is "1" or "true";
        }
        set
        {
            var tcPr = EnsureTcPr();
            tcPr.SetAttributeValue("anchorCtr", value ? "1" : "0");
            _slidePart?.Save();
        }
    }

    /// <inheritdoc />
    public IRow FirstRow => _table.Rows[_rowIndex];

    /// <inheritdoc />
    public IColumn FirstColumn => _table.Columns[_colIndex];

    /// <inheritdoc />
    public int ColSpan
    {
        get
        {
            var val = _tcElement?.Attribute("gridSpan")?.Value;
            return val is not null && int.TryParse(val, out var span) ? span : 1;
        }
    }

    /// <inheritdoc />
    public int RowSpan
    {
        get
        {
            var val = _tcElement?.Attribute("rowSpan")?.Value;
            return val is not null && int.TryParse(val, out var span) ? span : 1;
        }
    }

    /// <inheritdoc />
    public ITextFrame? TextFrame
    {
        get
        {
            var txBody = _tcElement?.Element(ANs + "txBody");
            if (txBody is null)
                return null;
            var tf = new TextFrame();
            tf.InitInternal(txBody, _slidePart, _parentSlide, parentShape: null);
            return tf;
        }
    }

    /// <inheritdoc />
    public ITable Table => _table;

    /// <inheritdoc />
    public bool IsMergedCell
    {
        get
        {
            if (ColSpan > 1 || RowSpan > 1)
                return true;
            var hMerge = _tcElement?.Attribute("hMerge")?.Value;
            if (hMerge is "1" or "true")
                return true;
            var vMerge = _tcElement?.Attribute("vMerge")?.Value;
            if (vMerge is "1" or "true")
                return true;
            return false;
        }
    }

    /// <inheritdoc />
    public ICellFormat CellFormat
    {
        get
        {
            var tcPr = _tcElement?.Element(ANs + "tcPr");
            var cf = new CellFormat();
            cf.InitInternal(tcPr, _slidePart, _parentSlide);
            return cf;
        }
    }

    /// <inheritdoc />
    public override IBaseSlide? Slide => _parentSlide;

    /// <inheritdoc />
    public override IPresentation? Presentation => _parentSlide?.Presentation;

    /// <inheritdoc />
    public override IPresentationComponent AsIPresentationComponent => this;

    /// <summary>
    /// Returns this instance as an <see cref="ISlideComponent"/>.
    /// </summary>
    public ISlideComponent AsISlideComponent => this;

    private XElement? GetTcPr()
    {
        return _tcElement?.Element(ANs + "tcPr");
    }

    private XElement EnsureTcPr()
    {
        var tcPr = _tcElement!.Element(ANs + "tcPr");
        if (tcPr is null)
        {
            tcPr = new XElement(ANs + "tcPr");
            _tcElement.Add(tcPr);
        }
        return tcPr;
    }

    private float GetMargin(string attributeName, int defaultEmu)
    {
        var tcPr = GetTcPr();
        var val = tcPr?.Attribute(attributeName)?.Value;
        if (val is not null && long.TryParse(val, out var emu))
            return emu / EmuPerPoint;
        return defaultEmu / EmuPerPoint;
    }

    private void SetMargin(string attributeName, float points)
    {
        var tcPr = EnsureTcPr();
        long emu = (long)Math.Round(points * EmuPerPoint);
        tcPr.SetAttributeValue(attributeName, emu.ToString());
        _slidePart?.Save();
    }
}
