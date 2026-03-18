using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents the formatting properties of a table cell, providing access to fill formatting
/// and six border line formats (left, top, right, bottom, diagonal-down, diagonal-up).
/// </summary>
public sealed class CellFormat : PVIObject, ICellFormat
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private XElement? _tcPrElement;
    private SlidePart? _slidePart;

    /// <summary>
    /// Initializes internal state. Called after default construction to inject dependencies.
    /// </summary>
    /// <param name="tcPrElement">The XML element representing table cell properties (<c>&lt;a:tcPr&gt;</c>).</param>
    /// <param name="slidePart">The slide part containing this cell.</param>
    /// <param name="parentSlide">The slide that owns this cell.</param>
    /// <returns>This instance for chaining.</returns>
    internal CellFormat InitInternal(XElement? tcPrElement, SlidePart? slidePart, IBaseSlide? parentSlide)
    {
        _tcPrElement = tcPrElement;
        _slidePart = slidePart;
        _parentSlide = parentSlide;
        return this;
    }

    /// <inheritdoc/>
    public IFillFormat FillFormat
    {
        get
        {
            var ff = new FillFormat();
            if (_tcPrElement is not null)
                ff.InitInternal(_tcPrElement, _parentSlide);
            return ff;
        }
    }

    /// <inheritdoc/>
    public ILineFormat BorderLeft => GetBorder("lnL");

    /// <inheritdoc/>
    public ILineFormat BorderTop => GetBorder("lnT");

    /// <inheritdoc/>
    public ILineFormat BorderRight => GetBorder("lnR");

    /// <inheritdoc/>
    public ILineFormat BorderBottom => GetBorder("lnB");

    /// <inheritdoc/>
    public ILineFormat BorderDiagonalDown => GetBorder("lnTlToBr");

    /// <inheritdoc/>
    public ILineFormat BorderDiagonalUp => GetBorder("lnBlToTr");

    private ILineFormat GetBorder(string tag)
    {
        var lf = new LineFormat();
        if (_tcPrElement is not null)
            lf.InitInternal(_tcPrElement, _parentSlide, ANs + tag);
        return lf;
    }
}
