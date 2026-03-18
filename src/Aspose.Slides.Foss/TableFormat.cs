using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents format of a table.
/// </summary>
public sealed class TableFormat : ITableFormat
{
    private XElement? _tblPr;
    private SlidePart? _slidePart;
    private IBaseSlide? _parentSlide;

    /// <summary>
    /// Initializes internal state from the table properties XML element.
    /// </summary>
    internal TableFormat InitInternal(XElement? tblPr, SlidePart? slidePart, IBaseSlide? parentSlide)
    {
        _tblPr = tblPr;
        _slidePart = slidePart;
        _parentSlide = parentSlide;
        return this;
    }

    /// <inheritdoc />
    public IFillFormat FillFormat
    {
        get
        {
            var ff = new FillFormat();
            if (_tblPr is not null)
                ff.InitInternal(_tblPr, _parentSlide, _slidePart);
            return ff;
        }
    }
}
