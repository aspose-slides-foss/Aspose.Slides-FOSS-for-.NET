using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a portion (run) of text inside a text paragraph.
/// </summary>
public sealed class Portion : IPortion
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XName AR = ANs + "r";
    private static readonly XName AT = ANs + "t";
    private static readonly XName ARPr = ANs + "rPr";

    private XElement _rElement;
    private XElement? _pElement;
    private XElement? _txBodyElement;
    private SlidePart? _slidePart;
    private IBaseSlide? _parentSlide;

    /// <summary>
    /// Initializes a new instance of the <see cref="Portion"/> class with empty text.
    /// </summary>
    public Portion()
    {
        _rElement = new XElement(AR, new XElement(AT, ""));
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Portion"/> class with the specified text.
    /// </summary>
    /// <param name="text">The initial text content.</param>
    public Portion(string text)
    {
        _rElement = new XElement(AR, new XElement(AT, text ?? ""));
    }

    /// <summary>
    /// Replaces the detached element with a live one from a parsed document.
    /// </summary>
    internal Portion InitInternal(XElement rElement, XElement? pElement, XElement? txBodyElement, SlidePart? slidePart, IBaseSlide? parentSlide)
    {
        _rElement = rElement;
        _pElement = pElement;
        _txBodyElement = txBodyElement;
        _slidePart = slidePart;
        _parentSlide = parentSlide;
        return this;
    }

    /// <summary>
    /// Gets the underlying &lt;a:r&gt; XML element.
    /// </summary>
    internal XElement RElement => _rElement;

    /// <inheritdoc/>
    public override IBasePortionFormat? PortionFormat
    {
        get
        {
            var rpr = _rElement.Element(ARPr);
            if (rpr is null)
            {
                rpr = new XElement(ARPr);
                _rElement.AddFirst(rpr);
            }

            var pf = new PortionFormat();
            pf.InitInternal(rpr, _slidePart, _parentSlide);
            return pf;
        }
    }

    /// <inheritdoc/>
    public override string Text
    {
        get
        {
            var tElem = _rElement.Element(AT);
            return tElem?.Value ?? "";
        }
        set
        {
            var tElem = _rElement.Element(AT);
            if (tElem is null)
            {
                tElem = new XElement(AT);
                _rElement.Add(tElem);
            }

            tElem.Value = value ?? "";
            _slidePart?.Save();
        }
    }

    /// <inheritdoc/>
    public override ISlideComponent AsISlideComponent => this;

    /// <inheritdoc/>
    public override IBaseSlide? Slide => _parentSlide;

    /// <inheritdoc/>
    public override IPresentationComponent AsIPresentationComponent => this;

    /// <inheritdoc/>
    public override IPresentation? Presentation => _parentSlide?.Presentation;
}
