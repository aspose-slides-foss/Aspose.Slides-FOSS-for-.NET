using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a paragraph of text.
/// </summary>
public sealed class Paragraph : IParagraph
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private XElement? _pElement;
    private XElement? _txBodyElement;
    private string _text = "";
    private readonly PortionCollection _portions = new();
    private readonly ParagraphFormat _paragraphFormat = new();
    private IBaseSlide? _parentSlide;

    /// <summary>
    /// Initializes a new instance of the <see cref="Paragraph"/> class.
    /// </summary>
    public Paragraph()
    {
        _pElement = new XElement(ANs + "p");
    }

    /// <summary>
    /// Initializes the internal state of this paragraph.
    /// </summary>
    /// <param name="parentSlide">The parent slide that owns this paragraph.</param>
    /// <returns>This instance for fluent chaining.</returns>
    internal Paragraph InitInternal(IBaseSlide? parentSlide)
    {
        _parentSlide = parentSlide;
        _portions.InitInternal(parentSlide);
        return this;
    }

    /// <summary>
    /// Initializes the internal state of this paragraph from an XML element.
    /// </summary>
    /// <param name="pElement">The &lt;a:p&gt; XML element.</param>
    /// <param name="txBodyElement">The parent &lt;a:txBody&gt; XML element.</param>
    /// <param name="slidePart">The slide part for relationship resolution.</param>
    /// <param name="parentSlide">The parent slide that owns this paragraph.</param>
    /// <returns>This instance for fluent chaining.</returns>
    internal Paragraph InitInternal(XElement pElement, XElement? txBodyElement, SlidePart? slidePart, IBaseSlide? parentSlide)
    {
        _pElement = pElement;
        _txBodyElement = txBodyElement;
        _parentSlide = parentSlide;
        _portions.InitInternal(pElement, txBodyElement, slidePart, parentSlide);

        if (slidePart is not null && parentSlide is not null)
        {
            Picture.FlushPendingBlipImages(pElement, slidePart, parentSlide);
        }

        var ppr = pElement.Element(ANs + "pPr");
        if (ppr is null)
        {
            ppr = new XElement(ANs + "pPr");
            pElement.AddFirst(ppr);
        }
        _paragraphFormat.InitInternal(ppr, slidePart, parentSlide);

        return this;
    }

    /// <inheritdoc/>
    public override IPortionCollection Portions => _portions;

    /// <inheritdoc/>
    public override IParagraphFormat ParagraphFormat => _paragraphFormat;

    /// <inheritdoc/>
    public override string Text
    {
        get => _text;
        set => _text = value;
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
