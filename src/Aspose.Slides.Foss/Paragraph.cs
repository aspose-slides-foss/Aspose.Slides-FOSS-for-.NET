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

    /// <summary>
    /// Returns the <c>&lt;a:p&gt;</c> element to write for this paragraph, carrying a run for every
    /// portion it was given and, failing that, for the text it was given directly.
    /// </summary>
    /// <remarks>
    /// A paragraph built with <c>new Paragraph()</c> is detached: its portions and its text live
    /// beside an element nothing has written yet. This is the point at which they become markup.
    /// </remarks>
    internal XElement ElementForAttaching()
    {
        var element = _pElement ??= new XElement(ANs + "p");

        _portions.WriteInto(element);

        if (_text.Length > 0 && !element.Elements(ANs + "r").Any())
            element.Add(new XElement(ANs + "r", new XElement(ANs + "t", _text)));

        return element;
    }

    /// <summary>
    /// Binds this paragraph to the element it was written as, so that later changes to it reach the
    /// document rather than a copy of it.
    /// </summary>
    internal void BindTo(XElement pElement, XElement txBodyElement, SlidePart? slidePart, IBaseSlide? parentSlide) =>
        InitInternal(pElement, txBodyElement, slidePart, parentSlide);

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
