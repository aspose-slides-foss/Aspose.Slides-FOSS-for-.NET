using System.Collections;
using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a collection of paragraphs.
/// </summary>
public sealed class ParagraphCollection : ISlideComponent, IParagraphCollection, IEnumerable<IParagraph>
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XName AP = ANs + "p";
    private static readonly XName AR = ANs + "r";

    private readonly List<IParagraph> _paragraphs = [];
    private XElement? _txBodyElement;
    private SlidePart? _slidePart;
    private IBaseSlide? _parentSlide;

    /// <summary>
    /// Initializes the internal state of this collection.
    /// </summary>
    /// <param name="parentSlide">The parent slide that owns this collection.</param>
    /// <returns>This instance for fluent chaining.</returns>
    internal ParagraphCollection InitInternal(IBaseSlide? parentSlide)
    {
        _parentSlide = parentSlide;
        return this;
    }

    /// <summary>
    /// Initializes the internal state of this collection from an XML text body element.
    /// </summary>
    /// <param name="txBodyElement">The &lt;a:txBody&gt; XML element containing paragraphs.</param>
    /// <param name="slidePart">The slide part for persistence.</param>
    /// <param name="parentSlide">The parent slide that owns this collection.</param>
    /// <returns>This instance for fluent chaining.</returns>
    internal ParagraphCollection InitInternal(XElement? txBodyElement, SlidePart? slidePart, IBaseSlide? parentSlide)
    {
        _txBodyElement = txBodyElement;
        _slidePart = slidePart;
        _parentSlide = parentSlide;
        return this;
    }

    /// <summary>
    /// Gets the paragraphs from the underlying XML text body element.
    /// Returns an empty list if no text body element is set.
    /// </summary>
    /// <returns>A list of paragraphs read from the XML.</returns>
    public List<IParagraph> GetParagraphs()
    {
        if (_txBodyElement is null)
            return [];

        var paragraphs = new List<IParagraph>();
        foreach (var pElem in _txBodyElement.Elements(AP))
        {
            var para = new Paragraph();
            para.InitInternal(pElem, _txBodyElement, _slidePart, _parentSlide);

            foreach (var rElem in pElem.Elements(AR))
            {
                var portion = new Portion();
                portion.InitInternal(rElem, pElem, _txBodyElement, _slidePart, _parentSlide);
                para.Portions.Add(portion);
            }

            // Sync paragraph text from portions
            var portions = para.Portions;
            if (portions.Count > 0)
            {
                var texts = new string[portions.Count];
                for (var i = 0; i < portions.Count; i++)
                    texts[i] = portions[i].Text;
                para.Text = string.Join("", texts);
            }

            paragraphs.Add(para);
        }

        return paragraphs;
    }

    /// <inheritdoc/>
    public override IBaseSlide? Slide => _parentSlide;

    /// <inheritdoc/>
    public override IPresentation? Presentation => _parentSlide?.Presentation;

    /// <inheritdoc/>
    public override IPresentationComponent AsIPresentationComponent => this;

    /// <inheritdoc/>
    public ISlideComponent AsISlideComponent => this;

    /// <inheritdoc/>
    public IEnumerable<IParagraph> AsIEnumerable => Enumerate();

    /// <inheritdoc/>
    public IParagraph this[int index] =>
        _txBodyElement is not null ? GetParagraphs()[index] : _paragraphs[index];

    /// <inheritdoc/>
    public int Count =>
        _txBodyElement is not null ? GetParagraphs().Count : _paragraphs.Count;

    /// <inheritdoc/>
    public void Add(IParagraph value) => _paragraphs.Add(value);

    /// <inheritdoc/>
    public void Insert(int index, IParagraph value) => _paragraphs.Insert(index, value);

    /// <inheritdoc/>
    public void Clear() => _paragraphs.Clear();

    /// <inheritdoc/>
    public void RemoveAt(int index) => _paragraphs.RemoveAt(index);

    /// <inheritdoc/>
    public bool Remove(IParagraph item) => _paragraphs.Remove(item);

    /// <inheritdoc/>
    public bool Contains(IParagraph item) => _paragraphs.Contains(item);

    /// <inheritdoc/>
    public int IndexOf(IParagraph item) => _paragraphs.IndexOf(item);

    /// <inheritdoc/>
    public bool IsReadOnly => false;

    /// <inheritdoc/>
    public IEnumerator<IParagraph> GetEnumerator() =>
        _txBodyElement is not null ? GetParagraphs().GetEnumerator() : _paragraphs.GetEnumerator();

    /// <inheritdoc/>
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    private IEnumerable<IParagraph> Enumerate()
    {
        foreach (var paragraph in this)
            yield return paragraph;
    }
}
