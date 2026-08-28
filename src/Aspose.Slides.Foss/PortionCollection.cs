using System.Collections;
using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a mutable collection of portions belonging to a slide component.
/// </summary>
public sealed class PortionCollection : ISlideComponent, IPortionCollection, IEnumerable<IPortion>
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XName AR = ANs + "r";
    private static readonly XName AT = ANs + "t";

    private readonly List<IPortion> _portions = [];
    private XElement? _pElement;
    private XElement? _txBodyElement;
    private SlidePart? _slidePart;
    private IBaseSlide? _parentSlide;

    /// <summary>
    /// Initializes the internal state of this collection.
    /// </summary>
    /// <param name="parentSlide">The parent slide that owns this collection.</param>
    /// <returns>This instance for fluent chaining.</returns>
    internal PortionCollection InitInternal(IBaseSlide? parentSlide)
    {
        _parentSlide = parentSlide;
        return this;
    }

    /// <summary>
    /// Initializes the internal state of this collection from XML paragraph and text body elements.
    /// </summary>
    /// <param name="pElement">The &lt;a:p&gt; XML element containing text runs.</param>
    /// <param name="txBodyElement">The &lt;a:txBody&gt; XML element.</param>
    /// <param name="slidePart">The slide part for persistence.</param>
    /// <param name="parentSlide">The parent slide that owns this collection.</param>
    /// <returns>This instance for fluent chaining.</returns>
    internal PortionCollection InitInternal(XElement? pElement, XElement? txBodyElement, SlidePart? slidePart, IBaseSlide? parentSlide)
    {
        _pElement = pElement;
        _txBodyElement = txBodyElement;
        _slidePart = slidePart;
        _parentSlide = parentSlide;
        return this;
    }

    /// <summary>
    /// Gets the portions from the underlying XML paragraph element.
    /// Returns an empty list if no paragraph element is set.
    /// </summary>
    /// <returns>A list of portions read from the XML.</returns>
    public List<IPortion> GetPortions()
    {
        if (_pElement is null)
            return [];

        var portions = new List<IPortion>();
        foreach (var rElem in _pElement.Elements(AR))
        {
            var portion = new Portion();
            portion.InitInternal(rElem, _pElement, _txBodyElement, _slidePart, _parentSlide);
            portions.Add(portion);
        }

        return portions;
    }

    /// <summary>
    /// Writes the portions held in this collection into <paramref name="pElement"/> as
    /// <c>&lt;a:r&gt;</c> children, skipping any that are already there.
    /// </summary>
    /// <remarks>
    /// Called when the paragraph that owns this collection joins a text body. Until then a portion
    /// added to a detached paragraph has nowhere to be written to.
    /// </remarks>
    internal void WriteInto(XElement pElement)
    {
        foreach (var portion in _portions)
        {
            if (portion is Portion concrete)
            {
                if (concrete.RElement.Parent is null)
                    pElement.Add(concrete.RElement);
                else if (concrete.RElement.Parent != pElement)
                    pElement.Add(new XElement(concrete.RElement));
            }
            else
            {
                pElement.Add(new XElement(AR, new XElement(AT, portion.Text ?? string.Empty)));
            }
        }
    }

    /// <inheritdoc/>
    public override IBaseSlide? Slide => _parentSlide;

    /// <inheritdoc/>
    public override IPresentation? Presentation => _parentSlide?.Presentation;

    /// <inheritdoc/>
    public override IPresentationComponent AsIPresentationComponent => this;

    /// <inheritdoc/>
    public IEnumerable<IPortion> AsIEnumerable => Enumerate();

    /// <inheritdoc/>
    public IPortion this[int index] =>
        _pElement is not null ? GetPortions()[index] : _portions[index];

    /// <inheritdoc/>
    public int Count =>
        _pElement is not null ? GetPortions().Count : _portions.Count;

    /// <inheritdoc/>
    public void Add(IPortion value)
    {
        _portions.Add(value);

        // If backed by XML, also append the <a:r> element to the paragraph
        // (only if it's not already a child, to avoid duplication).
        if (_pElement is not null && value is Portion concreteValue
            && concreteValue.RElement.Parent != _pElement)
        {
            _pElement.Add(concreteValue.RElement);
            _slidePart?.Save();
        }
    }

    /// <inheritdoc/>
    public int IndexOf(IPortion item) => _portions.IndexOf(item);

    /// <inheritdoc/>
    public void Insert(int index, IPortion value) => _portions.Insert(index, value);

    /// <inheritdoc/>
    public void Clear() => _portions.Clear();

    /// <inheritdoc/>
    public bool Contains(IPortion item) => _portions.Contains(item);

    /// <inheritdoc/>
    public bool IsReadOnly => false;

    /// <inheritdoc/>
    public bool Remove(IPortion item) => _portions.Remove(item);

    /// <inheritdoc/>
    public void RemoveAt(int index) => _portions.RemoveAt(index);

    /// <inheritdoc/>
    public IEnumerator<IPortion> GetEnumerator() =>
        _pElement is not null ? GetPortions().GetEnumerator() : _portions.GetEnumerator();

    /// <inheritdoc/>
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    private IEnumerable<IPortion> Enumerate()
    {
        foreach (var portion in this)
            yield return portion;
    }
}
