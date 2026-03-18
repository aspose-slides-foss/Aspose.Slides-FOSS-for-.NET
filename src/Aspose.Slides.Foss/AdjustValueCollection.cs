using System.Collections;
using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a collection of shape's adjustment values.
/// </summary>
public sealed class AdjustValueCollection : IAdjustValueCollection, IEnumerable<IAdjustValue>
{
    private static readonly XName GdName =
        XName.Get("gd", "http://schemas.openxmlformats.org/drawingml/2006/main");

    private XElement? _avLst;
    private SlidePart? _slidePart;

    /// <summary>
    /// Initializes internal state. Called after default construction to inject dependencies.
    /// </summary>
    /// <param name="avLstElement">The <c>a:avLst</c> XML element, or <c>null</c>.</param>
    /// <param name="slidePart">The owning slide part.</param>
    /// <returns>This instance for fluent chaining.</returns>
    internal AdjustValueCollection InitInternal(XElement? avLstElement, SlidePart? slidePart)
    {
        _avLst = avLstElement;
        _slidePart = slidePart;
        return this;
    }

    /// <summary>
    /// Gets the guide definition (<c>a:gd</c>) child elements from the adjustment list.
    /// </summary>
    /// <returns>A list of <c>a:gd</c> elements, or an empty list if no adjustment list is set.</returns>
    public List<XElement> GetGdElements()
    {
        if (_avLst is null)
            return [];

        return [.. _avLst.Elements(GdName)];
    }

    /// <inheritdoc />
    public int Count => GetGdElements().Count;

    /// <inheritdoc />
    public IAdjustValue this[int index]
    {
        get
        {
            var gdElements = GetGdElements();
            if (index < 0 || index >= gdElements.Count)
                throw new IndexOutOfRangeException($"Index {index} is out of range");

            var gd = gdElements[index];
            var av = new AdjustValue();
            av.InitInternal(gd, _slidePart);
            return av;
        }
    }

    /// <inheritdoc />
    public IList<IAdjustValue> AsICollection
    {
        get
        {
            var result = new List<IAdjustValue>();
            foreach (var gd in GetGdElements())
            {
                var av = new AdjustValue();
                av.InitInternal(gd, _slidePart);
                result.Add(av);
            }
            return result;
        }
    }

    /// <inheritdoc />
    public IEnumerable<IAdjustValue> AsIEnumerable => AsICollection;

    /// <inheritdoc />
    public IEnumerator<IAdjustValue> GetEnumerator()
    {
        foreach (var gd in GetGdElements())
        {
            var av = new AdjustValue();
            av.InitInternal(gd, _slidePart);
            yield return av;
        }
    }

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
