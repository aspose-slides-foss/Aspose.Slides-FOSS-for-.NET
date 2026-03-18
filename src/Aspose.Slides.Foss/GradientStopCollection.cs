using System.Collections;
using System.Xml.Linq;
using Aspose.Slides.Foss.Drawing;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Manages a collection of <c>&lt;a:gs&gt;</c> child elements within an <c>&lt;a:gsLst&gt;</c> XML element.
/// Each mutation persists changes by calling <see cref="Save"/> on the owning <see cref="SlidePart"/>.
/// </summary>
public sealed class GradientStopCollection : PVIObject, IGradientStopCollection, IEnumerable<IGradientStop>
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XName GsName = ANs + "gs";

    private XElement? _gsLst;
    private SlidePart? _slidePart;

    /// <summary>
    /// Initializes internal state. Called after default construction to inject dependencies.
    /// </summary>
    /// <param name="gsLstElement">The <c>&lt;a:gsLst&gt;</c> XML element.</param>
    /// <param name="slidePart">The owning slide part for persistence.</param>
    /// <param name="parentSlide">The parent slide reference.</param>
    internal void InitInternal(XElement gsLstElement, SlidePart? slidePart, IBaseSlide? parentSlide)
    {
        _gsLst = gsLstElement;
        _slidePart = slidePart;
        _parentSlide = parentSlide;
    }

    /// <inheritdoc/>
    public IList<IGradientStop> AsICollection
    {
        get
        {
            var list = new List<IGradientStop>();
            for (int i = 0; i < Count; i++)
                list.Add(this[i]);
            return list;
        }
    }

    /// <inheritdoc/>
    public IEnumerable<IGradientStop> AsIEnumerable => AsICollection;

    /// <inheritdoc/>
    public int Count
    {
        get
        {
            if (_gsLst is null)
                return 0;
            return _gsLst.Elements(GsName).Count();
        }
    }

    /// <inheritdoc/>
    public IGradientStop this[int index]
    {
        get
        {
            var elements = _gsLst!.Elements(GsName).ToList();
            if (index < 0 || index >= elements.Count)
                throw new IndexOutOfRangeException($"Index {index} is out of range. Collection has {elements.Count} elements.");

            var stop = new GradientStop();
            stop.InitInternal(elements[index], _slidePart!, _parentSlide);
            return stop;
        }
    }

    /// <inheritdoc/>
    public IGradientStop Add(float position, Color color) => AddCore(position, color);

    /// <inheritdoc/>
    public IGradientStop Add(float position, PresetColor presetColor) => AddCore(position, presetColor);

    /// <inheritdoc/>
    public IGradientStop Add(float position, SchemeColor schemeColor) => AddCore(position, schemeColor);

    /// <inheritdoc/>
    public void Insert(int index, float position, Color color) => InsertCore(index, position, color);

    /// <inheritdoc/>
    public void Insert(int index, float position, PresetColor presetColor) => InsertCore(index, position, presetColor);

    /// <inheritdoc/>
    public void Insert(int index, float position, SchemeColor schemeColor) => InsertCore(index, position, schemeColor);

    /// <inheritdoc/>
    public void RemoveAt(int index)
    {
        var elements = _gsLst!.Elements(GsName).ToList();
        if (index >= 0 && index < elements.Count)
        {
            elements[index].Remove();
            Save();
        }
    }

    /// <inheritdoc/>
    public void Clear()
    {
        _gsLst!.Elements(GsName).Remove();
        Save();
    }

    /// <inheritdoc/>
    public IEnumerator<IGradientStop> GetEnumerator()
    {
        for (int i = 0; i < Count; i++)
            yield return this[i];
    }

    /// <inheritdoc/>
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    private static string ConvertPosition(float position) =>
        ((int)Math.Round(position * 100000)).ToString();

    private IGradientStop AddCore(float position, object colorArg)
    {
        var posStr = ConvertPosition(position);
        var gsElement = new XElement(GsName, new XAttribute("pos", posStr));
        _gsLst!.Add(gsElement);

        var cf = new ColorFormat();
        cf.InitInternal(gsElement, _parentSlide, _slidePart);
        SetColorFromArg(cf, colorArg);

        Save();

        var stop = new GradientStop();
        stop.InitInternal(gsElement, _slidePart!, _parentSlide);
        return stop;
    }

    private void InsertCore(int index, float position, object colorArg)
    {
        var posStr = ConvertPosition(position);
        var gsElement = new XElement(GsName, new XAttribute("pos", posStr));

        var cf = new ColorFormat();
        cf.InitInternal(gsElement, _parentSlide, _slidePart);
        SetColorFromArg(cf, colorArg);

        var existing = _gsLst!.Elements(GsName).ToList();
        if (index >= existing.Count)
        {
            _gsLst.Add(gsElement);
        }
        else
        {
            existing[index].AddBeforeSelf(gsElement);
        }

        Save();
    }

    private static void SetColorFromArg(ColorFormat cf, object colorArg)
    {
        switch (colorArg)
        {
            case Color color:
                cf.Color = color;
                break;
            case PresetColor presetColor:
                cf.PresetColor = presetColor;
                break;
            case SchemeColor schemeColor:
                cf.SchemeColor = schemeColor;
                break;
        }
    }

    private void Save()
    {
        _slidePart?.Save();
    }
}
