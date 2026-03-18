using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a layout slide in a presentation.
/// </summary>
public sealed class LayoutSlide : BaseSlide, ILayoutSlide
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";

    private static readonly Dictionary<string, SlideLayoutType> OoxmlLayoutTypeMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["blank"] = SlideLayoutType.Blank,
        ["chart"] = SlideLayoutType.Chart,
        ["chartAndTx"] = SlideLayoutType.ChartAndText,
        ["clipArtAndTx"] = SlideLayoutType.ClipArtAndText,
        ["clipArtAndVertTx"] = SlideLayoutType.ClipArtAndVerticalText,
        ["cust"] = SlideLayoutType.Custom,
        ["dgm"] = SlideLayoutType.Diagram,
        ["fourObj"] = SlideLayoutType.FourObjects,
        ["mediaAndTx"] = SlideLayoutType.MediaClipAndText,
        ["obj"] = SlideLayoutType.BigObject,
        ["objAndTwoObj"] = SlideLayoutType.ObjectAndTwoObject,
        ["objAndTx"] = SlideLayoutType.ObjectAndText,
        ["objOnly"] = SlideLayoutType.BigObject,
        ["objOverTx"] = SlideLayoutType.ObjectOverText,
        ["objTx"] = SlideLayoutType.ObjectAndTextOverChart,
        ["picTx"] = SlideLayoutType.PictureAndCaption,
        ["secHead"] = SlideLayoutType.SectionHeader,
        ["tbl"] = SlideLayoutType.Table,
        ["title"] = SlideLayoutType.Title,
        ["titleOnly"] = SlideLayoutType.TitleOnly,
        ["tx"] = SlideLayoutType.Text,
        ["txAndChart"] = SlideLayoutType.TextAndChart,
        ["txAndClipArt"] = SlideLayoutType.TextAndClipArt,
        ["txAndMedia"] = SlideLayoutType.TextAndMediaClip,
        ["txAndObj"] = SlideLayoutType.TextAndObject,
        ["txAndTwoObj"] = SlideLayoutType.TextAndTwoObjects,
        ["txOverObj"] = SlideLayoutType.TextOverObject,
        ["twoColTx"] = SlideLayoutType.TwoColumnText,
        ["twoObj"] = SlideLayoutType.TwoObjects,
        ["twoObjAndObj"] = SlideLayoutType.TwoObjectsAndObject,
        ["twoObjAndTx"] = SlideLayoutType.TwoObjectsAndText,
        ["twoObjOverTx"] = SlideLayoutType.TwoObjectsOverText,
        ["twoTxTwoObj"] = SlideLayoutType.TwoObjects2,
        ["vertTitleAndTx"] = SlideLayoutType.VerticalTitleAndText,
        ["vertTitleAndTxOverChart"] = SlideLayoutType.VerticalTitleAndTextOverChart,
        ["vertTx"] = SlideLayoutType.VerticalText,
    };

    private OpcPackage? _package;
    private Func<string, IMasterSlide?>? _masterResolver;
    private IMasterSlide? _masterSlideCache;
    private SlideLayoutType? _layoutTypeCache;

    /// <summary>
    /// Initializes internal state. Called after default construction to inject dependencies.
    /// </summary>
    /// <param name="presentation">The parent Presentation object.</param>
    /// <param name="package">The OPC package.</param>
    /// <param name="partName">The part name of this layout slide.</param>
    /// <param name="layoutPart">The parsed SlidePart for this layout.</param>
    /// <param name="masterResolver">Callable that resolves a master part name to a MasterSlide.</param>
    internal void InitInternal(
        IPresentation presentation,
        OpcPackage package,
        string partName,
        SlidePart layoutPart,
        Func<string, IMasterSlide?>? masterResolver = null)
    {
        _presentationRef = presentation;
        _package = package;
        _partName = partName;
        _layoutPart = layoutPart;
        _masterResolver = masterResolver;
    }

    /// <inheritdoc />
    public IMasterSlide? MasterSlide
    {
        get
        {
            if (_masterResolver is not null && _masterSlideCache is null)
            {
                var masterPartName = ResolveMasterPartName();
                if (masterPartName is not null)
                    _masterSlideCache = _masterResolver(masterPartName);
            }

            return _masterSlideCache;
        }
        set => _masterSlideCache = value;
    }

    /// <inheritdoc />
    public SlideLayoutType LayoutType
    {
        get
        {
            if (_layoutTypeCache.HasValue)
                return _layoutTypeCache.Value;

            _layoutTypeCache = ResolveLayoutType();
            return _layoutTypeCache.Value;
        }
        internal set => _layoutTypeCache = value;
    }

    /// <summary>
    /// Gets or sets the OPC part name for this layout slide.
    /// </summary>
    internal string InternalPartName
    {
        get => _partName ?? "";
        set => _partName = value;
    }

    /// <summary>
    /// Gets or sets the OPC package reference for this layout slide.
    /// </summary>
    internal OpcPackage? PackageInternal
    {
        get => _package;
        set => _package = value;
    }

    /// <summary>
    /// Resolves the master slide part name from this layout's relationships.
    /// </summary>
    private string? ResolveMasterPartName()
    {
        if (_partName is null || _package is null)
            return null;

        var relsPath = GetRelsPath(_partName);
        var relsData = _package.GetPart(relsPath);
        if (relsData is null)
            return null;

        var rels = new RelsManager();
        rels.Load(relsData);

        const string slideMasterRelType =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";

        var masterRels = rels.FindByType(slideMasterRelType);
        foreach (var rel in masterRels)
        {
            return ResolvePartName(GetDirectoryPath(_partName), rel.Target);
        }

        return null;
    }

    /// <summary>
    /// Resolves the layout type from the layout slide XML.
    /// </summary>
    private SlideLayoutType ResolveLayoutType()
    {
        var part = _layoutPart;
        if (part?.Element is null)
            return SlideLayoutType.Custom;

        var typeAttr = part.Element.Attribute("type");
        if (typeAttr is null)
            return SlideLayoutType.Custom;

        return OoxmlLayoutTypeMap.TryGetValue(typeAttr.Value, out var layoutType)
            ? layoutType
            : SlideLayoutType.Custom;
    }

    private static string ResolvePartName(string basePath, string target)
    {
        var combined = basePath + target;
        var segments = combined.Split('/');
        var resolved = new Stack<string>();

        foreach (var segment in segments)
        {
            if (segment == "..")
            {
                if (resolved.Count > 0)
                    resolved.Pop();
            }
            else if (segment != "." && segment.Length > 0)
            {
                resolved.Push(segment);
            }
        }

        return string.Join("/", resolved.Reverse());
    }

    private static string GetDirectoryPath(string partName)
    {
        var lastSlash = partName.LastIndexOf('/');
        return lastSlash >= 0 ? partName[..(lastSlash + 1)] : "";
    }

    private static string GetRelsPath(string partName)
    {
        var dir = GetDirectoryPath(partName);
        var file = partName[dir.Length..];
        return $"{dir}_rels/{file}.rels";
    }
}
