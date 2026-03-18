using System.Collections;
using System.Runtime.CompilerServices;
using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents an ordered, mutable collection of <see cref="IShape"/> objects
/// belonging to a slide or group shape. Wraps the <c>&lt;p:spTree&gt;</c> XML element
/// and provides lazy-loaded, cache-backed access to shape wrappers.
/// </summary>
public sealed class ShapeCollection : IShapeCollection
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace RNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private const float EmuPerPoint = 12700f;

    private static readonly HashSet<string> SkippedLocalNames = new(StringComparer.Ordinal)
    {
        "nvGrpSpPr",
        "grpSpPr",
    };

    private SlidePart? _slidePart;
    private IBaseSlide? _parentSlide;
    private List<IShape>? _shapesCache;
    private readonly Dictionary<int, IShape> _elementToShape = new();

    /// <summary>
    /// Initializes a new instance of the <see cref="ShapeCollection"/> class.
    /// </summary>
    public ShapeCollection()
    {
    }

    /// <summary>
    /// Binds this collection to an actual slide. Resets cache and element map.
    /// </summary>
    /// <param name="slidePart">The slide part providing access to the slide XML root.</param>
    /// <param name="parentSlide">The owning slide object.</param>
    internal void InitInternal(SlidePart? slidePart, IBaseSlide? parentSlide)
    {
        _slidePart = slidePart;
        _parentSlide = parentSlide;
        _shapesCache = null;
        _elementToShape.Clear();
    }

    /// <inheritdoc />
    public IGroupShape? ParentGroup { get; internal set; }

    /// <inheritdoc />
    public IList<IShape> AsICollection => LoadShapes();

    /// <inheritdoc />
    public IEnumerable<IShape> AsIEnumerable => LoadShapes();

    /// <inheritdoc />
    public int Count => LoadShapes().Count;

    /// <inheritdoc />
    public IShape this[int index] => LoadShapes()[index];

    /// <inheritdoc />
    public IShape[] ToArray() => LoadShapes().ToArray();

    /// <inheritdoc />
    public IShape[] ToArray(int startIndex, int count)
    {
        return LoadShapes().GetRange(startIndex, count).ToArray();
    }

    /// <inheritdoc />
    public void Reorder(int index, IShape shape)
    {
        var spTree = GetSpTree();
        if (spTree is null)
            return;

        var shapeElement = FindShapeElement(spTree, shape);
        if (shapeElement is null)
            return;

        shapeElement.Remove();
        var shapeElements = GetShapeElements(spTree);
        if (index >= shapeElements.Count)
        {
            spTree.Add(shapeElement);
        }
        else
        {
            shapeElements[index].AddBeforeSelf(shapeElement);
        }

        InvalidateCache();
        Save();
    }

    /// <inheritdoc />
    public void Reorder(int index, IShape[] shapes)
    {
        var spTree = GetSpTree();
        if (spTree is null)
            return;

        var elements = new List<XElement>(shapes.Length);
        foreach (var shape in shapes)
        {
            var el = FindShapeElement(spTree, shape);
            if (el is not null)
            {
                elements.Add(el);
                el.Remove();
            }
        }

        var shapeElements = GetShapeElements(spTree);
        XElement? insertBefore = index < shapeElements.Count ? shapeElements[index] : null;

        foreach (var el in elements)
        {
            if (insertBefore is not null)
                insertBefore.AddBeforeSelf(el);
            else
                spTree.Add(el);
        }

        InvalidateCache();
        Save();
    }

    /// <inheritdoc />
    public IAutoShape AddAutoShape(ShapeType shapeType, float x, float y, float width, float height)
    {
        return AddAutoShape(shapeType, x, y, width, height, true);
    }

    /// <inheritdoc />
    public IAutoShape AddAutoShape(ShapeType shapeType, float x, float y, float width, float height, bool createFromTemplate)
    {
        var spTree = GetSpTree();
        var shapeId = NextShapeId();
        var name = $"AutoShape {shapeId}";
        var element = BuildAutoShapeXml(shapeId, name, shapeType, x, y, width, height, createFromTemplate);

        spTree?.Add(element);

        var shape = new AutoShape();
        shape.InitInternal(element, _slidePart, _parentSlide);
        _elementToShape[RuntimeHelpers.GetHashCode(element)] = shape;

        InvalidateCache();
        Save();
        return shape;
    }

    /// <inheritdoc />
    public IAutoShape InsertAutoShape(int index, ShapeType shapeType, float x, float y, float width, float height)
    {
        return InsertAutoShape(index, shapeType, x, y, width, height, true);
    }

    /// <inheritdoc />
    public IAutoShape InsertAutoShape(int index, ShapeType shapeType, float x, float y, float width, float height, bool createFromTemplate)
    {
        var spTree = GetSpTree();
        var shapeId = NextShapeId();
        var name = $"AutoShape {shapeId}";
        var element = BuildAutoShapeXml(shapeId, name, shapeType, x, y, width, height, createFromTemplate);

        InsertElementAt(spTree, index, element);

        var shape = new AutoShape();
        shape.InitInternal(element, _slidePart, _parentSlide);
        _elementToShape[RuntimeHelpers.GetHashCode(element)] = shape;

        InvalidateCache();
        Save();
        return shape;
    }

    /// <inheritdoc />
    public IConnector AddConnector(ShapeType shapeType, float x, float y, float width, float height)
    {
        return AddConnector(shapeType, x, y, width, height, true);
    }

    /// <inheritdoc />
    public IConnector AddConnector(ShapeType shapeType, float x, float y, float width, float height, bool createFromTemplate)
    {
        var spTree = GetSpTree();
        var shapeId = NextShapeId();
        var name = $"Connector {shapeId}";
        var element = BuildConnectorXml(shapeId, name, shapeType, x, y, width, height, createFromTemplate);

        spTree?.Add(element);

        var connector = new Connector();
        connector.InitInternal(element, _slidePart, _parentSlide);
        _elementToShape[RuntimeHelpers.GetHashCode(element)] = connector;

        InvalidateCache();
        Save();
        return connector;
    }

    /// <inheritdoc />
    public IConnector InsertConnector(int index, ShapeType shapeType, float x, float y, float width, float height)
    {
        return InsertConnector(index, shapeType, x, y, width, height, true);
    }

    /// <inheritdoc />
    public IConnector InsertConnector(int index, ShapeType shapeType, float x, float y, float width, float height, bool createFromTemplate)
    {
        var spTree = GetSpTree();
        var shapeId = NextShapeId();
        var name = $"Connector {shapeId}";
        var element = BuildConnectorXml(shapeId, name, shapeType, x, y, width, height, createFromTemplate);

        InsertElementAt(spTree, index, element);

        var connector = new Connector();
        connector.InitInternal(element, _slidePart, _parentSlide);
        _elementToShape[RuntimeHelpers.GetHashCode(element)] = connector;

        InvalidateCache();
        Save();
        return connector;
    }

    /// <inheritdoc />
    public IPictureFrame AddPictureFrame(ShapeType shapeType, float x, float y, float width, float height, IPPImage image)
    {
        var spTree = GetSpTree();
        var shapeId = NextShapeId();
        var name = $"Picture {shapeId}";
        var element = BuildPictureFrameXml(shapeId, name, shapeType, x, y, width, height, image);

        spTree?.Add(element);

        var frame = new PictureFrame();
        frame.InitInternal(element, _slidePart, _parentSlide);
        _elementToShape[RuntimeHelpers.GetHashCode(element)] = frame;

        InvalidateCache();
        Save();
        return frame;
    }

    /// <inheritdoc />
    public IPictureFrame InsertPictureFrame(int index, ShapeType shapeType, float x, float y, float width, float height, IPPImage image)
    {
        var spTree = GetSpTree();
        var shapeId = NextShapeId();
        var name = $"Picture {shapeId}";
        var element = BuildPictureFrameXml(shapeId, name, shapeType, x, y, width, height, image);

        InsertElementAt(spTree, index, element);

        var frame = new PictureFrame();
        frame.InitInternal(element, _slidePart, _parentSlide);
        _elementToShape[RuntimeHelpers.GetHashCode(element)] = frame;

        InvalidateCache();
        Save();
        return frame;
    }

    /// <inheritdoc />
    public ITable AddTable(float x, float y, double[] columnWidths, double[] rowHeights)
    {
        var spTree = GetSpTree();
        var shapeId = NextShapeId();
        var name = $"Table {shapeId}";
        var element = BuildTableXml(shapeId, name, x, y, columnWidths, rowHeights);

        spTree?.Add(element);

        var table = new Table();
        table.InitInternal(element, _slidePart, _parentSlide);
        _elementToShape[RuntimeHelpers.GetHashCode(element)] = table;

        InvalidateCache();
        Save();
        return table;
    }

    /// <inheritdoc />
    public ITable InsertTable(int index, float x, float y, double[] columnWidths, double[] rowHeights)
    {
        var spTree = GetSpTree();
        var shapeId = NextShapeId();
        var name = $"Table {shapeId}";
        var element = BuildTableXml(shapeId, name, x, y, columnWidths, rowHeights);

        InsertElementAt(spTree, index, element);

        var table = new Table();
        table.InitInternal(element, _slidePart, _parentSlide);
        _elementToShape[RuntimeHelpers.GetHashCode(element)] = table;

        InvalidateCache();
        Save();
        return table;
    }

    /// <inheritdoc />
    public int IndexOf(IShape shape)
    {
        var shapes = LoadShapes();
        for (int i = 0; i < shapes.Count; i++)
        {
            if (ReferenceEquals(shapes[i], shape))
                return i;
        }

        return -1;
    }

    /// <inheritdoc />
    public void RemoveAt(int index)
    {
        var shapes = LoadShapes();
        if (index < 0 || index >= shapes.Count)
            throw new ArgumentOutOfRangeException(nameof(index));

        var shape = shapes[index];
        var spTree = GetSpTree();
        var element = FindShapeElement(spTree, shape);
        element?.Remove();

        RemoveFromElementMap(element);
        InvalidateCache();
        Save();
    }

    /// <inheritdoc />
    public void Remove(IShape shape)
    {
        var spTree = GetSpTree();
        var element = FindShapeElement(spTree, shape);
        if (element is null)
            throw new ArgumentException("Shape is not in the collection.", nameof(shape));

        element.Remove();
        RemoveFromElementMap(element);
        InvalidateCache();
        Save();
    }

    /// <inheritdoc />
    public void Clear()
    {
        var spTree = GetSpTree();
        if (spTree is null)
            return;

        var toRemove = GetShapeElements(spTree);
        foreach (var el in toRemove)
        {
            el.Remove();
        }

        _elementToShape.Clear();
        InvalidateCache();
        Save();
    }

    /// <inheritdoc />
    public IEnumerator<IShape> GetEnumerator() => LoadShapes().GetEnumerator();

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    // ───── Internal helpers ─────

    private XElement? GetSpTree()
    {
        var root = _slidePart?.Element;
        return root?.Descendants(PNs + "spTree").FirstOrDefault();
    }

    private List<IShape> LoadShapes()
    {
        if (_shapesCache is not null)
            return _shapesCache;

        var spTree = GetSpTree();
        if (spTree is null)
        {
            _shapesCache = new List<IShape>();
            return _shapesCache;
        }

        var result = new List<IShape>();
        foreach (var child in spTree.Elements())
        {
            if (SkippedLocalNames.Contains(child.Name.LocalName))
                continue;

            var key = RuntimeHelpers.GetHashCode(child);
            if (_elementToShape.TryGetValue(key, out var existing))
            {
                result.Add(existing);
            }
            else
            {
                var shape = ShapeFactory.CreateShape(child, _slidePart, _parentSlide);
                if (shape is not null)
                {
                    _elementToShape[key] = shape;
                    result.Add(shape);
                }
            }
        }

        _shapesCache = result;
        return _shapesCache;
    }

    private void InvalidateCache()
    {
        _shapesCache = null;
    }

    private void Save()
    {
        _slidePart?.Save();
    }

    private int NextShapeId()
    {
        var spTree = GetSpTree();
        if (spTree is null)
            return 2;

        int maxId = 1;
        foreach (var desc in spTree.Descendants())
        {
            var idAttr = desc.Attribute("id");
            if (idAttr is not null && int.TryParse(idAttr.Value, out var id) && id > maxId)
                maxId = id;
        }

        return maxId + 1;
    }

    private static string ToEmu(float points)
    {
        return ((int)MathF.Round(points * EmuPerPoint)).ToString();
    }

    private static List<XElement> GetShapeElements(XElement? spTree)
    {
        if (spTree is null)
            return new List<XElement>();

        var result = new List<XElement>();
        foreach (var child in spTree.Elements())
        {
            if (!SkippedLocalNames.Contains(child.Name.LocalName))
                result.Add(child);
        }

        return result;
    }

    private static XElement? FindShapeElement(XElement? spTree, IShape shape)
    {
        if (spTree is null || shape is not Shape s)
            return null;

        var target = s.GetElement();
        if (target is null)
            return null;

        foreach (var child in spTree.Elements())
        {
            if (SkippedLocalNames.Contains(child.Name.LocalName))
                continue;

            if (ReferenceEquals(target, child))
                return child;
        }

        return null;
    }

    private static void InsertElementAt(XElement? spTree, int index, XElement element)
    {
        if (spTree is null)
            return;

        var shapeElements = GetShapeElements(spTree);
        if (index >= shapeElements.Count)
        {
            spTree.Add(element);
        }
        else
        {
            shapeElements[index].AddBeforeSelf(element);
        }
    }

    private void RemoveFromElementMap(XElement? element)
    {
        if (element is not null)
            _elementToShape.Remove(RuntimeHelpers.GetHashCode(element));
    }

    // ───── XML builders ─────

    private static XElement BuildAutoShapeXml(int shapeId, string name, ShapeType shapeType,
        float x, float y, float width, float height, bool createFromTemplate)
    {
        var preset = OoxmlPresetMapping.ToPreset(shapeType) ?? "rect";

        var sp = new XElement(PNs + "sp",
            new XElement(PNs + "nvSpPr",
                new XElement(PNs + "cNvPr",
                    new XAttribute("id", shapeId),
                    new XAttribute("name", name)),
                new XElement(PNs + "cNvSpPr"),
                new XElement(PNs + "nvPr")),
            new XElement(PNs + "spPr",
                new XElement(ANs + "xfrm",
                    new XElement(ANs + "off",
                        new XAttribute("x", ToEmu(x)),
                        new XAttribute("y", ToEmu(y))),
                    new XElement(ANs + "ext",
                        new XAttribute("cx", ToEmu(width)),
                        new XAttribute("cy", ToEmu(height)))),
                new XElement(ANs + "prstGeom",
                    new XAttribute("prst", preset),
                    new XElement(ANs + "avLst"))));

        if (createFromTemplate)
        {
            sp.Add(BuildAutoShapeStyle());
        }

        sp.Add(new XElement(PNs + "txBody",
            new XElement(ANs + "bodyPr",
                new XAttribute("rtlCol", "0"),
                new XAttribute("anchor", "ctr")),
            new XElement(ANs + "lstStyle"),
            new XElement(ANs + "p",
                new XElement(ANs + "endParaRPr"))));

        return sp;
    }

    private static XElement BuildConnectorXml(int shapeId, string name, ShapeType shapeType,
        float x, float y, float width, float height, bool createFromTemplate)
    {
        var preset = OoxmlPresetMapping.ToPreset(shapeType) ?? "bentConnector3";

        var avLst = new XElement(ANs + "avLst");
        if (ConnectorDefaultAdjustments.Defaults.TryGetValue(preset, out var adjustments))
        {
            foreach (var (adjName, adjValue) in adjustments)
            {
                avLst.Add(new XElement(ANs + "gd",
                    new XAttribute("name", adjName),
                    new XAttribute("fmla", $"val {adjValue}")));
            }
        }

        var cxnSp = new XElement(PNs + "cxnSp",
            new XElement(PNs + "nvCxnSpPr",
                new XElement(PNs + "cNvPr",
                    new XAttribute("id", shapeId),
                    new XAttribute("name", name)),
                new XElement(PNs + "cNvCxnSpPr"),
                new XElement(PNs + "nvPr")),
            new XElement(PNs + "spPr",
                new XElement(ANs + "xfrm",
                    new XElement(ANs + "off",
                        new XAttribute("x", ToEmu(x)),
                        new XAttribute("y", ToEmu(y))),
                    new XElement(ANs + "ext",
                        new XAttribute("cx", ToEmu(width)),
                        new XAttribute("cy", ToEmu(height)))),
                new XElement(ANs + "prstGeom",
                    new XAttribute("prst", preset),
                    avLst)));

        if (createFromTemplate)
        {
            cxnSp.Add(BuildConnectorStyle());
        }

        return cxnSp;
    }

    private XElement BuildPictureFrameXml(int shapeId, string name, ShapeType shapeType,
        float x, float y, float width, float height, IPPImage image)
    {
        var preset = OoxmlPresetMapping.ToPreset(shapeType) ?? "rect";

        var rId = _slidePart?.RelsManager.AddImageRelationship(image) ?? "rId1";

        return new XElement(PNs + "pic",
            new XElement(PNs + "nvPicPr",
                new XElement(PNs + "cNvPr",
                    new XAttribute("id", shapeId),
                    new XAttribute("name", name)),
                new XElement(PNs + "cNvPicPr",
                    new XElement(ANs + "picLocks",
                        new XAttribute("noChangeAspect", "1"))),
                new XElement(PNs + "nvPr")),
            new XElement(PNs + "blipFill",
                new XElement(ANs + "blip",
                    new XAttribute(RNs + "embed", rId)),
                new XElement(ANs + "stretch",
                    new XElement(ANs + "fillRect"))),
            new XElement(PNs + "spPr",
                new XElement(ANs + "xfrm",
                    new XElement(ANs + "off",
                        new XAttribute("x", ToEmu(x)),
                        new XAttribute("y", ToEmu(y))),
                    new XElement(ANs + "ext",
                        new XAttribute("cx", ToEmu(width)),
                        new XAttribute("cy", ToEmu(height)))),
                new XElement(ANs + "prstGeom",
                    new XAttribute("prst", preset),
                    new XElement(ANs + "avLst"))));
    }

    private static XElement BuildTableXml(int shapeId, string name,
        float x, float y, double[] columnWidths, double[] rowHeights)
    {
        var totalWidth = 0.0;
        foreach (var w in columnWidths) totalWidth += w;
        var totalHeight = 0.0;
        foreach (var h in rowHeights) totalHeight += h;

        var tbl = new XElement(ANs + "tbl",
            new XElement(ANs + "tblPr",
                new XAttribute("firstRow", "1"),
                new XAttribute("bandRow", "1"),
                new XElement(ANs + "tblStyleId", "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")),
            new XElement(ANs + "tblGrid"));

        var tblGrid = tbl.Element(ANs + "tblGrid")!;
        foreach (var colWidth in columnWidths)
        {
            tblGrid.Add(new XElement(ANs + "gridCol",
                new XAttribute("w", ((int)Math.Round(colWidth * EmuPerPoint)).ToString())));
        }

        foreach (var rowHeight in rowHeights)
        {
            var tr = new XElement(ANs + "tr",
                new XAttribute("h", ((int)Math.Round(rowHeight * EmuPerPoint)).ToString()));

            for (int c = 0; c < columnWidths.Length; c++)
            {
                tr.Add(new XElement(ANs + "tc",
                    new XElement(ANs + "txBody",
                        new XElement(ANs + "bodyPr"),
                        new XElement(ANs + "lstStyle"),
                        new XElement(ANs + "p",
                            new XElement(ANs + "endParaRPr"))),
                    new XElement(ANs + "tcPr")));
            }

            tbl.Add(tr);
        }

        return new XElement(PNs + "graphicFrame",
            new XElement(PNs + "nvGraphicFramePr",
                new XElement(PNs + "cNvPr",
                    new XAttribute("id", shapeId),
                    new XAttribute("name", name)),
                new XElement(PNs + "cNvGraphicFramePr",
                    new XElement(ANs + "graphicFrameLocks",
                        new XAttribute("noGrp", "1"))),
                new XElement(PNs + "nvPr")),
            new XElement(PNs + "xfrm",
                new XElement(ANs + "off",
                    new XAttribute("x", ((int)MathF.Round(x * EmuPerPoint)).ToString()),
                    new XAttribute("y", ((int)MathF.Round(y * EmuPerPoint)).ToString())),
                new XElement(ANs + "ext",
                    new XAttribute("cx", ((int)Math.Round(totalWidth * EmuPerPoint)).ToString()),
                    new XAttribute("cy", ((int)Math.Round(totalHeight * EmuPerPoint)).ToString()))),
            new XElement(ANs + "graphic",
                new XElement(ANs + "graphicData",
                    new XAttribute("uri", "http://schemas.openxmlformats.org/drawingml/2006/table"),
                    tbl)));
    }

    private static XElement BuildAutoShapeStyle()
    {
        return new XElement(PNs + "style",
            new XElement(ANs + "lnRef",
                new XAttribute("idx", "2"),
                new XElement(ANs + "schemeClr",
                    new XAttribute("val", "accent1"),
                    new XElement(ANs + "shade",
                        new XAttribute("val", "50000")))),
            new XElement(ANs + "fillRef",
                new XAttribute("idx", "1"),
                new XElement(ANs + "schemeClr",
                    new XAttribute("val", "accent1"))),
            new XElement(ANs + "effectRef",
                new XAttribute("idx", "0"),
                new XElement(ANs + "schemeClr",
                    new XAttribute("val", "accent1"))),
            new XElement(ANs + "fontRef",
                new XAttribute("idx", "minor"),
                new XElement(ANs + "schemeClr",
                    new XAttribute("val", "lt1"))));
    }

    private static XElement BuildConnectorStyle()
    {
        return new XElement(PNs + "style",
            new XElement(ANs + "lnRef",
                new XAttribute("idx", "1"),
                new XElement(ANs + "schemeClr",
                    new XAttribute("val", "accent1"))),
            new XElement(ANs + "fillRef",
                new XAttribute("idx", "0"),
                new XElement(ANs + "schemeClr",
                    new XAttribute("val", "accent1"))),
            new XElement(ANs + "effectRef",
                new XAttribute("idx", "0"),
                new XElement(ANs + "schemeClr",
                    new XAttribute("val", "accent1"))),
            new XElement(ANs + "fontRef",
                new XAttribute("idx", "minor"),
                new XElement(ANs + "schemeClr",
                    new XAttribute("val", "tx1"))));
    }
}
