using System.Globalization;
using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a picture fill style.
/// </summary>
public sealed class PictureFillFormat : PVIObject, IPictureFillFormat, IFillParamSource
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static readonly Dictionary<PictureFillMode, string> FillModeToXml = new()
    {
        [PictureFillMode.Stretch] = "stretch",
        [PictureFillMode.Tile] = "tile",
    };

    private static readonly Dictionary<RectangleAlignment, string> AlignToXml = new()
    {
        [RectangleAlignment.TopLeft] = "tl",
        [RectangleAlignment.Top] = "t",
        [RectangleAlignment.TopRight] = "tr",
        [RectangleAlignment.Left] = "l",
        [RectangleAlignment.Center] = "ctr",
        [RectangleAlignment.Right] = "r",
        [RectangleAlignment.BottomLeft] = "bl",
        [RectangleAlignment.Bottom] = "b",
        [RectangleAlignment.BottomRight] = "br",
    };

    private static readonly Dictionary<string, RectangleAlignment> XmlToAlign =
        AlignToXml.ToDictionary(kvp => kvp.Value, kvp => kvp.Key);

    private static readonly Dictionary<TileFlip, string> FlipToXml = new()
    {
        [TileFlip.NoFlip] = "none",
        [TileFlip.FlipX] = "x",
        [TileFlip.FlipY] = "y",
        [TileFlip.FlipBoth] = "xy",
    };

    private static readonly Dictionary<string, TileFlip> XmlToFlip =
        FlipToXml.ToDictionary(kvp => kvp.Value, kvp => kvp.Key);

    private XElement? _blipFill;
    private SlidePart? _slidePart;

    /// <summary>
    /// Initializes internal state with the blipFill XML element.
    /// </summary>
    /// <param name="blipFill">The a:blipFill element.</param>
    /// <param name="slidePart">The SlidePart for relationship resolution.</param>
    /// <param name="parentSlide">The parent slide object.</param>
    internal void InitInternal(XElement blipFill, SlidePart? slidePart, IBaseSlide? parentSlide = null)
    {
        _blipFill = blipFill;
        _slidePart = slidePart;
        _parentSlide = parentSlide;
    }

    /// <summary>
    /// Persists changes to the underlying package.
    /// </summary>
    internal void Save() => _slidePart?.Save();

    /// <summary>
    /// Gets the a:blip element.
    /// </summary>
    /// <returns>The blip element, or <c>null</c> if not found.</returns>
    internal XElement? GetBlip() => _blipFill?.Element(ANs + "blip");

    /// <summary>
    /// Gets the a:stretch element.
    /// </summary>
    /// <returns>The stretch element, or <c>null</c> if not found.</returns>
    internal XElement? GetStretch() => _blipFill?.Element(ANs + "stretch");

    /// <summary>
    /// Gets or creates the a:srcRect element for crop values.
    /// </summary>
    /// <returns>The srcRect element.</returns>
    internal XElement? GetOrCreateSrcRect()
    {
        if (_blipFill is null) return null;
        var srcRect = _blipFill.Element(ANs + "srcRect");
        if (srcRect is null)
        {
            srcRect = new XElement(ANs + "srcRect");
            // srcRect goes after blip
            var blip = _blipFill.Element(ANs + "blip");
            if (blip is not null)
                blip.AddAfterSelf(srcRect);
            else
                _blipFill.AddFirst(srcRect);
        }
        return srcRect;
    }

    /// <summary>
    /// Gets or creates the a:fillRect element under a:stretch.
    /// </summary>
    /// <returns>The fillRect element.</returns>
    internal XElement? GetOrCreateFillRect()
    {
        if (_blipFill is null) return null;
        var stretch = _blipFill.Element(ANs + "stretch");
        if (stretch is null)
        {
            stretch = new XElement(ANs + "stretch");
            _blipFill.Add(stretch);
        }
        var fillRect = stretch.Element(ANs + "fillRect");
        if (fillRect is null)
        {
            fillRect = new XElement(ANs + "fillRect");
            stretch.Add(fillRect);
        }
        return fillRect;
    }

    /// <summary>
    /// Gets a crop value from a:srcRect. Returns percentage (e.g., 10.0 for 10%).
    /// </summary>
    /// <param name="attr">The attribute name (l, t, r, b).</param>
    /// <returns>The crop value as a percentage.</returns>
    internal float GetCropValue(string attr)
    {
        var srcRect = _blipFill?.Element(ANs + "srcRect");
        var val = srcRect?.Attribute(attr)?.Value;
        if (val is not null && int.TryParse(val, out var emu))
            return emu / 1000f;
        return 0f;
    }

    /// <summary>
    /// Sets a crop value on a:srcRect. Value is percentage (e.g., 10.0 for 10%).
    /// </summary>
    /// <param name="attr">The attribute name (l, t, r, b).</param>
    /// <param name="value">The crop value as a percentage.</param>
    internal void SetCropValue(string attr, float value)
    {
        var srcRect = GetOrCreateSrcRect();
        if (srcRect is null) return;
        srcRect.SetAttributeValue(attr, ((int)(value * 1000)).ToString(CultureInfo.InvariantCulture));
        Save();
    }

    /// <summary>
    /// Gets a stretch offset value from a:fillRect under a:stretch.
    /// </summary>
    /// <param name="attr">The attribute name (l, t, r, b).</param>
    /// <returns>The stretch offset value as a percentage.</returns>
    internal float GetStretchOffset(string attr)
    {
        var fillRect = _blipFill?.Element(ANs + "stretch")?.Element(ANs + "fillRect");
        var val = fillRect?.Attribute(attr)?.Value;
        if (val is not null && int.TryParse(val, out var emu))
            return emu / 1000f;
        return 0f;
    }

    /// <summary>
    /// Sets a stretch offset value on a:fillRect under a:stretch.
    /// </summary>
    /// <param name="attr">The attribute name (l, t, r, b).</param>
    /// <param name="value">The stretch offset value as a percentage.</param>
    internal void SetStretchOffset(string attr, float value)
    {
        var fillRect = GetOrCreateFillRect();
        if (fillRect is null) return;
        fillRect.SetAttributeValue(attr, ((int)(value * 1000)).ToString(CultureInfo.InvariantCulture));
        Save();
    }

    /// <summary>
    /// Gets the a:tile element.
    /// </summary>
    /// <returns>The tile element, or <c>null</c> if not found.</returns>
    internal XElement? GetTile() => _blipFill?.Element(ANs + "tile");

    /// <summary>
    /// Gets or creates the a:tile element, removing a:stretch if present (switching from stretch to tile mode).
    /// </summary>
    /// <returns>The tile element.</returns>
    internal XElement? EnsureTile()
    {
        if (_blipFill is null) return null;
        var tile = _blipFill.Element(ANs + "tile");
        if (tile is not null) return tile;
        // Remove stretch if present (switching to tile mode)
        _blipFill.Element(ANs + "stretch")?.Remove();
        tile = new XElement(ANs + "tile");
        _blipFill.Add(tile);
        return tile;
    }

    /// <inheritdoc/>
    public int Dpi
    {
        get
        {
            var val = _blipFill?.Attribute("dpi")?.Value;
            return val is not null && int.TryParse(val, out var dpi) ? dpi : 0;
        }
        set
        {
            if (_blipFill is null) return;
            _blipFill.SetAttributeValue("dpi", value);
            Save();
        }
    }

    /// <inheritdoc/>
    public PictureFillMode PictureFillMode
    {
        get
        {
            if (_blipFill is null) return PictureFillMode.Tile;
            if (_blipFill.Element(ANs + "stretch") is not null)
                return PictureFillMode.Stretch;
            return PictureFillMode.Tile;
        }
        set
        {
            if (_blipFill is null) return;

            // Remove existing stretch/tile elements
            _blipFill.Element(ANs + "stretch")?.Remove();
            _blipFill.Element(ANs + "tile")?.Remove();

            if (value == PictureFillMode.Stretch)
            {
                var stretch = new XElement(ANs + "stretch");
                stretch.Add(new XElement(ANs + "fillRect"));
                _blipFill.Add(stretch);
            }
            else
            {
                _blipFill.Add(new XElement(ANs + "tile"));
            }
            Save();
        }
    }

    /// <inheritdoc/>
    public ISlidesPicture Picture
    {
        get
        {
            var blip = _blipFill?.Element(ANs + "blip");
            if (blip is null && _blipFill is not null)
            {
                blip = new XElement(ANs + "blip");
                _blipFill.AddFirst(blip);
            }
            var pic = new Picture();
            pic.InitInternal(blip, _slidePart, parentSlide: _parentSlide);
            return pic;
        }
    }

    /// <inheritdoc/>
    public float CropLeft
    {
        get => GetCropValue("l");
        set => SetCropValue("l", value);
    }

    /// <inheritdoc/>
    public float CropTop
    {
        get => GetCropValue("t");
        set => SetCropValue("t", value);
    }

    /// <inheritdoc/>
    public float CropRight
    {
        get => GetCropValue("r");
        set => SetCropValue("r", value);
    }

    /// <inheritdoc/>
    public float CropBottom
    {
        get => GetCropValue("b");
        set => SetCropValue("b", value);
    }

    /// <inheritdoc/>
    public float StretchOffsetLeft
    {
        get => GetStretchOffset("l");
        set => SetStretchOffset("l", value);
    }

    /// <inheritdoc/>
    public float StretchOffsetTop
    {
        get => GetStretchOffset("t");
        set => SetStretchOffset("t", value);
    }

    /// <inheritdoc/>
    public float StretchOffsetRight
    {
        get => GetStretchOffset("r");
        set => SetStretchOffset("r", value);
    }

    /// <inheritdoc/>
    public float StretchOffsetBottom
    {
        get => GetStretchOffset("b");
        set => SetStretchOffset("b", value);
    }

    /// <inheritdoc/>
    public float TileOffsetX
    {
        get => GetTileAttribute("tx");
        set => SetTileAttribute("tx", value);
    }

    /// <inheritdoc/>
    public float TileOffsetY
    {
        get => GetTileAttribute("ty");
        set => SetTileAttribute("ty", value);
    }

    /// <inheritdoc/>
    public float TileScaleX
    {
        get => GetTileAttribute("sx");
        set => SetTileAttribute("sx", value);
    }

    /// <inheritdoc/>
    public float TileScaleY
    {
        get => GetTileAttribute("sy");
        set => SetTileAttribute("sy", value);
    }

    /// <inheritdoc/>
    public RectangleAlignment TileAlignment
    {
        get
        {
            var tile = _blipFill?.Element(ANs + "tile");
            var val = tile?.Attribute("algn")?.Value;
            if (val is not null && XmlToAlign.TryGetValue(val, out var align))
                return align;
            return RectangleAlignment.NotDefined;
        }
        set
        {
            var tile = EnsureTile();
            if (tile is null) return;
            if (value == RectangleAlignment.NotDefined)
                tile.Attribute("algn")?.Remove();
            else if (AlignToXml.TryGetValue(value, out var xmlVal))
                tile.SetAttributeValue("algn", xmlVal);
            Save();
        }
    }

    /// <inheritdoc/>
    public TileFlip TileFlip
    {
        get
        {
            var tile = _blipFill?.Element(ANs + "tile");
            var val = tile?.Attribute("flip")?.Value;
            if (val is not null && XmlToFlip.TryGetValue(val, out var flip))
                return flip;
            return TileFlip.NotDefined;
        }
        set
        {
            var tile = EnsureTile();
            if (tile is null) return;
            if (value == TileFlip.NotDefined)
                tile.Attribute("flip")?.Remove();
            else if (FlipToXml.TryGetValue(value, out var xmlVal))
                tile.SetAttributeValue("flip", xmlVal);
            Save();
        }
    }

    // --- private helpers ---

    private float GetTileAttribute(string attr)
    {
        var tile = _blipFill?.Element(ANs + "tile");
        var val = tile?.Attribute(attr)?.Value;
        if (val is not null && float.TryParse(val, NumberStyles.Float, CultureInfo.InvariantCulture, out var result))
            return result;
        return 0f;
    }

    private void SetTileAttribute(string attr, float value)
    {
        var tile = EnsureTile();
        if (tile is null) return;
        tile.SetAttributeValue(attr, value.ToString(CultureInfo.InvariantCulture));
        Save();
    }
}
