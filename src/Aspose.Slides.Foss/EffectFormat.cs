using System.Xml.Linq;
using Aspose.Slides.Foss.Effects;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents effect formatting properties backed by an <c>effectLst</c> element.
/// </summary>
public sealed class EffectFormat : PVIObject, IEffectFormat, IEffectParamSource
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private const float EmuPerPoint = 12700f;

    /// <summary>
    /// Defines the OOXML-mandated child ordering within <c>effectLst</c>.
    /// </summary>
    private static readonly string[] EffectLstOrder =
    [
        "blur",
        "fillOverlay",
        "glow",
        "innerShdw",
        "outerShdw",
        "prstShdw",
        "reflection",
        "softEdge",
    ];

    /// <summary>
    /// Tags that must appear after <c>effectLst</c> within the parent element (e.g., <c>spPr</c>).
    /// </summary>
    private static readonly string[] AfterEffectLstTags = ["scene3d", "sp3d", "extLst"];

    /// <summary>
    /// Alpha applied to the default effect colour, in thousandths of a percent (40%).
    /// </summary>
    private const string DefaultAlpha = "40000";

    /// <summary>
    /// Default radius for effects whose radius attribute is required or whose zero value would make
    /// the effect invisible, in EMU (5 pt).
    /// </summary>
    private const string DefaultRadiusEmu = "63500";

    private XElement? _parentElement;
    private SlidePart? _slidePart;

    /// <summary>
    /// Initializes internal state for this effect format.
    /// </summary>
    /// <param name="parentElement">The XML element that may contain <c>&lt;a:effectLst&gt;</c>.</param>
    /// <param name="parentSlide">The parent slide object.</param>
    /// <param name="slidePart">The slide part for saving changes.</param>
    internal void InitInternal(XElement parentElement, IBaseSlide? parentSlide, SlidePart? slidePart = null)
    {
        _parentElement = parentElement;
        _parentSlide = parentSlide;
        _slidePart = slidePart;
    }

    private XElement? GetEffectLst() => _parentElement?.Element(ANs + "effectLst");

    private XElement EnsureEffectLst()
    {
        var el = _parentElement!.Element(ANs + "effectLst");
        if (el is not null)
            return el;

        el = new XElement(ANs + "effectLst");

        // Insert before scene3d, sp3d, or extLst to maintain OOXML ordering.
        foreach (var tag in AfterEffectLstTags)
        {
            var sibling = _parentElement!.Element(ANs + tag);
            if (sibling is not null)
            {
                sibling.AddBeforeSelf(el);
                return el;
            }
        }

        _parentElement!.Add(el);
        return el;
    }

    private XElement? GetEffectChild(string tag)
    {
        return GetEffectLst()?.Element(ANs + tag);
    }

    private XElement EnsureEffectChild(string tag)
    {
        var lst = EnsureEffectLst();
        var existing = lst.Element(ANs + tag);
        if (existing is not null)
            return existing;

        var newEl = new XElement(ANs + tag);
        ApplySchemaDefaults(newEl, tag);
        int rank = Array.IndexOf(EffectLstOrder, tag);

        if (rank < 0)
        {
            lst.Add(newEl);
            return newEl;
        }

        // Insert before the first existing sibling with a higher rank.
        foreach (var child in lst.Elements())
        {
            var childLocal = child.Name.LocalName;
            int childRank = Array.IndexOf(EffectLstOrder, childLocal);
            if (childRank > rank)
            {
                child.AddBeforeSelf(newEl);
                return newEl;
            }
        }

        lst.Add(newEl);
        return newEl;
    }

    /// <summary>
    /// Fills in the attributes and children ECMA-376 §20.1.8 requires of a newly created effect
    /// element.
    /// </summary>
    /// <remarks>
    /// <para>
    /// An effect element that omits a required attribute, or a shadow or glow that carries no
    /// colour child, is not a subtle formatting difference: the schema calls the element
    /// incomplete and PowerPoint refuses to open the whole file. The values here are neutral
    /// starting points that the caller is expected to overwrite through the effect's own
    /// properties; the point is that the element is complete the moment it exists.
    /// </para>
    /// <para>
    /// The colour is a literal <c>srgbClr</c> rather than a <c>schemeClr</c> so that it does not
    /// depend on a theme carrying any particular named colour.
    /// </para>
    /// <para>
    /// Applied only to elements this method creates. An element that is already in the document
    /// carries the caller's values and is left untouched.
    /// </para>
    /// </remarks>
    /// <param name="element">The freshly created effect element.</param>
    /// <param name="tag">Its local name.</param>
    private static void ApplySchemaDefaults(XElement element, string tag)
    {
        switch (tag)
        {
            // CT_OuterShadowEffect / CT_InnerShadowEffect: exactly one EG_ColorChoice is required.
            case "outerShdw":
            case "innerShdw":
                element.Add(DefaultColor());
                break;

            // CT_GlowEffect: one EG_ColorChoice is required. rad is optional but defaults to 0,
            // which is a glow of no width, so it is set as well.
            case "glow":
                element.SetAttributeValue("rad", DefaultRadiusEmu);
                element.Add(DefaultColor());
                break;

            // CT_PresetShadowEffect: prst is required, and one EG_ColorChoice with it.
            case "prstShdw":
                element.SetAttributeValue("prst", "shdw1");
                element.Add(DefaultColor());
                break;

            // CT_SoftEdgesEffect: rad is required.
            case "softEdge":
                element.SetAttributeValue("rad", DefaultRadiusEmu);
                break;

            // CT_FillOverlayEffect: blend is required, and exactly one EG_FillProperties.
            case "fillOverlay":
                element.SetAttributeValue("blend", "over");
                element.Add(new XElement(ANs + "noFill"));
                break;

            // CT_BlurEffect and CT_ReflectionEffect have no required attribute and no required
            // child; an empty element is already schema-valid.
            default:
                break;
        }
    }

    private static XElement DefaultColor() =>
        new(ANs + "srgbClr",
            new XAttribute("val", "000000"),
            new XElement(ANs + "alpha", new XAttribute("val", DefaultAlpha)));

    private void RemoveEffectChild(string tag)
    {
        var lst = GetEffectLst();
        var child = lst?.Element(ANs + tag);
        if (child is null)
            return;

        child.Remove();

        // lst is guaranteed non-null here because child was found within it.
        if (lst is not null && !lst.HasElements)
            lst.Remove();
    }

    private void Save()
    {
        _slidePart?.Save();
    }

    /// <inheritdoc/>
    public bool IsNoEffects => GetEffectLst()?.HasElements != true;

    /// <inheritdoc/>
    public IBlur? BlurEffect
    {
        get
        {
            var el = GetEffectChild("blur");
            return el is not null ? new Blur(el) : null;
        }
        set
        {
            if (value is null)
            {
                RemoveEffectChild("blur");
            }
            else
            {
                var el = EnsureEffectChild("blur");
                CopyAttributes((value as Blur)?.Element, el);
            }
            Save();
        }
    }

    /// <inheritdoc/>
    public IFillOverlay? FillOverlayEffect
    {
        get
        {
            var el = GetEffectChild("fillOverlay");
            return el is not null ? new FillOverlay(el) : null;
        }
        set
        {
            if (value is null)
            {
                RemoveEffectChild("fillOverlay");
            }
            else
            {
                var el = EnsureEffectChild("fillOverlay");
                CopyAttributes((value as FillOverlay)?.Element, el);
            }
            Save();
        }
    }

    /// <inheritdoc/>
    public IGlow? GlowEffect
    {
        get
        {
            var el = GetEffectChild("glow");
            return el is not null ? new Glow(el) : null;
        }
        set
        {
            if (value is null)
            {
                RemoveEffectChild("glow");
            }
            else
            {
                var el = EnsureEffectChild("glow");
                CopyAttributes((value as Glow)?.Element, el);
            }
            Save();
        }
    }

    /// <inheritdoc/>
    public IInnerShadow? InnerShadowEffect
    {
        get
        {
            var el = GetEffectChild("innerShdw");
            return el is not null ? new InnerShadow(el) : null;
        }
        set
        {
            if (value is null)
            {
                RemoveEffectChild("innerShdw");
            }
            else
            {
                var el = EnsureEffectChild("innerShdw");
                CopyAttributes((value as InnerShadow)?.Element, el);
            }
            Save();
        }
    }

    /// <inheritdoc/>
    public IOuterShadow? OuterShadowEffect
    {
        get
        {
            var el = GetEffectChild("outerShdw");
            return el is not null ? new OuterShadow(el) : null;
        }
        set
        {
            if (value is null)
            {
                RemoveEffectChild("outerShdw");
            }
            else
            {
                var el = EnsureEffectChild("outerShdw");
                CopyAttributes((value as OuterShadow)?.Element, el);
            }
            Save();
        }
    }

    /// <inheritdoc/>
    public IPresetShadow? PresetShadowEffect
    {
        get
        {
            var el = GetEffectChild("prstShdw");
            return el is not null ? new PresetShadow(el) : null;
        }
        set
        {
            if (value is null)
            {
                RemoveEffectChild("prstShdw");
            }
            else
            {
                var el = EnsureEffectChild("prstShdw");
                CopyAttributes((value as PresetShadow)?.Element, el);
            }
            Save();
        }
    }

    /// <inheritdoc/>
    public IReflection? ReflectionEffect
    {
        get
        {
            var el = GetEffectChild("reflection");
            return el is not null ? new Reflection(el) : null;
        }
        set
        {
            if (value is null)
            {
                RemoveEffectChild("reflection");
            }
            else
            {
                var el = EnsureEffectChild("reflection");
                CopyAttributes((value as Reflection)?.Element, el);
            }
            Save();
        }
    }

    /// <inheritdoc/>
    public ISoftEdge? SoftEdgeEffect
    {
        get
        {
            var el = GetEffectChild("softEdge");
            return el is not null ? new SoftEdge(el) : null;
        }
        set
        {
            if (value is null)
            {
                RemoveEffectChild("softEdge");
            }
            else
            {
                var el = EnsureEffectChild("softEdge");
                CopyAttributes((value as SoftEdge)?.Element, el);
            }
            Save();
        }
    }

    /// <inheritdoc/>
    public IEffectParamSource AsIEffectParamSource => this;

    /// <inheritdoc/>
    public void SetBlurEffect(float radius, bool grow)
    {
        var el = EnsureEffectChild("blur");
        el.SetAttributeValue("rad", (int)(radius * EmuPerPoint));
        el.SetAttributeValue("grow", grow ? "1" : "0");
        Save();
    }

    /// <inheritdoc/>
    public void EnableBlurEffect()
    {
        EnsureEffectChild("blur");
        Save();
    }

    /// <inheritdoc/>
    public void DisableBlurEffect()
    {
        RemoveEffectChild("blur");
        Save();
    }

    /// <inheritdoc/>
    public void EnableFillOverlayEffect()
    {
        EnsureEffectChild("fillOverlay");
        Save();
    }

    /// <inheritdoc/>
    public void DisableFillOverlayEffect()
    {
        RemoveEffectChild("fillOverlay");
        Save();
    }

    /// <inheritdoc/>
    public void EnableGlowEffect()
    {
        EnsureEffectChild("glow");
        Save();
    }

    /// <inheritdoc/>
    public void DisableGlowEffect()
    {
        RemoveEffectChild("glow");
        Save();
    }

    /// <inheritdoc/>
    public void EnableInnerShadowEffect()
    {
        EnsureEffectChild("innerShdw");
        Save();
    }

    /// <inheritdoc/>
    public void DisableInnerShadowEffect()
    {
        RemoveEffectChild("innerShdw");
        Save();
    }

    /// <inheritdoc/>
    public void EnableOuterShadowEffect()
    {
        EnsureEffectChild("outerShdw");
        Save();
    }

    /// <inheritdoc/>
    public void DisableOuterShadowEffect()
    {
        RemoveEffectChild("outerShdw");
        Save();
    }

    /// <inheritdoc/>
    public void EnablePresetShadowEffect()
    {
        EnsureEffectChild("prstShdw");
        Save();
    }

    /// <inheritdoc/>
    public void DisablePresetShadowEffect()
    {
        RemoveEffectChild("prstShdw");
        Save();
    }

    /// <inheritdoc/>
    public void EnableReflectionEffect()
    {
        EnsureEffectChild("reflection");
        Save();
    }

    /// <inheritdoc/>
    public void DisableReflectionEffect()
    {
        RemoveEffectChild("reflection");
        Save();
    }

    /// <inheritdoc/>
    public void EnableSoftEdgeEffect()
    {
        EnsureEffectChild("softEdge");
        Save();
    }

    /// <inheritdoc/>
    public void DisableSoftEdgeEffect()
    {
        RemoveEffectChild("softEdge");
        Save();
    }

    /// <summary>
    /// Copies all attributes from a source element to a target element.
    /// </summary>
    private static void CopyAttributes(XElement? source, XElement target)
    {
        if (source is null)
            return;

        foreach (var attr in source.Attributes())
        {
            target.SetAttributeValue(attr.Name, attr.Value);
        }
    }
}
