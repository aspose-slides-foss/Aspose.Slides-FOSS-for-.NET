using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Concrete picture reference backed by an <c>a:blip</c> XML element in a slide's XML.
/// </summary>
public sealed class Picture : ISlidesPicture
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace RNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly string HyperlinkRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    private static readonly XName PendingPartNameAttr = "_pendingPartName";

    private XElement? _blipElement;
    private SlidePart? _slidePart;
    private IBaseSlide? _parentSlideRef;

    /// <summary>
    /// Initializes internal state with the blip element, slide part, and parent slide.
    /// </summary>
    internal void InitInternal(XElement? blipElement, SlidePart? slidePart, IBaseSlide? parentSlide)
    {
        _blipElement = blipElement;
        _slidePart = slidePart;
        _parentSlideRef = parentSlide;
    }

    /// <inheritdoc/>
    public override IPresentationComponent AsIPresentationComponent => this;

    /// <inheritdoc/>
    public override ISlideComponent AsISlideComponent => this;

    /// <inheritdoc/>
    public override IBaseSlide? Slide => _parentSlideRef;

    /// <inheritdoc/>
    public override IPresentation? Presentation => _parentSlideRef?.Presentation;

    /// <inheritdoc/>
    public override IPPImage? Image
    {
        get
        {
            if (_blipElement is null || _slidePart is null)
                return null;

            var embedAttr = _blipElement.Attribute(RNs + "embed");
            if (embedAttr is null)
                return null;

            var rel = _slidePart.RelsManager.GetById(embedAttr.Value);
            if (rel is null)
                return null;

            var absolutePartName = OpcPaths.Resolve(
                OpcPaths.DirectoryOf(_slidePart.PartName), rel.Target);

            var presentation = _parentSlideRef?.Presentation;
            if (presentation is not Presentation pres)
                return null;

            return pres.ImagesInternal?.FindByPartName(absolutePartName);
        }
        set
        {
            if (_blipElement is null || value is not PPImage ppImage)
                return;

            if (_slidePart is null)
            {
                // Detached mode: store pending part name for later resolution.
                _blipElement.SetAttributeValue(PendingPartNameAttr, ppImage.PartName);
            }
            else
            {
                SetBlipImage(_blipElement, _slidePart, ppImage);
            }
        }
    }

    /// <inheritdoc/>
    public override string LinkPathLong
    {
        get
        {
            if (_blipElement is null || _slidePart is null)
                return string.Empty;

            var linkAttr = _blipElement.Attribute(RNs + "link");
            if (linkAttr is null)
                return string.Empty;

            var rel = _slidePart.RelsManager.GetById(linkAttr.Value);
            return rel?.Target ?? string.Empty;
        }
        set
        {
            if (_blipElement is null || _slidePart is null)
                return;

            // Remove existing link relationship.
            var linkAttr = _blipElement.Attribute(RNs + "link");
            if (linkAttr is not null)
            {
                _slidePart.RelsManager.Remove(linkAttr.Value);
                linkAttr.Remove();
            }

            if (!string.IsNullOrEmpty(value))
            {
                var relId = _slidePart.RelsManager.Add(HyperlinkRelType, value, isExternal: true);
                _blipElement.SetAttributeValue(RNs + "link", relId);
                _slidePart.RelsManager.Save();
            }
        }
    }

    /// <summary>
    /// Resolves deferred image references in all <c>a:blip</c> elements within the given XML subtree.
    /// </summary>
    /// <param name="element">The root XML element to scan.</param>
    /// <param name="slidePart">The slide part for relationship resolution.</param>
    /// <param name="parentSlide">The parent slide for presentation access.</param>
    internal static void FlushPendingBlipImages(XElement element, SlidePart slidePart, IBaseSlide parentSlide)
    {
        foreach (var blip in element.Descendants(ANs + "blip"))
        {
            var pendingAttr = blip.Attribute(PendingPartNameAttr);
            if (pendingAttr is null)
                continue;

            var pendingPartName = pendingAttr.Value;
            pendingAttr.Remove();

            var presentation = parentSlide.Presentation;
            if (presentation is not Presentation pres)
                continue;

            var ppImage = pres.ImagesInternal?.FindByPartName(pendingPartName);
            if (ppImage is null)
                continue;

            SetBlipImage(blip, slidePart, ppImage);
        }
    }

    private static void SetBlipImage(XElement blip, SlidePart slidePart, PPImage ppImage)
    {
        var relId = slidePart.RelsManager.EnsureImageRelationship(ppImage.PartName);

        blip.SetAttributeValue(RNs + "embed", relId);
        slidePart.Save();
    }
}
