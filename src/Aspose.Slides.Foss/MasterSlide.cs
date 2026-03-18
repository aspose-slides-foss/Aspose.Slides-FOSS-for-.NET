using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a master slide in a presentation.
/// </summary>
public sealed class MasterSlide : BaseSlide, IMasterSlide
{
    private MasterLayoutSlideCollection? _layoutSlides;

    /// <inheritdoc />
    public IMasterLayoutSlideCollection LayoutSlides =>
        _layoutSlides ??= new MasterLayoutSlideCollection();

    /// <summary>
    /// Initializes internal state. Called after default construction to inject dependencies.
    /// </summary>
    /// <param name="presentation">The parent Presentation object.</param>
    /// <param name="package">The OPC package.</param>
    /// <param name="partName">The part name of this master slide.</param>
    /// <param name="masterPart">The parsed SlidePart for this master slide.</param>
    /// <param name="layoutSlides">Optional list of LayoutSlide objects belonging to this master.</param>
    internal void InitInternal(
        IPresentation presentation,
        OpcPackage package,
        string partName,
        SlidePart masterPart,
        List<LayoutSlide>? layoutSlides = null)
    {
        _presentationRef = presentation;
        PackageInternal = package;
        _partName = partName;
        _masterPart = masterPart;

        if (layoutSlides is not null)
        {
            var collection = new MasterLayoutSlideCollection();
            collection.InitInternal(layoutSlides.Cast<ILayoutSlide>().ToList());
            _layoutSlides = collection;
        }
    }

    /// <summary>
    /// Gets or sets the OPC part name for this master slide (e.g. "ppt/slideMasters/slideMaster1.xml").
    /// </summary>
    internal string PartName
    {
        get => _partName ?? "";
        set => _partName = value;
    }

    /// <summary>
    /// Gets or sets the OPC package reference for this master slide.
    /// </summary>
    internal OpcPackage? PackageInternal { get; set; }

    /// <summary>
    /// Sets the presentation reference for this master slide. Internal use only.
    /// </summary>
    internal void SetPresentation(IPresentation? presentation)
    {
        _presentationRef = presentation;
    }

    /// <summary>
    /// Sets the master part for this master slide. Internal use only.
    /// </summary>
    internal void SetMasterPart(SlidePart? masterPart)
    {
        _masterPart = masterPart;
    }

    /// <summary>
    /// Sets the layout slides collection for this master slide.
    /// </summary>
    internal void SetLayoutSlides(MasterLayoutSlideCollection layouts)
    {
        _layoutSlides = layouts;
    }
}
