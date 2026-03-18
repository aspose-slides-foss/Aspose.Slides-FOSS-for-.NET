using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;
using Aspose.Slides.Foss.Internal.Pptx.PresentationPart;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a slide in a presentation.
/// </summary>
public sealed class Slide : BaseSlide, ISlide
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";

    private int _slideNumber;
    private bool _hiddenFallback;
    private string _nameFallback = string.Empty;
    private ILayoutSlide? _layoutSlide;
    private NotesSlideManager? _notesSlideManager;
    private Func<string, ILayoutSlide?>? _layoutResolver;

    /// <summary>
    /// Gets or sets the OPC part name for this slide (e.g. "ppt/slides/slide1.xml").
    /// </summary>
    internal string PartName { get; set; } = "";

    /// <summary>
    /// Gets or sets the OPC package reference for this slide.
    /// </summary>
    internal OpcPackage? PackageInternal { get; set; }

    /// <summary>
    /// Initializes the slide with all required internal references.
    /// </summary>
    /// <param name="presentation">The parent presentation.</param>
    /// <param name="package">The OPC package.</param>
    /// <param name="partName">The part name of this slide.</param>
    /// <param name="slideRef">The slide reference from the presentation part.</param>
    /// <param name="slidePart">The parsed slide part.</param>
    /// <param name="layoutResolver">Optional callable that resolves a part name to a layout slide.</param>
    internal void InitInternal(
        IPresentation presentation,
        OpcPackage package,
        string partName,
        SlideReference slideRef,
        SlidePart slidePart,
        Func<string, ILayoutSlide?>? layoutResolver = null)
    {
        _presentationRef = presentation;
        PackageInternal = package;
        PartName = partName;
        _partName = partName;
        _slideRef = slideRef;
        _slidePart = slidePart;
        _layoutResolver = layoutResolver;
        _layoutSlide = null;
        _notesSlideManager = null;
    }

    /// <inheritdoc/>
    public int SlideNumber
    {
        get => _slideNumber;
        set => _slideNumber = value;
    }

    /// <inheritdoc/>
    public bool Hidden
    {
        get
        {
            var el = _slidePart?.Element;
            if (el is null)
                return _hiddenFallback;
            var show = (string?)el.Attribute("show");
            return show == "0";
        }
        set
        {
            var el = _slidePart?.Element;
            if (el is null)
            {
                _hiddenFallback = value;
                return;
            }
            if (value)
                el.SetAttributeValue("show", "0");
            else
                el.Attribute("show")?.Remove();
        }
    }

    /// <inheritdoc/>
    public override string Name
    {
        get
        {
            var cSld = _slidePart?.Element?.Element(PNs + "cSld");
            if (cSld is null)
                return _nameFallback;
            return (string?)cSld.Attribute("name") ?? string.Empty;
        }
        set
        {
            var cSld = _slidePart?.Element?.Element(PNs + "cSld");
            if (cSld is null)
            {
                _nameFallback = value;
                return;
            }
            if (string.IsNullOrEmpty(value))
                cSld.Attribute("name")?.Remove();
            else
                cSld.SetAttributeValue("name", value);
        }
    }

    /// <inheritdoc/>
    public ILayoutSlide? LayoutSlide
    {
        get => _layoutSlide;
        set => _layoutSlide = value;
    }

    /// <inheritdoc/>
    public INotesSlideManager NotesSlideManager
    {
        get
        {
            if (_notesSlideManager is null)
            {
                _notesSlideManager = new NotesSlideManager();
                _notesSlideManager.InitInternal(this, PackageInternal, _slidePart);
            }

            return _notesSlideManager;
        }
    }

    /// <inheritdoc/>
    public List<IComment> GetSlideComments(ICommentAuthor? author)
    {
        var presentation = Presentation as Presentation;
        if (presentation is null)
            return [];

        var allComments = new List<IComment>();
        foreach (var commentAuthor in presentation.CommentAuthors.ToArray())
        {
            if (author is not null && !ReferenceEquals(commentAuthor, author) &&
                !(commentAuthor.Name == author.Name && commentAuthor.Initials == author.Initials))
            {
                continue;
            }

            foreach (var comment in commentAuthor.Comments.ToArray())
            {
                if (comment.Slide is Slide s &&
                    string.Equals(s.PartName, PartName, StringComparison.OrdinalIgnoreCase))
                {
                    allComments.Add(comment);
                }
            }
        }

        return allComments;
    }

    /// <summary>
    /// Sets the presentation reference for this slide. Internal use only.
    /// </summary>
    internal void SetPresentation(IPresentation? presentation)
    {
        _presentationRef = presentation;
    }

    /// <summary>
    /// Sets the slide part for this slide. Internal use only.
    /// </summary>
    internal void SetSlidePart(SlidePart? slidePart)
    {
        _slidePart = slidePart;
    }

    /// <summary>
    /// Gets the slide part for this slide. Internal use only.
    /// </summary>
    internal SlidePart? GetSlidePartInternal() => _slidePart;

    /// <inheritdoc/>
    public void Remove()
    {
        var presentation = Presentation as Presentation;
        presentation?.SlidesInternal.Remove(this);
    }
}
