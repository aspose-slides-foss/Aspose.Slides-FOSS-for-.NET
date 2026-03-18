using System.Text.RegularExpressions;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Base class for Slide, LayoutSlide, and MasterSlide providing common slide functionality.
/// </summary>
public class BaseSlide : IBaseSlide
{
    private protected SlidePart? _slidePart;
    private protected SlidePart? _layoutPart;
    private protected SlidePart? _masterPart;
    private protected IPresentation? _presentationRef;
    private protected object? _slideRef;
    private protected string? _partName;

    private ShapeCollection? _shapes;

    /// <inheritdoc/>
    public IPresentation? Presentation => _presentationRef;

    /// <inheritdoc/>
    public IShapeCollection? Shapes
    {
        get
        {
            if (_shapes is null)
            {
                var slidePart = GetSlidePart();
                if (slidePart is not null)
                {
                    var collection = new ShapeCollection();
                    collection.InitInternal(slidePart, this);
                    _shapes = collection;
                }
            }

            return _shapes;
        }
    }

    /// <inheritdoc/>
    public virtual string Name { get; set; } = string.Empty;

    /// <inheritdoc/>
    public virtual int SlideId
    {
        get
        {
            if (_partName is not null)
            {
                var fileName = _partName.Contains('/')
                    ? _partName[((_partName.LastIndexOf('/') + 1))..]
                    : _partName;
                var match = Regex.Match(fileName, @"(\d+)");
                if (match.Success)
                    return int.Parse(match.Groups[1].Value);
            }

            return 0;
        }
    }

    /// <summary>
    /// Gets the slide part object for this slide, resolving from the available
    /// part references (_slidePart, _layoutPart, or _masterPart).
    /// Subclasses set the appropriate field so this method returns the correct part.
    /// </summary>
    /// <returns>The active <see cref="SlidePart"/>, or <c>null</c> if none is set.</returns>
    internal SlidePart? GetSlidePart()
    {
        return _slidePart ?? _layoutPart ?? _masterPart;
    }
}
