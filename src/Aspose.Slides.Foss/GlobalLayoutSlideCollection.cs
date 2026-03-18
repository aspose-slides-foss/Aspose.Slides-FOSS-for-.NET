using System.Collections;

namespace Aspose.Slides.Foss;

/// <summary>
/// Aggregates all layout slides across all master slides in a presentation.
/// </summary>
public sealed class GlobalLayoutSlideCollection : LayoutSlideCollection, IGlobalLayoutSlideCollection
{
    private readonly List<ILayoutSlide> _layouts = new();

    /// <summary>
    /// Initializes internal state with a flat list of all layout slides.
    /// </summary>
    internal void InitInternal(List<ILayoutSlide> layouts)
    {
        _layouts.Clear();
        _layouts.AddRange(layouts);
    }

    /// <inheritdoc />
    public override ILayoutSlide this[int index] => _layouts[index];

    /// <summary>
    /// Gets the number of layout slides.
    /// </summary>
    public int Count => _layouts.Count;

    /// <inheritdoc />
    public override ICollection AsICollection => _layouts;

    /// <inheritdoc />
    public override IEnumerable AsIEnumerable => _layouts;

    /// <summary>
    /// Returns this collection as an <see cref="ILayoutSlideCollection"/>.
    /// </summary>
    public ILayoutSlideCollection AsILayoutSlideCollection => this;

    /// <inheritdoc />
    public override ILayoutSlide? GetByType(SlideLayoutType type)
    {
        foreach (var layout in _layouts)
        {
            if (layout.LayoutType == type)
                return layout;
        }
        return null;
    }

    /// <inheritdoc />
    public override IEnumerator<ILayoutSlide> GetEnumerator() => _layouts.GetEnumerator();
}
