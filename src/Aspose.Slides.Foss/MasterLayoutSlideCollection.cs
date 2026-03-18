using System.Collections;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a collections of all layout slides of defined master slide.
/// Extends <see cref="LayoutSlideCollection"/> with methods for adding/inserting/removing/cloning/reordering
/// layout slides in context of the individual collections of master's layout slides.
/// </summary>
public sealed class MasterLayoutSlideCollection : LayoutSlideCollection, IMasterLayoutSlideCollection
{
    private List<ILayoutSlide> _layouts = [];

    /// <summary>
    /// Initializes internal state with the layout slides for this master.
    /// </summary>
    internal void InitInternal(List<ILayoutSlide> layouts)
    {
        _layouts = layouts;
    }

    /// <inheritdoc />
    public override ILayoutSlide this[int index] => _layouts[index];

    /// <inheritdoc />
    public override ICollection AsICollection => _layouts;

    /// <inheritdoc />
    public override IEnumerable AsIEnumerable => _layouts;

    /// <inheritdoc />
    public int Count => _layouts.Count;

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
