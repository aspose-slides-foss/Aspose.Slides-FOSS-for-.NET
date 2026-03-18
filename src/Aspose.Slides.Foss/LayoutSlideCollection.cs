using System.Collections;

namespace Aspose.Slides.Foss;

/// <summary>
/// Base class for collections of layout slides.
/// </summary>
public abstract class LayoutSlideCollection : ILayoutSlideCollection
{
    /// <inheritdoc />
    public abstract ILayoutSlide this[int index] { get; }

    /// <inheritdoc />
    public abstract ICollection AsICollection { get; }

    /// <inheritdoc />
    public abstract IEnumerable AsIEnumerable { get; }

    /// <inheritdoc />
    public abstract ILayoutSlide? GetByType(SlideLayoutType type);

    /// <inheritdoc />
    public abstract IEnumerator<ILayoutSlide> GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
