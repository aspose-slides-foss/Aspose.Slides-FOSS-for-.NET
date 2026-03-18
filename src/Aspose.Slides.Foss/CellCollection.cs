using System.Collections;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a read-only collection of table cells associated with a parent slide and slide part.
/// </summary>
public class CellCollection : ISlideComponent, ICellCollection, IEnumerable<ICell>
{
    private List<Cell>? _cells;
    private SlidePart? _slidePart;
    private IBaseSlide? _parentSlide;

    /// <summary>
    /// Initializes the internal state of this collection.
    /// </summary>
    /// <param name="cells">The backing list of cells.</param>
    /// <param name="slidePart">The underlying slide part.</param>
    /// <param name="parentSlide">The parent slide that owns this collection.</param>
    /// <returns>This instance for fluent chaining.</returns>
    internal CellCollection InitInternal(List<Cell> cells, SlidePart? slidePart, IBaseSlide? parentSlide)
    {
        _cells = cells;
        _slidePart = slidePart;
        _parentSlide = parentSlide;
        return this;
    }

    /// <inheritdoc/>
    public override IBaseSlide? Slide => _parentSlide;

    /// <inheritdoc/>
    public override IPresentation? Presentation => _parentSlide?.Presentation;

    /// <inheritdoc/>
    public ISlideComponent AsISlideComponent => this;

    /// <inheritdoc/>
    public override IPresentationComponent AsIPresentationComponent => this;

    /// <inheritdoc/>
    public IList<ICell> AsICollection => new List<ICell>(_cells ?? []);

    /// <inheritdoc/>
    public IEnumerable<ICell> AsIEnumerable => Enumerate();

    /// <inheritdoc/>
    public ICell this[int index] => (_cells ?? [])[index];

    /// <inheritdoc/>
    public int Count => _cells?.Count ?? 0;

    /// <inheritdoc/>
    public IEnumerator<ICell> GetEnumerator() =>
        (_cells ?? Enumerable.Empty<Cell>()).GetEnumerator();

    /// <inheritdoc/>
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    private IEnumerable<ICell> Enumerate()
    {
        if (_cells is null) yield break;
        foreach (var cell in _cells)
            yield return cell;
    }
}
