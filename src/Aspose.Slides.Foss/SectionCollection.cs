using System.Collections;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a collection of sections in a presentation.
/// </summary>
public sealed class SectionCollection : ISectionCollection
{
    private readonly List<ISection> _sections = [];

    /// <inheritdoc />
    public ISection this[int index] => _sections[index];

    /// <inheritdoc />
    public int Count => _sections.Count;

    /// <inheritdoc />
    public ISection AddSection(string name, ISlide startedFromSlide)
    {
        var section = new Section
        {
            Name = name,
            StartedFromSlide = startedFromSlide
        };
        _sections.Add(section);
        return section;
    }

    /// <inheritdoc />
    public ISection AppendEmptySection(string name, int startedFromIndex)
    {
        var section = new Section { Name = name };
        _sections.Add(section);
        return section;
    }

    /// <inheritdoc />
    public int IndexOf(ISection section) => _sections.IndexOf(section);

    /// <inheritdoc />
    public void RemoveSection(ISection section) => _sections.Remove(section);

    /// <inheritdoc />
    public void RemoveSectionWithSlides(ISection section) => _sections.Remove(section);

    /// <inheritdoc />
    public void ReorderSectionWithSlides(ISection section, int index)
    {
        _sections.Remove(section);
        if (index >= _sections.Count)
            _sections.Add(section);
        else
            _sections.Insert(index, section);
    }

    /// <inheritdoc />
    public void Clear() => _sections.Clear();

    /// <inheritdoc />
    public IEnumerator<ISection> GetEnumerator() => _sections.GetEnumerator();

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
