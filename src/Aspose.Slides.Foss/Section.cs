namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a section of slides in a presentation.
/// </summary>
/// <remarks>
/// Obtained from <see cref="ISectionCollection.AddSection"/> or
/// <see cref="ISectionCollection.AppendEmptySection"/>, which bind it to the collection that owns
/// it. There is deliberately no public constructor: a section that belongs to no collection lists no
/// slides and is never written to the file, so <c>new Section { Name = "…" }</c> could only produce
/// something that looked like a section and did nothing.
/// </remarks>
public sealed class Section : ISection
{
    private readonly SectionCollection _owner;
    private string _name = string.Empty;

    internal Section(SectionCollection owner, Guid? sectionId = null)
    {
        _owner = owner;
        SectionId = sectionId ?? Guid.NewGuid();
    }

    /// <inheritdoc />
    public string Name
    {
        get => _name;
        set
        {
            if (string.Equals(_name, value, StringComparison.Ordinal))
                return;

            _name = value;
            _owner.MarkDirty();
        }
    }

    /// <inheritdoc />
    public Guid SectionId { get; }

    /// <inheritdoc />
    public ISlide? StartedFromSlide { get; internal set; }

    /// <inheritdoc />
    public IList<ISlide> GetSlidesListOfSection() => _owner.SlidesOf(this);
}
