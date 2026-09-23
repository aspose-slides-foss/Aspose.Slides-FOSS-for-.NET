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

    internal Section(SectionCollection owner, Guid? sectionId = null, string name = "")
    {
        _owner = owner;
        SectionId = sectionId ?? Guid.NewGuid();

        // Set without marking the collection changed: a section read from a file already says
        // what the file says, and only an edit should cause the section list to be rewritten.
        _name = name;
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
