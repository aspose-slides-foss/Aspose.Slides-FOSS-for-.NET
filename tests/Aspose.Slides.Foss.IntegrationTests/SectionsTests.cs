using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.IntegrationTests;

/// <summary>
/// Tests for SectionCollection.
/// </summary>
public sealed class SectionsTests
{
    [Fact]
    public void Test_add_section()
    {
        // A section starting from a slide of this presentation is accepted and counted.
        using var pres = new Presentation();

        var section = pres.Sections.AddSection("Intro", pres.Slides[0]);

        section.Name.Should().Be("Intro");
        pres.Sections.Count.Should().Be(1);
    }

    [Fact]
    public void Test_add_section_with_foreign_slide_is_refused()
    {
        // A section can only start from a slide of the presentation that owns the collection.
        // Accepting a slide from elsewhere used to produce a section that was written with no
        // slides in it, which reads as an empty section and reports nothing to the caller.
        using var pres = new Presentation();
        using var other = new Presentation();

        var add = () => pres.Sections.AddSection("Intro", other.Slides[0]);

        add.Should().Throw<ArgumentException>();
        pres.Sections.Count.Should().Be(0);
    }
}
