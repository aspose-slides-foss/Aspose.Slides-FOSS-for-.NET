using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

public sealed class INotesSizeTests
{
    [Fact]
    public void NotesSize_Has_Positive_Width_And_Height()
    {
        using var pres = new Presentation();
        var ns = pres.NotesSize;

        ns.Size.Width.Should().BeGreaterThan(0);
        ns.Size.Height.Should().BeGreaterThan(0);
    }
}
