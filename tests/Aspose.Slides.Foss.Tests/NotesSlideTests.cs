using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests NotesSlide, NotesSlideManager, and NotesSize.
/// </summary>
public sealed class NotesSlideTests
{
    /// <summary>
    /// NotesSize class can be instantiated.
    /// </summary>
    [Fact]
    public void NotesSize_CanBeInstantiated()
    {
        var ns = new NotesSize();
        ns.Should().NotBeNull();
    }

    /// <summary>
    /// SizeF has positive width and height when constructed with values.
    /// </summary>
    [Fact]
    public void SizeF_PositiveWidthAndHeight()
    {
        var size = new SizeF(720, 540);
        size.Width.Should().BeGreaterThan(0);
        size.Height.Should().BeGreaterThan(0);
    }

    /// <summary>
    /// NotesSlideManager class can be instantiated.
    /// </summary>
    [Fact]
    public void NotesSlideManager_CanBeInstantiated()
    {
        var mgr = new NotesSlideManager();
        mgr.Should().NotBeNull();
    }

    /// <summary>
    /// Presentation class can be instantiated.
    /// </summary>
    [Fact]
    public void Presentation_CanBeInstantiated()
    {
        using var pres = new Presentation();
        pres.Should().NotBeNull();
    }
}
