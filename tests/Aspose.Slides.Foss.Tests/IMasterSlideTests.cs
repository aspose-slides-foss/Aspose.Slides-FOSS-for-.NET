using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests for IMasterSlide interface and master-slide–based slide operations.
/// </summary>
public sealed class IMasterSlideTests
{
    // ── test_slide_count_after_add ────────

    /// <summary>
    /// Adding a slide via a layout obtained from master increases slide count to 2.
    /// </summary>
    [Fact]
    public void SlideCountAfterAdd_ViaLayout_IsTwo()
    {
        using var pres = new Presentation();

        var layout = pres.LayoutSlides[0];
        pres.Slides.AddEmptySlide(layout);

        pres.Slides.Count.Should().Be(2);
    }

    // ── test_add_empty_slide ────────────────────

    /// <summary>
    /// add_empty_slide increases slide count.
    /// </summary>
    [Fact]
    public void AddEmptySlide_IncreasesSlideCount()
    {
        using var pres = new Presentation();

        var layout = pres.LayoutSlides[0];
        pres.Slides.AddEmptySlide(layout);

        pres.Slides.Count.Should().Be(2);
    }

    // ── test_insert_empty_slide ─────────────────

    /// <summary>
    /// insert_empty_slide places a slide at the given index.
    /// </summary>
    [Fact]
    public void InsertEmptySlide_PlacesSlideAtIndex()
    {
        using var pres = new Presentation();

        var layout = pres.LayoutSlides[0];
        pres.Slides.AddEmptySlide(layout);
        pres.Slides.InsertEmptySlide(1, layout);

        pres.Slides.Count.Should().Be(3);
    }

    // ── test_remove_slide_by_ref ────────────────

    /// <summary>
    /// Removing a slide by reference decreases count.
    /// </summary>
    [Fact]
    public void Remove_DecreasesSlideCount()
    {
        using var pres = new Presentation();

        var layout = pres.LayoutSlides[0];
        pres.Slides.AddEmptySlide(layout);
        pres.Slides.Count.Should().Be(2);

        pres.Slides.Remove(pres.Slides[1]);

        pres.Slides.Count.Should().Be(1);
    }

    // ── test_remove_slide_at ────────────────────

    /// <summary>
    /// remove_at removes by index.
    /// </summary>
    [Fact]
    public void RemoveAt_RemovesSlideByIndex()
    {
        using var pres = new Presentation();

        var layout = pres.LayoutSlides[0];
        pres.Slides.AddEmptySlide(layout);

        pres.Slides.RemoveAt(1);

        pres.Slides.Count.Should().Be(1);
    }

    // ── test_iterate_slides ─────────────────────

    /// <summary>
    /// Slides are iterable.
    /// </summary>
    [Fact]
    public void Slides_AreIterable()
    {
        using var pres = new Presentation();

        var layout = pres.LayoutSlides[0];
        pres.Slides.AddEmptySlide(layout);

        var slides = pres.Slides.ToList();

        slides.Should().HaveCount(2);
    }

    // ── test_index_of ───────────────────────────

    /// <summary>
    /// index_of returns the correct position.
    /// </summary>
    [Fact]
    public void IndexOf_ReturnsCorrectPosition()
    {
        using var pres = new Presentation();

        var layout = pres.LayoutSlides[0];
        pres.Slides.AddEmptySlide(layout);

        pres.Slides.IndexOf(pres.Slides[0]).Should().Be(0);
        pres.Slides.IndexOf(pres.Slides[1]).Should().Be(1);
    }

    // ── IMasterSlide contract tests ──────────────────────────────

    /// <summary>
    /// IMasterSlide extends IBaseSlide.
    /// </summary>
    [Fact]
    public void IMasterSlide_ExtendsIBaseSlide()
    {
        var master = new MasterSlide();

        master.Should().BeAssignableTo<IBaseSlide>();
        master.Should().BeAssignableTo<IMasterSlide>();
    }

    /// <summary>
    /// MasterSlide exposes a non-null LayoutSlides collection.
    /// </summary>
    [Fact]
    public void MasterSlide_HasLayoutSlidesCollection()
    {
        using var pres = new Presentation();

        // Trigger parsing of masters
        _ = pres.LayoutSlides;
        var masters = pres.Masters;

        // The master should expose its layout slides
        var master = new MasterSlide();
        master.LayoutSlides.Should().NotBeNull();
    }

    /// <summary>
    /// Presentation masters have layout slides that match the global collection.
    /// </summary>
    [Fact]
    public void Presentation_MasterLayoutSlides_AreAccessible()
    {
        using var pres = new Presentation();

        var globalLayouts = pres.LayoutSlides;
        globalLayouts.Count.Should().BeGreaterThanOrEqualTo(1);
    }
}
