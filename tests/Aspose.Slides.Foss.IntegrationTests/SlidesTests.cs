using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.IntegrationTests;

/// <summary>
/// Tests for SlideCollection operations and Slide properties.
/// </summary>
public sealed class SlidesTests : IDisposable
{
    private readonly string _tempDir;

    public SlidesTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, recursive: true);
    }

    [Fact]
    public void TestAddEmptySlide()
    {
        // add_empty_slide increases slide count.
        using var pres = new Presentation();
        var layout = pres.LayoutSlides[0];
        pres.Slides.AddEmptySlide(layout);
        pres.Slides.Count.Should().Be(2);
    }

    [Fact]
    public void TestInsertEmptySlide()
    {
        // insert_empty_slide places a slide at the given index.
        using var pres = new Presentation();
        var layout = pres.LayoutSlides[0];
        pres.Slides.AddEmptySlide(layout);
        pres.Slides.InsertEmptySlide(1, layout);
        pres.Slides.Count.Should().Be(3);
    }

    [Fact]
    public void TestRemoveSlideByRef()
    {
        // Removing a slide by reference decreases count.
        using var pres = new Presentation();
        var layout = pres.LayoutSlides[0];
        pres.Slides.AddEmptySlide(layout);
        pres.Slides.Count.Should().Be(2);
        pres.Slides.Remove(pres.Slides[1]);
        pres.Slides.Count.Should().Be(1);
    }

    [Fact]
    public void TestRemoveSlideAt()
    {
        // remove_at removes by index.
        using var pres = new Presentation();
        var layout = pres.LayoutSlides[0];
        pres.Slides.AddEmptySlide(layout);
        pres.Slides.RemoveAt(1);
        pres.Slides.Count.Should().Be(1);
    }

    [Fact]
    public void TestSlideHidden()
    {
        // Setting hidden persists across save/reload.
        using var pres = new Presentation();
        pres.Slides[0].Hidden = true;
        pres.Slides[0].Hidden.Should().BeTrue();

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        pres2.Slides[0].Hidden.Should().BeTrue();
    }

    [Fact]
    public void TestCloneSlide()
    {
        // add_clone duplicates a slide with its shapes.
        using var pres = new Presentation();
        var slide = pres.Slides[0];
        slide.Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        pres.Slides.AddClone(slide);
        pres.Slides.Count.Should().Be(2);
        pres.Slides[1].Shapes!.Count.Should().BeGreaterThanOrEqualTo(1);
    }

    [Fact]
    public void TestSlideLayoutAccess()
    {
        // Each slide exposes its layout_slide.
        using var pres = new Presentation();
        pres.Slides[0].LayoutSlide.Should().NotBeNull();
    }

    [Fact]
    public void TestSlideName()
    {
        // Slide name persists after save/reload.
        using var pres = new Presentation();
        pres.Slides[0].Name = "MySlide";
        pres.Slides[0].Name.Should().Be("MySlide");

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        pres2.Slides[0].Name.Should().Be("MySlide");
    }

    [Fact]
    public void TestIterateSlides()
    {
        // Slides are iterable.
        using var pres = new Presentation();
        var layout = pres.LayoutSlides[0];
        pres.Slides.AddEmptySlide(layout);
        var slides = pres.Slides.ToList();
        slides.Count.Should().Be(2);
    }

    [Fact]
    public void TestIndexOf()
    {
        // index_of returns the correct position.
        using var pres = new Presentation();
        var layout = pres.LayoutSlides[0];
        pres.Slides.AddEmptySlide(layout);
        pres.Slides.IndexOf(pres.Slides[0]).Should().Be(0);
        pres.Slides.IndexOf(pres.Slides[1]).Should().Be(1);
    }
}
