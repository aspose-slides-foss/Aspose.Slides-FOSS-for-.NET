using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Export;
using FluentAssertions;

namespace Aspose.Slides.Foss.IntegrationTests;

/// <summary>
/// Tests for Presentation create / load / save / properties.
/// </summary>
public sealed class PresentationTests : IDisposable
{
    private readonly string _tempDir;

    public PresentationTests()
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
    public void TestCreateEmpty()
    {
        // A brand-new presentation has exactly 1 slide.
        using var pres = new Presentation();
        pres.Slides.Count.Should().Be(1);
    }

    [Fact]
    public void TestSaveAndReload()
    {
        // Round-trip: create → save → reload preserves slide count.
        using var pres = new Presentation();
        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        pres2.Slides.Count.Should().Be(1);
    }

    [Fact]
    public void TestSaveToStream()
    {
        // Saving to a MemoryStream produces a non-empty buffer.
        using var pres = new Presentation();
        using var buf = new MemoryStream();
        pres.Save(buf, SaveFormat.Pptx);
        buf.Position.Should().BeGreaterThan(0);
    }

    [Fact]
    public void TestContextManager()
    {
        // Presentation can be used as a context manager (IDisposable / using).
        using var pres = new Presentation();
        pres.Slides.Count.Should().BeGreaterThanOrEqualTo(1);
    }

    [Fact]
    public void TestFirstSlideNumber()
    {
        // first_slide_number persists across save/reload.
        using var pres = new Presentation();
        pres.FirstSlideNumber = 5;
        pres.FirstSlideNumber.Should().Be(5);

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        pres2.FirstSlideNumber.Should().Be(5);
    }

    [Fact]
    public void TestLoadExisting()
    {
        // Load a known .pptx from test_data and verify it opens.
        var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test_data", "Presentation.pptx");

        using var pres = new Presentation(path);
        pres.Slides.Count.Should().BeGreaterThanOrEqualTo(1);
    }

    [Fact]
    public void TestDisposeIsIdempotent()
    {
        // Calling Dispose() twice must not raise.
        using var pres = new Presentation();
        pres.Dispose();
        pres.Dispose(); // second call should be harmless
    }

    [Fact]
    public void TestSlideCountAfterAdd()
    {
        // Adding a slide increases slide count to 2.
        using var pres = new Presentation();
        var layout = pres.LayoutSlides[0];
        pres.Slides.AddEmptySlide(layout);
        pres.Slides.Count.Should().Be(2);
    }
}
