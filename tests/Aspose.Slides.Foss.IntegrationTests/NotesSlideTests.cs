using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.IntegrationTests;

/// <summary>
/// Tests for NotesSlide, NotesSlideManager, header/footer, NotesSize.
/// </summary>
public sealed class NotesSlideTests : IDisposable
{
    private readonly string _tempDir;

    public NotesSlideTests()
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
    public void Test_add_notes()
    {
        // Notes text persists after save/reload.
        using var pres = new Presentation();
        var slide = pres.Slides[0];
        var notes = slide.NotesSlideManager.AddNotesSlide();
        notes.NotesTextFrame.Text = "Speaker notes";

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var ns2 = pres2.Slides[0].NotesSlideManager.NotesSlide;
        ns2.Should().NotBeNull();
        ns2!.NotesTextFrame.Text.Should().Be("Speaker notes");
    }

    [Fact]
    public void Test_remove_notes()
    {
        // Removing notes persists.
        using var pres = new Presentation();
        var mgr = pres.Slides[0].NotesSlideManager;
        mgr.AddNotesSlide();
        mgr.NotesSlide.Should().NotBeNull();

        mgr.RemoveNotesSlide();
        mgr.NotesSlide.Should().BeNull();

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        pres2.Slides[0].NotesSlideManager.NotesSlide.Should().BeNull();
    }

    [Fact]
    public void Test_notes_header_footer()
    {
        // Header/footer visibility persists.
        using var pres = new Presentation();
        var notes = pres.Slides[0].NotesSlideManager.AddNotesSlide();
        notes.NotesTextFrame.Text = "Notes";
        var hfm = notes.HeaderFooterManager;
        hfm.SetFooterVisibility(true);
        hfm.SetFooterText("Confidential");
        hfm.SetSlideNumberVisibility(true);

        hfm.IsFooterVisible.Should().BeTrue();
        hfm.IsSlideNumberVisible.Should().BeTrue();

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var ns2 = pres2.Slides[0].NotesSlideManager.NotesSlide;
        var hfm2 = ns2!.HeaderFooterManager;
        hfm2.IsFooterVisible.Should().BeTrue();
        hfm2.IsSlideNumberVisible.Should().BeTrue();
    }

    [Fact]
    public void Test_notes_parent_slide()
    {
        // Notes slide references its parent slide.
        using var pres = new Presentation();
        var slide = pres.Slides[0];
        var notes = slide.NotesSlideManager.AddNotesSlide();
        notes.ParentSlide.Should().BeSameAs(slide);
    }

    [Fact]
    public void Test_notes_size()
    {
        // Notes size has positive width and height.
        using var pres = new Presentation();
        var ns = pres.NotesSize;
        ns.Size.Width.Should().BeGreaterThan(0);
        ns.Size.Height.Should().BeGreaterThan(0);
    }
}
