using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Export;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests for INotesSlideManager interface contract: AddNotesSlide, RemoveNotesSlide, NotesSlide property.
/// </summary>
public sealed class INotesSlideManagerTests
{
    private static byte[] SaveToBytes(Presentation pres)
    {
        using var ms = new MemoryStream();
        pres.Save(ms, SaveFormat.Pptx);
        return ms.ToArray();
    }

    // ── test_add_notes ──

    /// <summary>
    /// AddNotesSlide returns a non-null notes slide with accessible text frame.
    /// </summary>
    [Fact]
    public void AddNotesSlide_ReturnsNonNull()
    {
        using var pres = new Presentation();
        var mgr = pres.Slides[0].NotesSlideManager;

        var notes = mgr.AddNotesSlide();

        notes.Should().NotBeNull();
        notes.NotesTextFrame.Should().NotBeNull();
    }

    /// <summary>
    /// Notes text can be set and read back via the manager.
    /// </summary>
    [Fact]
    public void AddNotesSlide_NotesTextCanBeSetAndRead()
    {
        using var pres = new Presentation();
        var slide = pres.Slides[0];
        var notes = slide.NotesSlideManager.AddNotesSlide();

        notes.NotesTextFrame.Text = "Speaker notes";

        var ns = slide.NotesSlideManager.NotesSlide;
        ns.Should().NotBeNull();
        ns!.NotesTextFrame.Text.Should().Be("Speaker notes");
    }

    /// <summary>
    /// Notes text produces non-empty save output.
    /// </summary>
    [Fact]
    public void AddNotesSlide_NotesTextProducesNonEmptySave()
    {
        using var pres = new Presentation();
        var notes = pres.Slides[0].NotesSlideManager.AddNotesSlide();
        notes.NotesTextFrame.Text = "Speaker notes";

        var bytes = SaveToBytes(pres);

        bytes.Should().NotBeEmpty();
    }

    /// <summary>
    /// AddNotesSlide is idempotent — calling it twice returns the same instance.
    /// </summary>
    [Fact]
    public void AddNotesSlide_IsIdempotent()
    {
        using var pres = new Presentation();
        var mgr = pres.Slides[0].NotesSlideManager;

        var first = mgr.AddNotesSlide();
        var second = mgr.AddNotesSlide();

        second.Should().BeSameAs(first);
    }

    // ── test_remove_notes ──

    /// <summary>
    /// Removing notes sets NotesSlide to null.
    /// </summary>
    [Fact]
    public void RemoveNotesSlide_SetsNotesSlideToNull()
    {
        using var pres = new Presentation();
        var mgr = pres.Slides[0].NotesSlideManager;
        mgr.AddNotesSlide();
        mgr.NotesSlide.Should().NotBeNull();

        mgr.RemoveNotesSlide();

        mgr.NotesSlide.Should().BeNull();
    }

    /// <summary>
    /// Removing notes produces non-empty save output (no crash).
    /// </summary>
    [Fact]
    public void RemoveNotesSlide_SaveProducesNonEmptyOutput()
    {
        using var pres = new Presentation();
        var mgr = pres.Slides[0].NotesSlideManager;
        mgr.AddNotesSlide();
        mgr.RemoveNotesSlide();

        var bytes = SaveToBytes(pres);

        bytes.Should().NotBeEmpty();
    }

    /// <summary>
    /// RemoveNotesSlide is safe to call when no notes exist.
    /// </summary>
    [Fact]
    public void RemoveNotesSlide_SafeWhenNoNotesExist()
    {
        using var pres = new Presentation();
        var mgr = pres.Slides[0].NotesSlideManager;

        var act = () => mgr.RemoveNotesSlide();

        act.Should().NotThrow();
        mgr.NotesSlide.Should().BeNull();
    }

    /// <summary>
    /// NotesSlide is null by default when no notes have been added.
    /// </summary>
    [Fact]
    public void NotesSlide_IsNullByDefault()
    {
        using var pres = new Presentation();
        var mgr = pres.Slides[0].NotesSlideManager;

        mgr.NotesSlide.Should().BeNull();
    }

    // ── test_notes_parent_slide ──

    /// <summary>
    /// Notes slide references its parent slide.
    /// </summary>
    [Fact]
    public void AddNotesSlide_ParentSlideReferencesOwner()
    {
        using var pres = new Presentation();
        var slide = pres.Slides[0];
        var notes = slide.NotesSlideManager.AddNotesSlide();

        notes.ParentSlide.Should().BeSameAs(slide);
    }

    // ── test_notes_header_footer ──

    /// <summary>
    /// Footer and slide number visibility can be set via the manager's notes slide.
    /// </summary>
    [Fact]
    public void AddNotesSlide_HeaderFooterVisibilityPersists()
    {
        using var pres = new Presentation();
        var notes = pres.Slides[0].NotesSlideManager.AddNotesSlide();
        notes.NotesTextFrame.Text = "Notes";
        var hfm = (NotesSlideHeaderFooterManager)notes.HeaderFooterManager;

        hfm.SetFooterVisibility(true);
        hfm.SetFooterText("Confidential");
        hfm.SetSlideNumberVisibility(true);

        hfm.IsFooterVisible.Should().BeTrue();
        hfm.IsSlideNumberVisible.Should().BeTrue();
    }

    /// <summary>
    /// Header/footer visibility produces non-empty save output.
    /// </summary>
    [Fact]
    public void AddNotesSlide_HeaderFooterSavesSuccessfully()
    {
        using var pres = new Presentation();
        var notes = pres.Slides[0].NotesSlideManager.AddNotesSlide();
        notes.NotesTextFrame.Text = "Notes";
        var hfm = (NotesSlideHeaderFooterManager)notes.HeaderFooterManager;
        hfm.SetFooterVisibility(true);
        hfm.SetFooterText("Confidential");
        hfm.SetSlideNumberVisibility(true);

        var bytes = SaveToBytes(pres);

        bytes.Should().NotBeEmpty();
    }

    /// <summary>
    /// INotesSlideManager is implemented by NotesSlideManager.
    /// </summary>
    [Fact]
    public void NotesSlideManager_ImplementsInterface()
    {
        using var pres = new Presentation();
        var mgr = pres.Slides[0].NotesSlideManager;

        mgr.Should().BeAssignableTo<INotesSlideManager>();
    }


    private static Presentation Reload(Presentation pres)
    {
        var bytes = SaveToBytes(pres);
        pres.Dispose();
        return new Presentation(new MemoryStream(bytes));
    }

    /// <summary>
    /// Notes text persists after save/reload.
    /// </summary>
    [Fact]
    public void AddNotesSlide_NotesTextPersistsAfterReload()
    {
        using var pres = new Presentation();
        var notes = pres.Slides[0].NotesSlideManager.AddNotesSlide();
        notes.NotesTextFrame.Text = "Speaker notes";

        using var pres2 = Reload(pres);
        var ns2 = pres2.Slides[0].NotesSlideManager.NotesSlide;

        ns2.Should().NotBeNull();
        ns2!.NotesTextFrame.Text.Should().Be("Speaker notes");
    }

    /// <summary>
    /// Removing notes persists after save/reload.
    /// </summary>
    [Fact]
    public void RemoveNotesSlide_PersistsAfterReload()
    {
        using var pres = new Presentation();
        var mgr = pres.Slides[0].NotesSlideManager;
        mgr.AddNotesSlide();
        mgr.NotesSlide.Should().NotBeNull();

        mgr.RemoveNotesSlide();
        mgr.NotesSlide.Should().BeNull();

        using var pres2 = Reload(pres);
        pres2.Slides[0].NotesSlideManager.NotesSlide.Should().BeNull();
    }

    /// <summary>
    /// Header/footer visibility persists after save/reload.
    /// </summary>
    [Fact]
    public void AddNotesSlide_HeaderFooterVisibilityPersistsAfterReload()
    {
        using var pres = new Presentation();
        var notes = pres.Slides[0].NotesSlideManager.AddNotesSlide();
        notes.NotesTextFrame.Text = "Notes";
        var hfm = notes.HeaderFooterManager;
        hfm.SetFooterVisibility(true);
        hfm.SetFooterText("Confidential");
        hfm.SetSlideNumberVisibility(true);

        hfm.IsFooterVisible.Should().BeTrue();
        hfm.IsSlideNumberVisible.Should().BeTrue();

        using var pres2 = Reload(pres);
        var ns2 = pres2.Slides[0].NotesSlideManager.NotesSlide;
        ns2.Should().NotBeNull();
        var hfm2 = ns2!.HeaderFooterManager;
        hfm2.IsFooterVisible.Should().BeTrue();
        hfm2.IsSlideNumberVisible.Should().BeTrue();
    }

    /// <summary>
    /// Notes slide references its parent slide after creation.
    /// </summary>
    [Fact]
    public void AddNotesSlide_ParentSlideIsCorrectSlide()
    {
        using var pres = new Presentation();
        var slide = pres.Slides[0];
        var notes = slide.NotesSlideManager.AddNotesSlide();

        notes.ParentSlide.Should().BeSameAs(slide);
    }
}
