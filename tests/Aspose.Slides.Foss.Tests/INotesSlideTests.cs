using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Internal;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests for INotesSlide interface contract: test_add_notes, test_notes_header_footer,
/// test_notes_parent_slide.
/// </summary>
public sealed class INotesSlideTests
{
    /// <summary>
    /// Creates a fully-initialized <see cref="NotesSlide"/> backed by a fresh
    /// <see cref="NotesSlidePart"/> and the given parent slide.
    /// </summary>
    private static NotesSlide CreateNotesSlide(ISlide parentSlide)
    {
        var notesPart = new NotesSlidePart();
        var package = new OpcPackage();
        using var presentation = new Presentation();
        var notesSlide = new NotesSlide();
        notesSlide.InitInternal(presentation, package, "ppt/notesSlides/notesSlide1.xml", notesPart, parentSlide);
        return notesSlide;
    }

    // ── test_add_notes ──

    /// <summary>
    /// NotesSlide implements INotesSlide.
    /// </summary>
    [Fact]
    public void NotesSlide_ImplementsINotesSlide()
    {
        var notesSlide = CreateNotesSlide(new Slide());

        notesSlide.Should().BeAssignableTo<INotesSlide>();
    }

    /// <summary>
    /// NotesSlide inherits from IBaseSlide through INotesSlide.
    /// </summary>
    [Fact]
    public void NotesSlide_ImplementsIBaseSlide()
    {
        var notesSlide = CreateNotesSlide(new Slide());

        notesSlide.Should().BeAssignableTo<IBaseSlide>();
    }

    /// <summary>
    /// NotesTextFrame is accessible after initialization.
    /// </summary>
    [Fact]
    public void NotesSlide_NotesTextFrameIsAccessible()
    {
        var notesSlide = CreateNotesSlide(new Slide());
        var tf = new TextFrame();
        notesSlide.SetNotesTextFrame(tf);

        INotesSlide iNotesSlide = notesSlide;

        iNotesSlide.NotesTextFrame.Should().NotBeNull();
        iNotesSlide.NotesTextFrame.Should().BeSameAs(tf);
    }

    // ── test_notes_header_footer ──

    /// <summary>
    /// HeaderFooterManager is accessible via INotesSlide and supports visibility settings.
    /// </summary>
    [Fact]
    public void NotesSlide_HeaderFooterManagerIsAccessible()
    {
        INotesSlide notesSlide = CreateNotesSlide(new Slide());

        notesSlide.HeaderFooterManager.Should().NotBeNull();
        notesSlide.HeaderFooterManager.Should().BeAssignableTo<INotesSlideHeaderFooterManager>();
    }

    /// <summary>
    /// Footer and slide number visibility persist through the INotesSlide interface.
    /// </summary>
    [Fact]
    public void NotesSlide_FooterAndSlideNumberVisibilityPersist()
    {
        INotesSlide notesSlide = CreateNotesSlide(new Slide());
        var hfm = (NotesSlideHeaderFooterManager)notesSlide.HeaderFooterManager;

        hfm.SetFooterVisibility(true);
        hfm.SetFooterText("Confidential");
        hfm.SetSlideNumberVisibility(true);

        hfm.IsFooterVisible.Should().BeTrue();
        hfm.IsSlideNumberVisible.Should().BeTrue();
    }

    /// <summary>
    /// Multiple header/footer visibility settings coexist through INotesSlide.
    /// </summary>
    [Fact]
    public void NotesSlide_MultipleVisibilitySettingsCoexist()
    {
        INotesSlide notesSlide = CreateNotesSlide(new Slide());
        var hfm = (NotesSlideHeaderFooterManager)notesSlide.HeaderFooterManager;

        hfm.SetFooterVisibility(true);
        hfm.SetFooterText("Confidential");
        hfm.SetHeaderVisibility(true);

        hfm.IsFooterVisible.Should().BeTrue();
        hfm.IsHeaderVisible.Should().BeTrue();
    }

    // ── test_notes_parent_slide ──

    /// <summary>
    /// Notes slide references its parent slide.
    /// </summary>
    [Fact]
    public void NotesSlide_ParentSlideReferencesOwner()
    {
        var slide = new Slide();
        INotesSlide notesSlide = CreateNotesSlide(slide);

        notesSlide.ParentSlide.Should().BeSameAs(slide);
    }

    /// <summary>
    /// AsIBaseSlide returns the notes slide itself.
    /// </summary>
    [Fact]
    public void NotesSlide_AsIBaseSlideReturnsSelf()
    {
        var notesSlide = CreateNotesSlide(new Slide());

        notesSlide.AsIBaseSlide.Should().BeSameAs(notesSlide);
    }
}
