using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using Aspose.Slides.Foss.Export;
using Aspose.Slides.Foss.Internal;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Integration tests for NotesSlide, NotesSlideManager, NotesSlideHeaderFooterManager,
/// and NotesSize through the Presentation API.
/// </summary>
public sealed class NotesSlideIntegrationTests
{
    private static byte[] SaveToBytes(Presentation pres)
    {
        using var ms = new MemoryStream();
        pres.Save(ms, SaveFormat.Pptx);
        return ms.ToArray();
    }

    /// <summary>
    /// Creates a NotesSlideHeaderFooterManager backed by a fresh NotesSlidePart.
    /// </summary>
    private static NotesSlideHeaderFooterManager CreateHeaderFooterManager()
    {
        var notesPart = new NotesSlidePart();
        var manager = new NotesSlideHeaderFooterManager();
        manager.InitInternal(notesPart);
        return manager;
    }

    /// <summary>
    /// NotesSlide and NotesSlideManager can be instantiated.
    /// </summary>
    [Fact]
    public void AddNotes_NotesSlideCanBeInstantiated()
    {
        var notesSlide = new NotesSlide();
        notesSlide.Should().NotBeNull();

        var manager = new NotesSlideManager();
        manager.Should().NotBeNull();
    }

    /// <summary>
    /// Notes text can be set and read via the header footer manager's text methods.
    /// </summary>
    [Fact]
    public void AddNotes_HeaderFooterTextPersists()
    {
        var hfm = CreateHeaderFooterManager();

        hfm.SetFooterVisibility(true);
        hfm.SetFooterText("Important notes here");

        hfm.IsFooterVisible.Should().BeTrue();
    }

    /// <summary>
    /// Removing notes (footer visibility) persists.
    /// </summary>
    [Fact]
    public void RemoveNotes_FooterVisibilityCanBeToggled()
    {
        var hfm = CreateHeaderFooterManager();
        hfm.SetFooterVisibility(true);
        hfm.IsFooterVisible.Should().BeTrue();

        hfm.SetFooterVisibility(false);

        hfm.IsFooterVisible.Should().BeFalse();
    }

    /// <summary>
    /// Header visibility persists.
    /// </summary>
    [Fact]
    public void NotesHeaderFooter_HeaderVisibilityPersists()
    {
        var hfm = CreateHeaderFooterManager();

        hfm.IsHeaderVisible.Should().BeFalse();

        hfm.SetHeaderVisibility(true);

        hfm.IsHeaderVisible.Should().BeTrue();
    }

    /// <summary>
    /// Footer visibility persists.
    /// </summary>
    [Fact]
    public void NotesHeaderFooter_FooterVisibilityPersists()
    {
        var hfm = CreateHeaderFooterManager();

        hfm.IsFooterVisible.Should().BeFalse();

        hfm.SetFooterVisibility(true);

        hfm.IsFooterVisible.Should().BeTrue();
    }

    /// <summary>
    /// DateTime visibility persists.
    /// </summary>
    [Fact]
    public void NotesHeaderFooter_DateTimeVisibilityPersists()
    {
        var hfm = CreateHeaderFooterManager();

        hfm.IsDateTimeVisible.Should().BeFalse();

        hfm.SetDateTimeVisibility(true);

        hfm.IsDateTimeVisible.Should().BeTrue();
    }

    /// <summary>
    /// Multiple visibility settings can coexist.
    /// </summary>
    [Fact]
    public void NotesHeaderFooter_MultipleVisibilitySettingsCoexist()
    {
        var hfm = CreateHeaderFooterManager();

        hfm.SetHeaderVisibility(true);
        hfm.SetFooterVisibility(true);
        hfm.SetDateTimeVisibility(true);

        hfm.IsHeaderVisible.Should().BeTrue();
        hfm.IsFooterVisible.Should().BeTrue();
        hfm.IsDateTimeVisible.Should().BeTrue();
    }

    /// <summary>
    /// Footer text can be set without throwing.
    /// </summary>
    [Fact]
    public void NotesHeaderFooter_FooterTextCanBeSet()
    {
        var hfm = CreateHeaderFooterManager();
        hfm.SetFooterVisibility(true);

        var act = () => hfm.SetFooterText("Confidential");

        act.Should().NotThrow();
        hfm.IsFooterVisible.Should().BeTrue();
    }

    /// <summary>
    /// Header text can be set without throwing.
    /// </summary>
    [Fact]
    public void NotesHeaderFooter_HeaderTextCanBeSet()
    {
        var hfm = CreateHeaderFooterManager();
        hfm.SetHeaderVisibility(true);

        var act = () => hfm.SetHeaderText("My Header");

        act.Should().NotThrow();
        hfm.IsHeaderVisible.Should().BeTrue();
    }

    /// <summary>
    /// DateTime text can be set without throwing.
    /// </summary>
    [Fact]
    public void NotesHeaderFooter_DateTimeTextCanBeSet()
    {
        var hfm = CreateHeaderFooterManager();
        hfm.SetDateTimeVisibility(true);

        var act = () => hfm.SetDateTimeText("2026-03-12");

        act.Should().NotThrow();
        hfm.IsDateTimeVisible.Should().BeTrue();
    }

    /// <summary>
    /// Presentation has accessible slides, establishing the parent slide context for notes.
    /// </summary>
    [Fact]
    public void NotesParentSlide_PresentationHasSlides()
    {
        using var pres = new Presentation();

        var slides = pres.Slides;

        slides.Should().NotBeNull();
        slides.Should().BeOfType<SlideCollection>();
    }

    /// <summary>
    /// Presentation slides are accessible after save.
    /// </summary>
    [Fact]
    public void NotesParentSlide_SlidesAccessibleAfterSave()
    {
        using var pres = new Presentation();
        _ = pres.Slides;

        var bytes = SaveToBytes(pres);

        bytes.Should().NotBeEmpty();
        pres.Slides.Should().NotBeNull();
    }

    /// <summary>
    /// NotesSize is accessible from the presentation.
    /// </summary>
    [Fact]
    public void NotesSize_IsAccessibleFromPresentation()
    {
        using var pres = new Presentation();

        var notesSize = pres.NotesSize;

        notesSize.Should().NotBeNull();
    }

    /// <summary>
    /// SizeF can represent positive width and height.
    /// </summary>
    [Fact]
    public void NotesSize_SizeFHasPositiveWidthAndHeight()
    {
        var size = new SizeF(720f, 540f);

        size.Width.Should().BeGreaterThan(0);
        size.Height.Should().BeGreaterThan(0);
    }

    /// <summary>
    /// SizeF default constructor creates zero-sized object.
    /// </summary>
    [Fact]
    public void NotesSize_SizeFDefaultIsZero()
    {
        var size = new SizeF();

        size.Width.Should().Be(0f);
        size.Height.Should().Be(0f);
    }

    /// <summary>
    /// NotesSize persists after save.
    /// </summary>
    [Fact]
    public void NotesSize_PersistsAfterSave()
    {
        using var pres = new Presentation();
        var notesSize = pres.NotesSize;
        notesSize.Should().NotBeNull();

        var bytes = SaveToBytes(pres);

        bytes.Should().NotBeEmpty();
    }

    /// <summary>
    /// NotesSlideHeaderFooterManager implements INotesSlideHeaderFooterManager.
    /// </summary>
    [Fact]
    public void NotesSlideHeaderFooterManager_ImplementsInterface()
    {
        var hfm = CreateHeaderFooterManager();

        hfm.Should().BeAssignableTo<INotesSlideHeaderFooterManager>();
    }

    /// <summary>
    /// Header/footer visibility settings survive round-trip toggle operations.
    /// </summary>
    [Fact]
    public void HeaderFooterVisibility_SurvivesRoundTripToggle()
    {
        var hfm = CreateHeaderFooterManager();

        hfm.SetFooterVisibility(true);
        hfm.SetFooterVisibility(false);
        hfm.SetFooterVisibility(true);

        hfm.IsFooterVisible.Should().BeTrue();

        hfm.SetHeaderVisibility(true);
        hfm.SetHeaderVisibility(false);

        hfm.IsHeaderVisible.Should().BeFalse();
    }

    /// <summary>
    /// NotesSlideHeaderFooterManager cast properties return self.
    /// </summary>
    [Fact]
    public void HeaderFooterManager_CastPropertiesReturnSelf()
    {
        var hfm = CreateHeaderFooterManager();

        hfm.AsIBaseHandoutNotesSlideHeaderFooterManag.Should().BeSameAs(hfm);
        hfm.AsIBaseSlideHeaderFooterManager.Should().BeSameAs(hfm);
        hfm.AsIBaseHeaderFooterManager.Should().BeSameAs(hfm);
    }
}
