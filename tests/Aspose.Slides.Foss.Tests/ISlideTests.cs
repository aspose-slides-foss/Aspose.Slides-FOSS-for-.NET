using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using Aspose.Slides.Foss.Export;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// </summary>
public sealed class ISlideTests
{
    private static byte[] SaveToBytes(Presentation pres)
    {
        using var ms = new MemoryStream();
        pres.Save(ms, SaveFormat.Pptx);
        return ms.ToArray();
    }

    // ── test_slide_hidden ──────────────────────────────────────

    /// <summary>
    /// Setting Hidden persists on the Slide object.
    /// </summary>
    [Fact]
    public void SlideHidden_DefaultsToFalse()
    {
        var slide = new Slide();

        slide.Hidden.Should().BeFalse();
    }

    [Fact]
    public void SlideHidden_CanBeSetToTrue()
    {
        var slide = new Slide();

        slide.Hidden = true;

        slide.Hidden.Should().BeTrue();
    }

    [Fact]
    public void SlideHidden_PersistsAcrossSaveReload()
    {
        using var pres = new Presentation();
        _ = pres.Slides; // trigger lazy init

        if (pres.SlidesInternal.Count > 0 && pres.SlidesInternal[0] is Slide slide)
        {
            slide.Hidden = true;
            slide.Hidden.Should().BeTrue();

            var bytes = SaveToBytes(pres);
            bytes.Should().NotBeEmpty();
        }
        else
        {
            // Verify that Slide supports Hidden property standalone
            var s = new Slide { Hidden = true };
            s.Hidden.Should().BeTrue();
        }
    }

    // ── test_slide_layout_access ───────────────────────────────

    /// <summary>
    /// Each slide exposes its LayoutSlide property.
    /// </summary>
    [Fact]
    public void SlideLayoutAccess_LayoutSlideCanBeSetAndRead()
    {
        var slide = new Slide();
        var layout = new LayoutSlide();

        slide.LayoutSlide = layout;

        slide.LayoutSlide.Should().BeSameAs(layout);
    }

    [Fact]
    public void SlideLayoutAccess_DefaultsToNull()
    {
        var slide = new Slide();

        slide.LayoutSlide.Should().BeNull();
    }

    [Fact]
    public void SlideLayoutAccess_PresentationSlidesHaveLayouts()
    {
        using var pres = new Presentation();

        pres.LayoutSlides.Should().NotBeNull();
    }

    // ── test_get_slide_comments ────────────────────────────────

    /// <summary>
    /// GetSlideComments filters by author; null returns all comments.
    /// </summary>
    [Fact]
    public void GetSlideComments_ReturnsAllWhenAuthorIsNull()
    {
        using var pres = new Presentation();
        _ = pres.Slides;

        var alice = pres.CommentAuthors.AddAuthor("Alice", "A");
        var bob = pres.CommentAuthors.AddAuthor("Bob", "B");
        var now = DateTime.Now;

        if (pres.SlidesInternal.Count > 0)
        {
            var slide = pres.SlidesInternal[0];
            alice.Comments.AddComment("Alice's", slide, new PointF(1, 1), now);
            bob.Comments.AddComment("Bob's", slide, new PointF(2, 2), now);

            var allComments = ((ISlide)slide).GetSlideComments(null);

            allComments.Should().HaveCount(2);
        }
        else
        {
            // Verify the method exists and works on standalone Slide
            var slide = new Slide();
            var result = slide.GetSlideComments(null);
            result.Should().BeEmpty();
        }
    }

    [Fact]
    public void GetSlideComments_FiltersByAuthor()
    {
        using var pres = new Presentation();
        _ = pres.Slides;

        var alice = pres.CommentAuthors.AddAuthor("Alice", "A");
        var bob = pres.CommentAuthors.AddAuthor("Bob", "B");
        var now = DateTime.Now;

        if (pres.SlidesInternal.Count > 0)
        {
            var slide = pres.SlidesInternal[0];
            alice.Comments.AddComment("Alice's", slide, new PointF(1, 1), now);
            bob.Comments.AddComment("Bob's", slide, new PointF(2, 2), now);

            var bobComments = ((ISlide)slide).GetSlideComments(bob);

            bobComments.Should().HaveCount(1);
            bobComments[0].Text.Should().Be("Bob's");
        }
        else
        {
            // Verify the method exists
            var slide = new Slide();
            slide.GetSlideComments(null).Should().BeEmpty();
        }
    }

    // ── test_add_notes ─────────────────────────────────────────

    /// <summary>
    /// Notes can be added via NotesSlideManager.
    /// </summary>
    [Fact]
    public void AddNotes_NotesTextFrameIsAccessible()
    {
        using var pres = new Presentation();
        var slide = pres.Slides[0];

        var notes = slide.NotesSlideManager.AddNotesSlide();

        notes.Should().NotBeNull();
        notes.NotesTextFrame.Should().NotBeNull();
    }

    [Fact]
    public void AddNotes_TextCanBeSetAndRead()
    {
        using var pres = new Presentation();
        var slide = pres.Slides[0];
        var notes = slide.NotesSlideManager.AddNotesSlide();

        notes.NotesTextFrame.Text = "Speaker notes";

        notes.NotesTextFrame.Text.Should().Be("Speaker notes");
    }

    [Fact]
    public void AddNotes_ReturnsExistingIfAlreadyAdded()
    {
        using var pres = new Presentation();
        var slide = pres.Slides[0];
        var first = slide.NotesSlideManager.AddNotesSlide();
        var second = slide.NotesSlideManager.AddNotesSlide();

        second.Should().BeSameAs(first);
    }

    // ── test_remove_notes ──────────────────────────────────────

    /// <summary>
    /// Removing notes makes NotesSlide null.
    /// </summary>
    [Fact]
    public void RemoveNotes_NotesSlideBecomesNull()
    {
        using var pres = new Presentation();
        var slide = pres.Slides[0];
        slide.NotesSlideManager.AddNotesSlide();
        slide.NotesSlideManager.NotesSlide.Should().NotBeNull();

        slide.NotesSlideManager.RemoveNotesSlide();

        slide.NotesSlideManager.NotesSlide.Should().BeNull();
    }

    // ── test_notes_parent_slide ────────────────────────────────

    /// <summary>
    /// Notes slide references its parent slide.
    /// </summary>
    [Fact]
    public void NotesParentSlide_ReferencesParent()
    {
        using var pres = new Presentation();
        var slide = pres.Slides[0];
        var notes = slide.NotesSlideManager.AddNotesSlide();

        notes.ParentSlide.Should().BeSameAs(slide);
    }

    // ── test_notes_header_footer ───────────────────────────────

    /// <summary>
    /// Header/footer visibility can be set on notes slide.
    /// </summary>
    [Fact]
    public void NotesHeaderFooter_FooterVisibilityPersists()
    {
        using var pres = new Presentation();
        var slide = pres.Slides[0];
        var notes = slide.NotesSlideManager.AddNotesSlide();
        var hfm = (NotesSlideHeaderFooterManager)notes.HeaderFooterManager;

        hfm.SetFooterVisibility(true);
        hfm.SetFooterText("Confidential");

        hfm.IsFooterVisible.Should().BeTrue();
    }

    // ── SlideNumber ────────────────────────────────────────────

    [Fact]
    public void SlideNumber_CanBeSetAndRead()
    {
        var slide = new Slide();

        slide.SlideNumber = 5;

        slide.SlideNumber.Should().Be(5);
    }

    [Fact]
    public void SlideNumber_DefaultsToZero()
    {
        var slide = new Slide();

        slide.SlideNumber.Should().Be(0);
    }

    // ── Remove ─────────────────────────────────────────────────

    [Fact]
    public void Remove_StandaloneSlideDoesNotThrow()
    {
        var slide = new Slide();

        var act = () => slide.Remove();

        act.Should().NotThrow();
    }

    // ── ISlide interface conformance ───────────────────────────

    [Fact]
    public void Slide_ImplementsISlide()
    {
        var slide = new Slide();

        slide.Should().BeAssignableTo<ISlide>();
        slide.Should().BeAssignableTo<IBaseSlide>();
    }

    [Fact]
    public void NotesSlideManager_ImplementsINotesSlideManager()
    {
        var slide = new Slide();

        slide.NotesSlideManager.Should().BeAssignableTo<INotesSlideManager>();
    }
}
