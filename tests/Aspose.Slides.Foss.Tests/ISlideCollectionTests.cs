using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using Aspose.Slides.Foss.Export;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests for ISlideCollection operations: add, insert, remove, clone, iterate, index_of.
/// </summary>
public sealed class ISlideCollectionTests
{
    private static byte[] SaveToBytes(Presentation pres)
    {
        using var ms = new MemoryStream();
        pres.Save(ms, SaveFormat.Pptx);
        return ms.ToArray();
    }

    private static Presentation Reload(Presentation pres)
    {
        var bytes = SaveToBytes(pres);
        pres.Dispose();
        return new Presentation(new MemoryStream(bytes));
    }

    // ── test_add_empty_slide ────────────────────

    /// <summary>
    /// add_empty_slide increases slide count.
    /// </summary>
    [Fact]
    public void AddEmptySlide_IncreasesSlideCount()
    {
        using var pres = new Presentation();

        var layout = ((GlobalLayoutSlideCollection)pres.LayoutSlides)[0];
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

        var layout = ((GlobalLayoutSlideCollection)pres.LayoutSlides)[0];
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

        var layout = ((GlobalLayoutSlideCollection)pres.LayoutSlides)[0];
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

        var layout = ((GlobalLayoutSlideCollection)pres.LayoutSlides)[0];
        pres.Slides.AddEmptySlide(layout);

        pres.Slides.RemoveAt(1);

        pres.Slides.Count.Should().Be(1);
    }

    // ── test_clone_slide ────────────────────────

    /// <summary>
    /// add_clone duplicates a slide with its shapes.
    /// </summary>
    [Fact]
    public void AddClone_DuplicatesSlideWithShapes()
    {
        using var pres = new Presentation();

        var slide = pres.Slides[0];
        slide.Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);

        pres.Slides.AddClone(slide);

        pres.Slides.Count.Should().Be(2);
        pres.Slides[1].Shapes.Should().NotBeNull();
    }

    // ── test_iterate_slides ─────────────────────

    /// <summary>
    /// Slides are iterable.
    /// </summary>
    [Fact]
    public void Slides_AreIterable()
    {
        using var pres = new Presentation();

        var layout = ((GlobalLayoutSlideCollection)pres.LayoutSlides)[0];
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

        var layout = ((GlobalLayoutSlideCollection)pres.LayoutSlides)[0];
        pres.Slides.AddEmptySlide(layout);

        pres.Slides.IndexOf(pres.Slides[0]).Should().Be(0);
        pres.Slides.IndexOf(pres.Slides[1]).Should().Be(1);
    }

    // ── test_slide_count_after_add ────────

    /// <summary>
    /// Adding a slide increases slide count to 2.
    /// </summary>
    [Fact]
    public void SlideCountAfterAdd_IsTwo()
    {
        using var pres = new Presentation();

        var layout = ((GlobalLayoutSlideCollection)pres.LayoutSlides)[0];
        pres.Slides.AddEmptySlide(layout);

        pres.Slides.Count.Should().Be(2);
    }

    // ── test_remove_at ──────────────────────────

    /// <summary>
    /// remove_at removes shape by index from a slide's shape collection.
    /// </summary>
    [Fact]
    public void ShapeRemoveAt_RemovesByIndex()
    {
        using var pres = new Presentation();

        var slide = pres.Slides[0];
        slide.Shapes!.Clear();
        slide.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        slide.Shapes.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);

        slide.Shapes.RemoveAt(0);

        slide.Shapes.ToArray().Should().HaveCount(1);
    }

    // ── test_remove_comment ───────────────────

    /// <summary>
    /// Removing a comment persists.
    /// </summary>
    [Fact]
    public void RemoveComment_Persists()
    {
        using var pres = new Presentation();

        var author = pres.CommentAuthors.AddAuthor("Alice", "A");
        var slide = pres.Slides[0];
        var now = DateTime.Now;

        author.Comments.AddComment("C1", slide, new PointF(1, 1), now);
        author.Comments.AddComment("C2", slide, new PointF(2, 2), now);
        author.Comments.AddComment("C3", slide, new PointF(3, 3), now);
        author.Comments.ToArray().Should().HaveCount(3);

        author.Comments.RemoveAt(1);

        author.Comments.ToArray().Should().HaveCount(2);

        using var pres2 = Reload(pres);
        var authors2 = (CommentAuthorCollection)pres2.CommentAuthors;
        authors2[0].Comments.ToArray().Should().HaveCount(2);
    }

    // ── Additional contract tests ────────────────────────────────

    /// <summary>
    /// A new presentation has exactly 1 slide.
    /// </summary>
    [Fact]
    public void NewPresentation_HasOneSlide()
    {
        using var pres = new Presentation();

        pres.Slides.Count.Should().Be(1);
    }

    /// <summary>
    /// AsICollection returns the slides as a list.
    /// </summary>
    [Fact]
    public void AsICollection_ReturnsSlidesAsList()
    {
        using var pres = new Presentation();

        pres.Slides.AsICollection.Should().NotBeNull();
        pres.Slides.AsICollection.Count.Should().Be(1);
    }

    /// <summary>
    /// AsIEnumerable returns the slides as an enumerable.
    /// </summary>
    [Fact]
    public void AsIEnumerable_ReturnsSlidesAsEnumerable()
    {
        using var pres = new Presentation();

        pres.Slides.AsIEnumerable.Should().NotBeNull();
        pres.Slides.AsIEnumerable.Count().Should().Be(1);
    }

    /// <summary>
    /// ToArray returns all slides as an array.
    /// </summary>
    [Fact]
    public void ToArray_ReturnsAllSlides()
    {
        using var pres = new Presentation();

        var layout = ((GlobalLayoutSlideCollection)pres.LayoutSlides)[0];
        pres.Slides.AddEmptySlide(layout);

        var array = pres.Slides.ToArray();

        array.Should().HaveCount(2);
    }

    /// <summary>
    /// ToArray with range returns subset of slides.
    /// </summary>
    [Fact]
    public void ToArray_WithRange_ReturnsSubset()
    {
        using var pres = new Presentation();

        var layout = ((GlobalLayoutSlideCollection)pres.LayoutSlides)[0];
        pres.Slides.AddEmptySlide(layout);
        pres.Slides.AddEmptySlide(layout);

        var subset = pres.Slides.ToArray(1, 2);

        subset.Should().HaveCount(2);
    }

    /// <summary>
    /// IndexOf returns -1 for a slide not in the collection.
    /// </summary>
    [Fact]
    public void IndexOf_ReturnsNegativeOneForMissing()
    {
        using var pres = new Presentation();

        var orphan = new Slide();

        pres.Slides.IndexOf(orphan).Should().Be(-1);
    }
}
