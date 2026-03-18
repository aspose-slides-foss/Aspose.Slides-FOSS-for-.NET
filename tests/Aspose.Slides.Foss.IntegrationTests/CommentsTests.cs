using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using Aspose.Slides.Foss.Export;
using FluentAssertions;

namespace Aspose.Slides.Foss.IntegrationTests;

/// <summary>
/// Tests for Comments: authors, comments CRUD, slide comments.
/// </summary>
public sealed class CommentsTests : IDisposable
{
    private readonly string _tempDir;

    public CommentsTests()
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
    public void TestAddAuthor()
    {
        // Author name and initials persist.
        using var pres = new Presentation();
        var author = pres.CommentAuthors.AddAuthor("Alice", "A");
        author.Name.Should().Be("Alice");
        author.Initials.Should().Be("A");
        pres.CommentAuthors.Count.Should().Be(1);

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        pres2.CommentAuthors.Count.Should().Be(1);
        pres2.CommentAuthors[0].Name.Should().Be("Alice");
        pres2.CommentAuthors[0].Initials.Should().Be("A");
    }

    [Fact]
    public void TestAddComment()
    {
        // Comment text, position, and time persist.
        using var pres = new Presentation();
        var author = pres.CommentAuthors.AddAuthor("Alice", "A");
        var slide = pres.Slides[0];
        var now = new DateTime(2026, 1, 15, 12, 0, 0);
        var comment = author.Comments.AddComment("Review note", slide, new PointF(2.0f, 3.0f), now);
        comment.Text.Should().Be("Review note");
        comment.Author.Name.Should().Be("Alice");

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var a2 = pres2.CommentAuthors[0];
        a2.Comments.Count.Should().Be(1);
        var c = a2.Comments[0];
        c.Text.Should().Be("Review note");
    }

    [Fact]
    public void TestMultipleAuthors()
    {
        // Multiple authors can coexist.
        using var pres = new Presentation();
        pres.CommentAuthors.AddAuthor("Alice", "A");
        pres.CommentAuthors.AddAuthor("Bob", "B");
        pres.CommentAuthors.Count.Should().Be(2);
    }

    [Fact]
    public void TestGetSlideComments()
    {
        // get_slide_comments filters by author.
        using var pres = new Presentation();
        var a1 = pres.CommentAuthors.AddAuthor("Alice", "A");
        var a2 = pres.CommentAuthors.AddAuthor("Bob", "B");
        var slide = pres.Slides[0];
        var now = DateTime.Now;
        a1.Comments.AddComment("Alice's", slide, new PointF(1, 1), now);
        a2.Comments.AddComment("Bob's", slide, new PointF(2, 2), now);

        var allC = slide.GetSlideComments(null);
        allC.Count.Should().Be(2);

        var bobC = slide.GetSlideComments(a2);
        bobC.Count.Should().Be(1);
        bobC[0].Text.Should().Be("Bob's");
    }

    [Fact]
    public void TestRemoveComment()
    {
        // Removing a comment persists.
        using var pres = new Presentation();
        var author = pres.CommentAuthors.AddAuthor("Alice", "A");
        var slide = pres.Slides[0];
        var now = DateTime.Now;
        author.Comments.AddComment("C1", slide, new PointF(1, 1), now);
        author.Comments.AddComment("C2", slide, new PointF(2, 2), now);
        author.Comments.AddComment("C3", slide, new PointF(3, 3), now);
        author.Comments.Count.Should().Be(3);

        author.Comments.RemoveAt(1);
        author.Comments.Count.Should().Be(2);

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        pres2.CommentAuthors[0].Comments.Count.Should().Be(2);
    }

    [Fact]
    public void TestInsertComment()
    {
        // insert_comment places at the correct index.
        using var pres = new Presentation();
        var author = pres.CommentAuthors.AddAuthor("Alice", "A");
        var slide = pres.Slides[0];
        var now = DateTime.Now;
        author.Comments.AddComment("First", slide, new PointF(1, 1), now);
        author.Comments.AddComment("Third", slide, new PointF(1, 3), now);
        author.Comments.InsertComment(1, "Second", slide, new PointF(1, 2), now);
        author.Comments.Count.Should().Be(3);
        author.Comments[1].Text.Should().Be("Second");
    }

    [Fact]
    public void TestClearComments()
    {
        // clear() removes all comments from an author.
        using var pres = new Presentation();
        var author = pres.CommentAuthors.AddAuthor("Alice", "A");
        var slide = pres.Slides[0];
        var now = DateTime.Now;
        author.Comments.AddComment("C1", slide, new PointF(1, 1), now);
        author.Comments.AddComment("C2", slide, new PointF(2, 2), now);
        author.Comments.Clear();
        author.Comments.Count.Should().Be(0);
    }

    [Fact]
    public void TestRemoveAuthor()
    {
        // Removing an author persists.
        using var pres = new Presentation();
        pres.CommentAuthors.AddAuthor("Alice", "A");
        pres.CommentAuthors.AddAuthor("Bob", "B");
        pres.CommentAuthors.Remove(pres.CommentAuthors[0]);
        pres.CommentAuthors.Count.Should().Be(1);

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        pres2.CommentAuthors.Count.Should().Be(1);
        pres2.CommentAuthors[0].Name.Should().Be("Bob");
    }
}
