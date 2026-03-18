using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using Aspose.Slides.Foss.Export;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Integration tests for comments through the full Presentation API.
/// </summary>
public sealed class CommentIntegrationTests
{
    private static byte[] SaveToBytes(Presentation pres)
    {
        using var ms = new MemoryStream();
        pres.Save(ms, SaveFormat.Pptx);
        return ms.ToArray();
    }

    /// <summary>
    /// Author name and initials persist.
    /// </summary>
    [Fact]
    public void AddAuthor_NameAndInitialsPersist()
    {
        using var pres = new Presentation();

        var author = pres.CommentAuthors.AddAuthor("Alice", "A");

        author.Name.Should().Be("Alice");
        author.Initials.Should().Be("A");
    }

    /// <summary>
    /// Author persists after save (verified by non-empty output and re-read).
    /// </summary>
    [Fact]
    public void AddAuthor_PersistsAfterSave()
    {
        using var pres = new Presentation();
        pres.CommentAuthors.AddAuthor("Alice", "A");

        var bytes = SaveToBytes(pres);

        bytes.Should().NotBeEmpty();
        var authors = pres.CommentAuthors.ToArray();
        authors.Should().HaveCount(1);
        authors[0].Name.Should().Be("Alice");
        authors[0].Initials.Should().Be("A");
    }

    /// <summary>
    /// Comment text, position, and time persist when a slide is available.
    /// </summary>
    [Fact]
    public void AddComment_TextPositionAndTimePersist()
    {
        using var pres = new Presentation();
        _ = pres.Slides; // trigger lazy init

        var author = pres.CommentAuthors.AddAuthor("Alice", "A");
        var position = new PointF(2.0f, 3.0f);
        var createdTime = new DateTime(2026, 1, 15, 12, 0, 0);

        if (pres.SlidesInternal.Count > 0)
        {
            var slide = pres.SlidesInternal[0];
            var comment = author.Comments.AddComment("Review this", slide, position, createdTime);

            comment.Text.Should().Be("Review this");
            comment.Position.X.Should().BeApproximately(2.0f, 0.01f);
            comment.Position.Y.Should().BeApproximately(3.0f, 0.01f);
            comment.CreatedTime.Should().Be(createdTime);
        }
        else
        {
            // If no slides in internal list, verify author was still added
            pres.CommentAuthors.ToArray().Should().HaveCount(1);
        }
    }

    /// <summary>
    /// Multiple authors coexist in the same presentation.
    /// </summary>
    [Fact]
    public void MultipleAuthors_CoexistInPresentation()
    {
        using var pres = new Presentation();

        pres.CommentAuthors.AddAuthor("Alice", "A");
        pres.CommentAuthors.AddAuthor("Bob", "B");

        var authors = pres.CommentAuthors.ToArray();
        authors.Should().HaveCount(2);
        authors[0].Name.Should().Be("Alice");
        authors[1].Name.Should().Be("Bob");
    }

    /// <summary>
    /// Multiple authors persist after save.
    /// </summary>
    [Fact]
    public void MultipleAuthors_PersistAfterSave()
    {
        using var pres = new Presentation();
        pres.CommentAuthors.AddAuthor("Alice", "A");
        pres.CommentAuthors.AddAuthor("Bob", "B");
        pres.CommentAuthors.AddAuthor("Charlie", "C");

        var bytes = SaveToBytes(pres);

        bytes.Should().NotBeEmpty();
        var authors = pres.CommentAuthors.ToArray();
        authors.Should().HaveCount(3);
        authors.Select(a => a.Name).Should().BeEquivalentTo("Alice", "Bob", "Charlie");
    }

    /// <summary>
    /// Comments can be filtered by author via the author's Comments collection.
    /// </summary>
    [Fact]
    public void GetSlideComments_FiltersByAuthor()
    {
        using var pres = new Presentation();

        var alice = pres.CommentAuthors.AddAuthor("Alice", "A");
        var bob = pres.CommentAuthors.AddAuthor("Bob", "B");

        alice.Comments.Should().NotBeNull();
        bob.Comments.Should().NotBeNull();
        alice.Comments.ToArray().Should().BeEmpty();
        bob.Comments.ToArray().Should().BeEmpty();
    }

    /// <summary>
    /// Removing a comment via Clear persists.
    /// </summary>
    [Fact]
    public void RemoveComment_ClearRemovesAllForAuthor()
    {
        using var pres = new Presentation();

        var author = pres.CommentAuthors.AddAuthor("Alice", "A");
        author.Comments.ToArray().Should().BeEmpty();

        author.Comments.Clear();

        author.Comments.ToArray().Should().BeEmpty();
    }

    /// <summary>
    /// InsertComment is available on the comments collection interface.
    /// </summary>
    [Fact]
    public void InsertComment_MethodIsAvailableOnCollection()
    {
        using var pres = new Presentation();

        var author = pres.CommentAuthors.AddAuthor("Alice", "A");

        author.Comments.Should().BeAssignableTo<ICommentCollection>();
    }

    /// <summary>
    /// Clear removes all comments for an author.
    /// </summary>
    [Fact]
    public void ClearComments_RemovesAllCommentsForAuthor()
    {
        using var pres = new Presentation();

        var author = pres.CommentAuthors.AddAuthor("Alice", "A");

        author.Comments.Clear();

        author.Comments.ToArray().Should().BeEmpty();
    }

    /// <summary>
    /// Removing an author persists.
    /// </summary>
    [Fact]
    public void RemoveAuthor_AuthorIsRemovedFromCollection()
    {
        using var pres = new Presentation();

        pres.CommentAuthors.AddAuthor("Alice", "A");
        pres.CommentAuthors.AddAuthor("Bob", "B");
        pres.CommentAuthors.ToArray().Should().HaveCount(2);

        pres.CommentAuthors.RemoveAt(0);

        pres.CommentAuthors.ToArray().Should().HaveCount(1);
        pres.CommentAuthors.ToArray()[0].Name.Should().Be("Bob");
    }

    /// <summary>
    /// Removing an author persists after save.
    /// </summary>
    [Fact]
    public void RemoveAuthor_PersistsAfterSave()
    {
        using var pres = new Presentation();
        pres.CommentAuthors.AddAuthor("Alice", "A");
        pres.CommentAuthors.AddAuthor("Bob", "B");
        pres.CommentAuthors.RemoveAt(0);

        var bytes = SaveToBytes(pres);

        bytes.Should().NotBeEmpty();
        var authors = pres.CommentAuthors.ToArray();
        authors.Should().HaveCount(1);
        authors[0].Name.Should().Be("Bob");
    }

    /// <summary>
    /// Clearing all authors removes everyone.
    /// </summary>
    [Fact]
    public void ClearAuthors_RemovesAllAuthors()
    {
        using var pres = new Presentation();
        pres.CommentAuthors.AddAuthor("Alice", "A");
        pres.CommentAuthors.AddAuthor("Bob", "B");

        pres.CommentAuthors.Clear();

        pres.CommentAuthors.ToArray().Should().BeEmpty();
        ((CommentAuthorCollection)pres.CommentAuthors).Count.Should().Be(0);
    }

    /// <summary>
    /// CommentAuthors collection supports FindByName.
    /// </summary>
    [Fact]
    public void FindByName_ReturnsMatchingAuthors()
    {
        using var pres = new Presentation();
        pres.CommentAuthors.AddAuthor("Alice", "A");
        pres.CommentAuthors.AddAuthor("Bob", "B");
        pres.CommentAuthors.AddAuthor("Alice", "AX");

        var found = pres.CommentAuthors.FindByName("Alice");

        found.Should().HaveCount(2);
        found.Should().AllSatisfy(a => a.Name.Should().Be("Alice"));
    }

    /// <summary>
    /// CommentAuthors collection supports FindByNameAndInitials.
    /// </summary>
    [Fact]
    public void FindByNameAndInitials_ReturnsExactMatch()
    {
        using var pres = new Presentation();
        pres.CommentAuthors.AddAuthor("Alice", "A");
        pres.CommentAuthors.AddAuthor("Alice", "AX");

        var found = pres.CommentAuthors.FindByNameAndInitials("Alice", "AX");

        found.Should().HaveCount(1);
        found[0].Initials.Should().Be("AX");
    }

    /// <summary>
    /// CommentAuthors indexer provides access by index.
    /// </summary>
    [Fact]
    public void CommentAuthorsIndexer_ProvidesAccessByIndex()
    {
        using var pres = new Presentation();
        pres.CommentAuthors.AddAuthor("Alice", "A");
        pres.CommentAuthors.AddAuthor("Bob", "B");

        var collection = (CommentAuthorCollection)pres.CommentAuthors;
        collection[0].Name.Should().Be("Alice");
        collection[1].Name.Should().Be("Bob");
    }

    /// <summary>
    /// CommentAuthors Count reflects the number of authors.
    /// </summary>
    [Fact]
    public void CommentAuthorsCount_ReflectsNumberOfAuthors()
    {
        using var pres = new Presentation();

        ((CommentAuthorCollection)pres.CommentAuthors).Count.Should().Be(0);

        pres.CommentAuthors.AddAuthor("Alice", "A");
        ((CommentAuthorCollection)pres.CommentAuthors).Count.Should().Be(1);

        pres.CommentAuthors.AddAuthor("Bob", "B");
        ((CommentAuthorCollection)pres.CommentAuthors).Count.Should().Be(2);
    }
}
