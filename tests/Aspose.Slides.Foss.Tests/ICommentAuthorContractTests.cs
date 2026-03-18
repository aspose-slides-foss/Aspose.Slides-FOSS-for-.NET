using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using Aspose.Slides.Foss.Internal;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Exercises the ICommentAuthor contract: Name, Initials, Comments, Remove.
/// </summary>
public sealed class ICommentAuthorContractTests
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";

    private const string SlidePartName = "ppt/slides/slide1.xml";
    private const string CommentsPartName = "ppt/comments/comment1.xml";

    private static XElement MakeAuthorElem(int id, string name, string initials)
    {
        return new XElement(PNs + "cmAuthor",
            new XAttribute("id", id),
            new XAttribute("name", name),
            new XAttribute("initials", initials),
            new XAttribute("lastIdx", 0),
            new XAttribute("clrIdx", id));
    }

    private static (CommentAuthorCollection authors, Slide slide) CreateSetup(
        params (int id, string name, string initials)[] authorDefs)
    {
        if (authorDefs.Length == 0)
            authorDefs = [(0, "Alice", "A")];

        var authorsRoot = new XElement(PNs + "cmAuthorLst",
            authorDefs.Select(a => MakeAuthorElem(a.id, a.name, a.initials)));
        var authorsPart = new CommentAuthorsPart(authorsRoot);

        var package = new OpcPackage();
        package.SetPart(SlidePartName, Array.Empty<byte>());

        var commentsPart = CommentsPart.CreateEmpty();
        authorsPart.CpCache[CommentsPartName] = commentsPart;

        var collection = new CommentAuthorCollection();
        collection.InitInternal(authorsPart, package);

        var slide = new Slide();
        slide.PartName = SlidePartName;

        return (collection, slide);
    }

    [Fact]
    public void AddAuthor_NameAndInitialsPersist()
    {
        var (authors, _) = CreateSetup();
        // Collection already has Alice from setup; verify via interface
        var author = authors[0];

        author.Name.Should().Be("Alice");
        author.Initials.Should().Be("A");
        authors.Count.Should().Be(1);
    }

    [Fact]
    public void AddComment_TextAndAuthorPersist()
    {
        var (authors, slide) = CreateSetup();
        var author = authors[0];
        var now = new DateTime(2026, 1, 15, 12, 0, 0);

        var comment = author.Comments.AddComment("Review note", slide, new PointF(2.0f, 3.0f), now);

        comment.Text.Should().Be("Review note");
        comment.Author.Name.Should().Be("Alice");
        author.Comments.ToArray().Should().HaveCount(1);
        author.Comments.ToArray()[0].Text.Should().Be("Review note");
    }

    [Fact]
    public void MultipleAuthors_Coexist()
    {
        var (authors, _) = CreateSetup(
            (0, "Alice", "A"),
            (1, "Bob", "B"));

        authors.Count.Should().Be(2);
    }

    [Fact]
    public void Comments_FilteredByAuthor()
    {
        var (authors, slide) = CreateSetup(
            (0, "Alice", "A"),
            (1, "Bob", "B"));
        var alice = authors[0];
        var bob = authors[1];
        var now = DateTime.Now;

        alice.Comments.AddComment("Alice's", slide, new PointF(1, 1), now);
        bob.Comments.AddComment("Bob's", slide, new PointF(2, 2), now);

        alice.Comments.ToArray().Should().HaveCount(1);
        alice.Comments.ToArray()[0].Text.Should().Be("Alice's");

        bob.Comments.ToArray().Should().HaveCount(1);
        bob.Comments.ToArray()[0].Text.Should().Be("Bob's");
    }

    [Fact]
    public void RemoveComment_ByIndex()
    {
        var (authors, slide) = CreateSetup();
        var author = authors[0];
        var now = DateTime.Now;

        author.Comments.AddComment("C1", slide, new PointF(1, 1), now);
        author.Comments.AddComment("C2", slide, new PointF(2, 2), now);
        author.Comments.AddComment("C3", slide, new PointF(3, 3), now);
        author.Comments.ToArray().Should().HaveCount(3);

        author.Comments.RemoveAt(1);

        author.Comments.ToArray().Should().HaveCount(2);
    }

    [Fact]
    public void InsertComment_PlacesAtCorrectIndex()
    {
        var (authors, slide) = CreateSetup();
        var author = authors[0];
        var now = DateTime.Now;

        author.Comments.AddComment("First", slide, new PointF(1, 1), now);
        author.Comments.AddComment("Third", slide, new PointF(1, 3), now);
        author.Comments.InsertComment(1, "Second", slide, new PointF(1, 2), now);

        var all = author.Comments.ToArray();
        all.Should().HaveCount(3);
        all[1].Text.Should().Be("Second");
    }

    [Fact]
    public void ClearComments_RemovesAllFromAuthor()
    {
        var (authors, slide) = CreateSetup();
        var author = authors[0];
        var now = DateTime.Now;

        author.Comments.AddComment("C1", slide, new PointF(1, 1), now);
        author.Comments.AddComment("C2", slide, new PointF(2, 2), now);

        author.Comments.Clear();

        author.Comments.ToArray().Should().HaveCount(0);
    }

    [Fact]
    public void RemoveAuthor_LeavesRemainingAuthors()
    {
        var (authors, _) = CreateSetup(
            (0, "Alice", "A"),
            (1, "Bob", "B"));

        authors.Remove(authors[0]);

        authors.Count.Should().Be(1);
        authors[0].Name.Should().Be("Bob");
    }
}
