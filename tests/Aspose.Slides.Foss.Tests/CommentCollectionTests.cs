using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using Aspose.Slides.Foss.Internal;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests CommentCollection operations: AddComment, InsertComment, RemoveAt, Remove, Clear,
/// ToArray, FindCommentByIdx, AsICollection, AsIEnumerable.
/// </summary>
public sealed class CommentCollectionTests
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

    /// <summary>
    /// Creates a CommentCollection for a single author with a slide and comments part wired up.
    /// </summary>
    private static (CommentCollection collection, Slide slide, CommentsPart commentsPart, CommentAuthorsPart authorsPart) CreateSetup(
        string authorName = "Alice", string authorInitials = "A", int authorId = 0)
    {
        var authorsRoot = new XElement(PNs + "cmAuthorLst",
            MakeAuthorElem(authorId, authorName, authorInitials));
        var authorsPart = new CommentAuthorsPart(authorsRoot);

        var package = new OpcPackage();
        package.SetPart(SlidePartName, Array.Empty<byte>());

        var commentsPart = CommentsPart.CreateEmpty();
        authorsPart.CpCache[CommentsPartName] = commentsPart;

        var authorData = authorsPart.FindAuthorById(authorId)!;

        var collection = new CommentCollection();
        collection.InitInternal(authorData, authorsPart, package, null);

        var slide = new Slide();
        slide.PartName = SlidePartName;

        return (collection, slide, commentsPart, authorsPart);
    }

    /// <summary>
    /// Creates a setup with two authors sharing the same comments part.
    /// </summary>
    private static (CommentCollection alice, CommentCollection bob, Slide slide, CommentsPart commentsPart) CreateTwoAuthorSetup()
    {
        var authorsRoot = new XElement(PNs + "cmAuthorLst",
            MakeAuthorElem(0, "Alice", "A"),
            MakeAuthorElem(1, "Bob", "B"));
        var authorsPart = new CommentAuthorsPart(authorsRoot);

        var package = new OpcPackage();
        package.SetPart(SlidePartName, Array.Empty<byte>());

        var commentsPart = CommentsPart.CreateEmpty();
        authorsPart.CpCache[CommentsPartName] = commentsPart;

        var aliceData = authorsPart.FindAuthorById(0)!;
        var bobData = authorsPart.FindAuthorById(1)!;

        var aliceCollection = new CommentCollection();
        aliceCollection.InitInternal(aliceData, authorsPart, package, null);

        var bobCollection = new CommentCollection();
        bobCollection.InitInternal(bobData, authorsPart, package, null);

        var slide = new Slide();
        slide.PartName = SlidePartName;

        return (aliceCollection, bobCollection, slide, commentsPart);
    }

    [Fact]
    public void AddComment_ReturnsCommentWithCorrectText()
    {
        var (collection, slide, _, _) = CreateSetup();
        var now = new DateTime(2026, 1, 15, 12, 0, 0);

        var comment = collection.AddComment("Review note", slide, new PointF(2.0f, 3.0f), now);

        comment.Text.Should().Be("Review note");
    }

    [Fact]
    public void AddComment_ReturnsCommentWithCorrectAuthor()
    {
        var (collection, slide, _, _) = CreateSetup();
        var now = new DateTime(2026, 1, 15, 12, 0, 0);

        var comment = collection.AddComment("Review note", slide, new PointF(2.0f, 3.0f), now);

        comment.Author.Name.Should().Be("Alice");
    }

    [Fact]
    public void AddComment_PersistsInToArray()
    {
        var (collection, slide, _, _) = CreateSetup();
        var now = new DateTime(2026, 1, 15, 12, 0, 0);

        collection.AddComment("Review note", slide, new PointF(2.0f, 3.0f), now);

        var all = collection.ToArray();
        all.Should().HaveCount(1);
        all[0].Text.Should().Be("Review note");
    }

    [Fact]
    public void AddComment_MultipleCommentsPersist()
    {
        var (collection, slide, _, _) = CreateSetup();
        var now = DateTime.Now;

        collection.AddComment("C1", slide, new PointF(1, 1), now);
        collection.AddComment("C2", slide, new PointF(2, 2), now);
        collection.AddComment("C3", slide, new PointF(3, 3), now);

        collection.ToArray().Should().HaveCount(3);
    }

    [Fact]
    public void RemoveAt_RemovesCommentByIndex()
    {
        var (collection, slide, _, _) = CreateSetup();
        var now = DateTime.Now;

        collection.AddComment("C1", slide, new PointF(1, 1), now);
        collection.AddComment("C2", slide, new PointF(2, 2), now);
        collection.AddComment("C3", slide, new PointF(3, 3), now);
        collection.ToArray().Should().HaveCount(3);

        collection.RemoveAt(1);

        collection.ToArray().Should().HaveCount(2);
    }

    [Fact]
    public void InsertComment_PlacesAtCorrectIndex()
    {
        var (collection, slide, _, _) = CreateSetup();
        var now = DateTime.Now;

        collection.AddComment("First", slide, new PointF(1, 1), now);
        collection.AddComment("Third", slide, new PointF(1, 3), now);
        collection.InsertComment(1, "Second", slide, new PointF(1, 2), now);

        var all = collection.ToArray();
        all.Should().HaveCount(3);
        all[1].Text.Should().Be("Second");
    }

    [Fact]
    public void Clear_RemovesAllCommentsFromAuthor()
    {
        var (collection, slide, _, _) = CreateSetup();
        var now = DateTime.Now;

        collection.AddComment("C1", slide, new PointF(1, 1), now);
        collection.AddComment("C2", slide, new PointF(2, 2), now);

        collection.Clear();

        collection.ToArray().Should().HaveCount(0);
    }

    [Fact]
    public void ToArray_ReturnsAllComments()
    {
        var (collection, slide, _, _) = CreateSetup();
        var now = DateTime.Now;

        collection.AddComment("First", slide, new PointF(1, 1), now);
        collection.AddComment("Second", slide, new PointF(2, 2), now);

        var array = collection.ToArray();

        array.Should().HaveCount(2);
        array[0].Text.Should().Be("First");
        array[1].Text.Should().Be("Second");
    }

    [Fact]
    public void ToArray_WithStartAndCount_ReturnsSubset()
    {
        var (collection, slide, _, _) = CreateSetup();
        var now = DateTime.Now;

        collection.AddComment("A", slide, new PointF(1, 1), now);
        collection.AddComment("B", slide, new PointF(2, 2), now);
        collection.AddComment("C", slide, new PointF(3, 3), now);

        var subset = collection.ToArray(1, 1);

        subset.Should().HaveCount(1);
        subset[0].Text.Should().Be("B");
    }

    [Fact]
    public void FindCommentByIdx_ReturnsMatchingComment()
    {
        var (collection, slide, _, _) = CreateSetup();
        var now = DateTime.Now;

        var added = collection.AddComment("Target", slide, new PointF(1, 1), now);

        // The idx is assigned internally by NextCommentIdx (starts at 1)
        var found = collection.FindCommentByIdx(1);

        found.Should().NotBeNull();
        found!.Text.Should().Be("Target");
    }

    [Fact]
    public void FindCommentByIdx_ReturnsNullForNonexistent()
    {
        var (collection, slide, _, _) = CreateSetup();

        var found = collection.FindCommentByIdx(999);

        found.Should().BeNull();
    }

    [Fact]
    public void AsICollection_ReturnsAllCommentsAsList()
    {
        var (collection, slide, _, _) = CreateSetup();
        var now = DateTime.Now;

        collection.AddComment("C1", slide, new PointF(1, 1), now);
        collection.AddComment("C2", slide, new PointF(2, 2), now);

        IList<IComment> list = collection.AsICollection;

        list.Should().HaveCount(2);
    }

    [Fact]
    public void AsIEnumerable_ReturnsAllCommentsAsEnumerable()
    {
        var (collection, slide, _, _) = CreateSetup();
        var now = DateTime.Now;

        collection.AddComment("C1", slide, new PointF(1, 1), now);
        collection.AddComment("C2", slide, new PointF(2, 2), now);

        collection.AsIEnumerable.Count().Should().Be(2);
    }

    [Fact]
    public void Remove_RemovesSpecificComment()
    {
        var (collection, slide, _, _) = CreateSetup();
        var now = DateTime.Now;

        collection.AddComment("C1", slide, new PointF(1, 1), now);
        var toRemove = collection.AddComment("C2", slide, new PointF(2, 2), now);
        collection.AddComment("C3", slide, new PointF(3, 3), now);

        collection.Remove(toRemove);

        var remaining = collection.ToArray();
        remaining.Should().HaveCount(2);
        remaining.Select(c => c.Text).Should().BeEquivalentTo("C1", "C3");
    }

    [Fact]
    public void Clear_OnlyAffectsOwnAuthor()
    {
        var (alice, bob, slide, commentsPart) = CreateTwoAuthorSetup();
        var now = DateTime.Now;

        alice.AddComment("Alice's", slide, new PointF(1, 1), now);
        bob.AddComment("Bob's", slide, new PointF(2, 2), now);

        alice.Clear();

        alice.ToArray().Should().HaveCount(0);
        bob.ToArray().Should().HaveCount(1);
        bob.ToArray()[0].Text.Should().Be("Bob's");
    }
}
