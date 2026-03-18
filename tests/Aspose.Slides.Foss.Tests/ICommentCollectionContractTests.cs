using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using Aspose.Slides.Foss.Internal;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Exercises the ICommentCollection contract: AddComment, InsertComment, RemoveAt,
/// Remove, Clear, ToArray, AsICollection, AsIEnumerable.
/// </summary>
public sealed class ICommentCollectionContractTests
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

    private static (ICommentCollection collection, Slide slide) CreateSetup(
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

        return (collection, slide);
    }

    private static (ICommentCollection alice, ICommentCollection bob, Slide slide) CreateTwoAuthorSetup()
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

        return (aliceCollection, bobCollection, slide);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void AddComment_TextAndAuthorPersist()
    {
        var (collection, slide) = CreateSetup();
        var now = new DateTime(2026, 1, 15, 12, 0, 0);

        var comment = collection.AddComment("Review note", slide, new PointF(2.0f, 3.0f), now);

        comment.Text.Should().Be("Review note");
        comment.Author.Name.Should().Be("Alice");
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void AddComment_PersistsInCollection()
    {
        var (collection, slide) = CreateSetup();
        var now = new DateTime(2026, 1, 15, 12, 0, 0);

        collection.AddComment("Review note", slide, new PointF(2.0f, 3.0f), now);

        var all = collection.ToArray();
        all.Should().HaveCount(1);
        all[0].Text.Should().Be("Review note");
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void RemoveAt_RemovesCommentByIndex()
    {
        var (collection, slide) = CreateSetup();
        var now = DateTime.Now;

        collection.AddComment("C1", slide, new PointF(1, 1), now);
        collection.AddComment("C2", slide, new PointF(2, 2), now);
        collection.AddComment("C3", slide, new PointF(3, 3), now);
        collection.ToArray().Should().HaveCount(3);

        collection.RemoveAt(1);

        collection.ToArray().Should().HaveCount(2);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void RemoveAt_RemovesFirstElement()
    {
        var (collection, slide) = CreateSetup();
        var now = DateTime.Now;

        collection.AddComment("First", slide, new PointF(1, 1), now);
        collection.AddComment("Second", slide, new PointF(2, 2), now);

        collection.RemoveAt(0);

        collection.ToArray().Should().HaveCount(1);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void RemoveAt_ReducesCountByOne()
    {
        var (collection, slide) = CreateSetup();
        var now = DateTime.Now;

        collection.AddComment("C1", slide, new PointF(1, 1), now);
        collection.AddComment("C2", slide, new PointF(2, 2), now);
        collection.ToArray().Should().HaveCount(2);

        collection.RemoveAt(1);

        collection.ToArray().Should().HaveCount(1);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void InsertComment_PlacesAtCorrectIndex()
    {
        var (collection, slide) = CreateSetup();
        var now = DateTime.Now;

        collection.AddComment("First", slide, new PointF(1, 1), now);
        collection.AddComment("Third", slide, new PointF(1, 3), now);
        collection.InsertComment(1, "Second", slide, new PointF(1, 2), now);

        var all = collection.ToArray();
        all.Should().HaveCount(3);
        all[1].Text.Should().Be("Second");
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void Clear_RemovesAllComments()
    {
        var (collection, slide) = CreateSetup();
        var now = DateTime.Now;

        collection.AddComment("C1", slide, new PointF(1, 1), now);
        collection.AddComment("C2", slide, new PointF(2, 2), now);

        collection.Clear();

        collection.ToArray().Should().HaveCount(0);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void TwoAuthors_CommentsAreIsolatedByAuthor()
    {
        var (alice, bob, slide) = CreateTwoAuthorSetup();
        var now = DateTime.Now;

        alice.AddComment("Alice's", slide, new PointF(1, 1), now);
        bob.AddComment("Bob's", slide, new PointF(2, 2), now);

        alice.ToArray().Should().HaveCount(1);
        alice.ToArray()[0].Text.Should().Be("Alice's");

        bob.ToArray().Should().HaveCount(1);
        bob.ToArray()[0].Text.Should().Be("Bob's");
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void TwoAuthors_TotalCommentsSumCorrectly()
    {
        var (alice, bob, slide) = CreateTwoAuthorSetup();
        var now = DateTime.Now;

        alice.AddComment("Alice's", slide, new PointF(1, 1), now);
        bob.AddComment("Bob's", slide, new PointF(2, 2), now);

        var totalCount = alice.ToArray().Length + bob.ToArray().Length;
        totalCount.Should().Be(2);
    }

    /// <summary>
    /// Remove a specific comment by reference.
    /// </summary>
    [Fact]
    public void Remove_RemovesSpecificComment()
    {
        var (collection, slide) = CreateSetup();
        var now = DateTime.Now;

        collection.AddComment("C1", slide, new PointF(1, 1), now);
        var toRemove = collection.AddComment("C2", slide, new PointF(2, 2), now);
        collection.AddComment("C3", slide, new PointF(3, 3), now);

        collection.Remove(toRemove);

        var remaining = collection.ToArray();
        remaining.Should().HaveCount(2);
        remaining.Select(c => c.Text).Should().BeEquivalentTo("C1", "C3");
    }

    /// <summary>
    /// AsICollection returns all comments as an IList.
    /// </summary>
    [Fact]
    public void AsICollection_ReturnsAllComments()
    {
        var (collection, slide) = CreateSetup();
        var now = DateTime.Now;

        collection.AddComment("C1", slide, new PointF(1, 1), now);
        collection.AddComment("C2", slide, new PointF(2, 2), now);

        IList<IComment> list = collection.AsICollection;

        list.Should().HaveCount(2);
    }

    /// <summary>
    /// AsIEnumerable returns all comments as an enumerable.
    /// </summary>
    [Fact]
    public void AsIEnumerable_ReturnsAllComments()
    {
        var (collection, slide) = CreateSetup();
        var now = DateTime.Now;

        collection.AddComment("C1", slide, new PointF(1, 1), now);
        collection.AddComment("C2", slide, new PointF(2, 2), now);

        collection.AsIEnumerable.Count().Should().Be(2);
    }

    /// <summary>
    /// ToArray returns all comments preserving order.
    /// </summary>
    [Fact]
    public void ToArray_ReturnsCommentsInOrder()
    {
        var (collection, slide) = CreateSetup();
        var now = DateTime.Now;

        collection.AddComment("First", slide, new PointF(1, 1), now);
        collection.AddComment("Second", slide, new PointF(2, 2), now);

        var array = collection.ToArray();

        array.Should().HaveCount(2);
        array[0].Text.Should().Be("First");
        array[1].Text.Should().Be("Second");
    }
}
