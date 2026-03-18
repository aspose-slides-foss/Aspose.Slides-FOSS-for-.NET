using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using Aspose.Slides.Foss.Internal;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

public sealed class CommentTests
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";

    private static XElement MakeCommentElem(int authorId, int idx, string text,
        float posX = 0f, float posY = 0f, string? dt = null, int? parentCmId = null)
    {
        var cm = new XElement(PNs + "cm",
            new XAttribute("authorId", authorId),
            new XAttribute("idx", idx));

        if (dt is not null)
            cm.Add(new XAttribute("dt", dt));

        if (parentCmId is not null)
            cm.Add(new XAttribute("parentCmId", parentCmId.Value));

        cm.Add(new XElement(PNs + "pos",
            new XAttribute("x", (int)Math.Round(posX * 360000)),
            new XAttribute("y", (int)Math.Round(posY * 360000))));

        cm.Add(new XElement(PNs + "text", text));

        return cm;
    }

    private static XElement MakeAuthorElem(int id, string name, string initials)
    {
        return new XElement(PNs + "cmAuthor",
            new XAttribute("id", id),
            new XAttribute("name", name),
            new XAttribute("initials", initials));
    }

    private static (Comment comment, CommentsPart commentsPart) CreateComment(
        string text, int authorId, int idx,
        float posX = 1f, float posY = 1f,
        string? dt = null, int? parentCmId = null,
        XElement? commentsRoot = null,
        XElement? authorsRoot = null,
        ISlide? slide = null,
        ICommentAuthor? author = null)
    {
        commentsRoot ??= new XElement(PNs + "cmLst");
        authorsRoot ??= new XElement(PNs + "cmAuthorLst",
            MakeAuthorElem(authorId, "Alice", "A"));

        var cmElem = MakeCommentElem(authorId, idx, text, posX, posY, dt, parentCmId);
        commentsRoot.Add(cmElem);

        var commentsPart = new CommentsPart(commentsRoot);
        var authorsPart = new CommentAuthorsPart(authorsRoot);
        var slideStub = slide ?? new SlideStub();
        author ??= new CommentAuthorImpl(new AuthorData(authorsRoot.Elements(PNs + "cmAuthor").First()));

        var comment = new Comment();
        comment.InitInternal(new CommentData(cmElem), commentsPart, authorsPart, slideStub, author);

        return (comment, commentsPart);
    }

    private sealed class SlideStub : ISlide
    {
        public int SlideNumber { get; set; }
        public bool Hidden { get; set; }
        public ILayoutSlide? LayoutSlide { get; set; }
        public INotesSlideManager NotesSlideManager => throw new InvalidOperationException("Not used in tests.");
        public List<IComment> GetSlideComments(ICommentAuthor? author) => [];
        public void Remove() { }
        public IPresentation? Presentation => null;
        public IShapeCollection? Shapes => null;
        public string Name { get; set; } = "";
        public int SlideId => 0;
    }

    [Fact]
    public void Text_ReturnsCommentText()
    {
        var (comment, _) = CreateComment("Review note", authorId: 0, idx: 1);

        comment.Text.Should().Be("Review note");
    }

    [Fact]
    public void Text_CanBeModified()
    {
        var (comment, _) = CreateComment("Original", authorId: 0, idx: 1);

        comment.Text = "Updated";

        comment.Text.Should().Be("Updated");
    }

    [Fact]
    public void Author_ReturnsCorrectAuthor()
    {
        var authorsRoot = new XElement(PNs + "cmAuthorLst",
            MakeAuthorElem(0, "Alice", "A"));
        var author = new CommentAuthorImpl(new AuthorData(authorsRoot.Elements(PNs + "cmAuthor").First()));

        var (comment, _) = CreateComment("Review note", authorId: 0, idx: 1,
            authorsRoot: authorsRoot, author: author);

        comment.Author.Name.Should().Be("Alice");
        comment.Author.Initials.Should().Be("A");
    }

    [Fact]
    public void Position_ReturnsCorrectCoordinates()
    {
        var (comment, _) = CreateComment("Note", authorId: 0, idx: 1, posX: 2.0f, posY: 3.0f);

        comment.Position.X.Should().BeApproximately(2.0f, 0.01f);
        comment.Position.Y.Should().BeApproximately(3.0f, 0.01f);
    }

    [Fact]
    public void Position_CanBeModified()
    {
        var (comment, _) = CreateComment("Note", authorId: 0, idx: 1, posX: 1.0f, posY: 1.0f);

        comment.Position = new PointF(5.0f, 7.0f);

        comment.Position.X.Should().BeApproximately(5.0f, 0.01f);
        comment.Position.Y.Should().BeApproximately(7.0f, 0.01f);
    }

    [Fact]
    public void CreatedTime_ReturnsSetDateTime()
    {
        var (comment, _) = CreateComment("Note", authorId: 0, idx: 1,
            dt: "2026-01-15T12:00:00");

        comment.CreatedTime.Should().Be(new DateTime(2026, 1, 15, 12, 0, 0));
    }

    [Fact]
    public void CreatedTime_CanBeModified()
    {
        var (comment, _) = CreateComment("Note", authorId: 0, idx: 1);

        var newTime = new DateTime(2026, 6, 1, 9, 30, 0);
        comment.CreatedTime = newTime;

        comment.CreatedTime.Should().Be(newTime);
    }

    [Fact]
    public void Slide_ReturnsAssociatedSlide()
    {
        var slideStub = new SlideStub();

        var (comment, _) = CreateComment("Note", authorId: 0, idx: 1, slide: slideStub);

        comment.Slide.Should().BeSameAs(slideStub);
    }

    [Fact]
    public void ParentComment_IsNullByDefault()
    {
        var (comment, _) = CreateComment("Note", authorId: 0, idx: 1);

        comment.ParentComment.Should().BeNull();
    }

    [Fact]
    public void ParentComment_CanBeSetAndRetrieved()
    {
        var commentsRoot = new XElement(PNs + "cmLst");
        var authorsRoot = new XElement(PNs + "cmAuthorLst",
            MakeAuthorElem(0, "Alice", "A"));

        // Create parent comment
        var parentElem = MakeCommentElem(0, 1, "Parent comment", 1f, 1f);
        commentsRoot.Add(parentElem);

        // Create child comment
        var childElem = MakeCommentElem(0, 2, "Reply", 1f, 1f);
        commentsRoot.Add(childElem);

        var commentsPart = new CommentsPart(commentsRoot);
        var authorsPart = new CommentAuthorsPart(authorsRoot);
        var slide = new SlideStub();
        var author = new CommentAuthorImpl(new AuthorData(authorsRoot.Elements(PNs + "cmAuthor").First()));

        var parent = new Comment();
        parent.InitInternal(new CommentData(parentElem), commentsPart, authorsPart, slide, author);

        var child = new Comment();
        child.InitInternal(new CommentData(childElem), commentsPart, authorsPart, slide, author);

        child.ParentComment = parent;

        child.ParentComment.Should().NotBeNull();
        child.ParentComment!.Text.Should().Be("Parent comment");
    }

    [Fact]
    public void ParentComment_CanBeCleared()
    {
        var commentsRoot = new XElement(PNs + "cmLst");
        var authorsRoot = new XElement(PNs + "cmAuthorLst",
            MakeAuthorElem(0, "Alice", "A"));

        var parentElem = MakeCommentElem(0, 1, "Parent", 1f, 1f);
        commentsRoot.Add(parentElem);

        var childElem = MakeCommentElem(0, 2, "Reply", 1f, 1f, parentCmId: 1);
        commentsRoot.Add(childElem);

        var commentsPart = new CommentsPart(commentsRoot);
        var authorsPart = new CommentAuthorsPart(authorsRoot);
        var slide = new SlideStub();
        var author = new CommentAuthorImpl(new AuthorData(authorsRoot.Elements(PNs + "cmAuthor").First()));

        var child = new Comment();
        child.InitInternal(new CommentData(childElem), commentsPart, authorsPart, slide, author);

        child.ParentComment.Should().NotBeNull();

        child.ParentComment = null;

        child.ParentComment.Should().BeNull();
    }

    [Fact]
    public void Remove_RemovesCommentFromXml()
    {
        var commentsRoot = new XElement(PNs + "cmLst");
        var authorsRoot = new XElement(PNs + "cmAuthorLst",
            MakeAuthorElem(0, "Alice", "A"));

        var cm1Elem = MakeCommentElem(0, 1, "C1", 1f, 1f);
        var cm2Elem = MakeCommentElem(0, 2, "C2", 2f, 2f);
        var cm3Elem = MakeCommentElem(0, 3, "C3", 3f, 3f);
        commentsRoot.Add(cm1Elem, cm2Elem, cm3Elem);

        var commentsPart = new CommentsPart(commentsRoot);
        var authorsPart = new CommentAuthorsPart(authorsRoot);
        var slide = new SlideStub();
        var author = new CommentAuthorImpl(new AuthorData(authorsRoot.Elements(PNs + "cmAuthor").First()));

        var comment2 = new Comment();
        comment2.InitInternal(new CommentData(cm2Elem), commentsPart, authorsPart, slide, author);

        commentsPart.GetComments().Should().HaveCount(3);

        comment2.Remove();

        commentsPart.GetComments().Should().HaveCount(2);
        commentsPart.GetComments().Select(c => c.Text).Should().BeEquivalentTo("C1", "C3");
    }

    [Fact]
    public void Remove_AlsoRemovesDirectReplies()
    {
        var commentsRoot = new XElement(PNs + "cmLst");
        var authorsRoot = new XElement(PNs + "cmAuthorLst",
            MakeAuthorElem(0, "Alice", "A"));

        var parentElem = MakeCommentElem(0, 1, "Parent", 1f, 1f);
        var replyElem = MakeCommentElem(0, 2, "Reply", 2f, 2f, parentCmId: 1);
        var otherElem = MakeCommentElem(0, 3, "Other", 3f, 3f);
        commentsRoot.Add(parentElem, replyElem, otherElem);

        var commentsPart = new CommentsPart(commentsRoot);
        var authorsPart = new CommentAuthorsPart(authorsRoot);
        var slide = new SlideStub();
        var author = new CommentAuthorImpl(new AuthorData(authorsRoot.Elements(PNs + "cmAuthor").First()));

        var parent = new Comment();
        parent.InitInternal(new CommentData(parentElem), commentsPart, authorsPart, slide, author);

        parent.Remove();

        commentsPart.GetComments().Should().HaveCount(1);
        commentsPart.GetComments()[0].Text.Should().Be("Other");
    }

    [Fact]
    public void MultipleComments_SameAuthor_CoexistInXml()
    {
        var commentsRoot = new XElement(PNs + "cmLst");
        var authorsRoot = new XElement(PNs + "cmAuthorLst",
            MakeAuthorElem(0, "Alice", "A"));

        commentsRoot.Add(MakeCommentElem(0, 1, "First", 1f, 1f));
        commentsRoot.Add(MakeCommentElem(0, 2, "Second", 1f, 2f));
        commentsRoot.Add(MakeCommentElem(0, 3, "Third", 1f, 3f));

        var commentsPart = new CommentsPart(commentsRoot);

        commentsPart.GetComments().Should().HaveCount(3);
        commentsPart.GetComments()[1].Text.Should().Be("Second");
    }

    [Fact]
    public void MultipleAuthors_CommentsFilterByAuthor()
    {
        var commentsRoot = new XElement(PNs + "cmLst");
        var authorsRoot = new XElement(PNs + "cmAuthorLst",
            MakeAuthorElem(0, "Alice", "A"),
            MakeAuthorElem(1, "Bob", "B"));

        commentsRoot.Add(MakeCommentElem(0, 1, "Alice's", 1f, 1f));
        commentsRoot.Add(MakeCommentElem(1, 1, "Bob's", 2f, 2f));

        var commentsPart = new CommentsPart(commentsRoot);

        commentsPart.GetComments().Should().HaveCount(2);

        var bobComments = commentsPart.GetCommentsByAuthor(1);
        bobComments.Should().HaveCount(1);
        bobComments[0].Text.Should().Be("Bob's");
    }
}
