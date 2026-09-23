using System.Globalization;
using System.Xml.Linq;
using Aspose.Slides.Foss.ConformanceTests.Harness;
using Aspose.Slides.Foss.Drawing;
using Aspose.Slides.Foss.Export;

namespace Aspose.Slides.Foss.ConformanceTests;

/// <summary>
/// Comments and their replies are written in the form PowerPoint reads, and survive an edit.
/// </summary>
/// <remarks>
/// PowerPoint writes a reply as a classic <c>&lt;p:cm&gt;</c> like any other, and ties it to its
/// parent with the <c>p15:threadingInfo</c> extension in the comment's <c>p:extLst</c>: the reply's
/// holds <c>&lt;p15:parentCm authorId="…" idx="…"/&gt;</c> naming the parent by author and index,
/// and the parent's is empty. Without <c>p15:parentCm</c> PowerPoint shows the reply as a second,
/// unrelated comment. An attribute the schema does not declare is the worst outcome of all: it makes
/// the package invalid, no reader renders the thread, and the API reports success.
/// </remarks>
public sealed class CommentConformanceTests : IDisposable
{
    /// <summary>The <c>p:ext/@uri</c> of the extension that carries <c>p15:threadingInfo</c>.</summary>
    private const string ThreadingExtensionUri = "{C676402C-5697-4E1C-873F-D02D1690AC5C}";

    private const string CommentsRelationshipType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

    private static readonly DateTime When = new(2024, 1, 1, 9, 0, 0, DateTimeKind.Utc);

    private readonly TestWorkspace _workspace = new();

    public void Dispose() => _workspace.Dispose();

    [Fact]
    public void AReplyNamesItsParentTheWayPowerPointDoes()
    {
        var path = WriteDeckWithAReply("comment-reply-threaded.pptx");

        using var package = PptxPackage.Open(path);
        var part = CommentsPartOfSlide(package, 0);
        var parent = CommentWithText(package, part, "Please check this");
        var reply = CommentWithText(package, part, "Checked");

        var parentInfo = ThreadingInfo(package, part, parent);
        Assert.True(!parentInfo.HasElements,
            $"The parent's p15:threadingInfo names a parent of its own: {parentInfo}");

        var replyInfo = ThreadingInfo(package, part, reply);
        var parentCm = replyInfo.Elements(Ns.P15 + "parentCm").ToList();
        Assert.True(parentCm.Count == 1,
            $"The reply's p15:threadingInfo carries {parentCm.Count} p15:parentCm element(s), expected 1." +
            $"{Environment.NewLine}{PackageAssert.Describe(package, part)}");
        Assert.True((string?)parentCm[0].Attribute("authorId") == (string?)parent.Attribute("authorId") &&
                    (string?)parentCm[0].Attribute("idx") == (string?)parent.Attribute("idx"),
            $"The reply names {parentCm[0]}, but its parent is authorId=\"{(string?)parent.Attribute("authorId")}\" " +
            $"idx=\"{(string?)parent.Attribute("idx")}\".{Environment.NewLine}{PackageAssert.Describe(package, part)}");

        foreach (var info in new[] { parentInfo, replyInfo })
        {
            Assert.True(int.TryParse((string?)info.Attribute("timeZoneBias"), NumberStyles.Integer, CultureInfo.InvariantCulture, out _),
                $"p15:threadingInfo has no integer timeZoneBias, which PowerPoint writes on every one: {info}");
        }

        PackageAssert.ChildrenInSchemaOrder(reply, Ns.P + "pos", Ns.P + "text", Ns.P + "extLst");
        Assert.DoesNotContain(package.PartNames, name => name.StartsWith("ppt/threadedComments/", StringComparison.Ordinal));
    }

    [Fact]
    public void AReplyByAnotherAuthorNamesItsParentByAuthorAndIndex()
    {
        var path = _workspace.PathFor("comment-reply-two-authors.pptx");

        using (var presentation = new Presentation())
        {
            var alice = presentation.CommentAuthors.AddAuthor("Alice", "A");
            var bob = presentation.CommentAuthors.AddAuthor("Bob", "B");
            var slide = presentation.Slides[0];
            var parent = alice.Comments.AddComment("Parent", slide, new PointF(1, 1), When);
            bob.Comments.AddComment("Unrelated", slide, new PointF(1, 2), When);
            var reply = bob.Comments.AddComment("Reply", slide, new PointF(1, 3), When);
            reply.ParentComment = parent;
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);
        var part = CommentsPartOfSlide(package, 0);
        var parentElement = CommentWithText(package, part, "Parent");
        var replyElement = CommentWithText(package, part, "Reply");
        Assert.NotEqual((string?)parentElement.Attribute("authorId"), (string?)replyElement.Attribute("authorId"));

        var parentCm = ThreadingInfo(package, part, replyElement).Element(Ns.P15 + "parentCm");
        Assert.True(parentCm is not null &&
                    (string?)parentCm.Attribute("authorId") == (string?)parentElement.Attribute("authorId") &&
                    (string?)parentCm.Attribute("idx") == (string?)parentElement.Attribute("idx"),
            $"Bob's reply must name Alice's comment by her author id and its index." +
            $"{Environment.NewLine}{PackageAssert.Describe(package, part)}");
        Assert.Null(ThreadingInfoOrNull(CommentWithText(package, part, "Unrelated"))?.Element(Ns.P15 + "parentCm"));
    }

    [Fact]
    public void AReplyWhoseParentIsClearedIsWrittenAsATopLevelComment()
    {
        var path = _workspace.PathFor("comment-reply-cleared.pptx");

        using (var presentation = new Presentation())
        {
            var author = presentation.CommentAuthors.AddAuthor("Reviewer", "RV");
            var parent = author.Comments.AddComment("Please check this", presentation.Slides[0], new PointF(10, 10), When);
            var reply = author.Comments.AddComment("Checked", presentation.Slides[0], new PointF(20, 20), When);
            reply.ParentComment = parent;
            reply.ParentComment = null;
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);
        var part = CommentsPartOfSlide(package, 0);
        Assert.True(PackageAssert.Select(package, part, "//p15:parentCm").Count == 0,
            $"A reply whose parent was cleared still names one.{Environment.NewLine}{PackageAssert.Describe(package, part)}");
    }

    [Fact]
    public void RemovingACommentRemovesItsRepliesByEveryAuthorFromTheFile()
    {
        var path = _workspace.PathFor("comment-parent-removed.pptx");

        using (var presentation = new Presentation())
        {
            var alice = presentation.CommentAuthors.AddAuthor("Alice", "A");
            var bob = presentation.CommentAuthors.AddAuthor("Bob", "B");
            var slide = presentation.Slides[0];
            var parent = alice.Comments.AddComment("Parent", slide, new PointF(1, 1), When);
            alice.Comments.AddComment("Kept", slide, new PointF(1, 2), When);
            alice.Comments.AddComment("Reply by Alice", slide, new PointF(1, 3), When).ParentComment = parent;
            bob.Comments.AddComment("Reply by Bob", slide, new PointF(1, 4), When).ParentComment = parent;
            parent.Remove();
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);
        var part = CommentsPartOfSlide(package, 0);
        var texts = PackageAssert.Select(package, part, "/p:cmLst/p:cm/p:text").Select(t => t.Value).ToList();
        Assert.True(texts.SequenceEqual(["Kept"]),
            $"Removing a comment must remove it and every reply to it; the file holds: {string.Join(", ", texts)}." +
            $"{Environment.NewLine}{PackageAssert.Describe(package, part)}");
    }

    [Fact]
    public void ACommentRemovedWithoutItsRepliesLeavesNoReplyNamingIt()
    {
        var path = _workspace.PathFor("comment-parent-removed-at.pptx");

        using (var presentation = new Presentation())
        {
            var author = presentation.CommentAuthors.AddAuthor("Reviewer", "RV");
            var parent = author.Comments.AddComment("Parent", presentation.Slides[0], new PointF(1, 1), When);
            author.Comments.AddComment("Reply", presentation.Slides[0], new PointF(1, 2), When).ParentComment = parent;
            author.Comments.RemoveAt(0);
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);
        var part = CommentsPartOfSlide(package, 0);
        var reply = CommentWithText(package, part, "Reply");
        Assert.True(ThreadingInfoOrNull(reply)?.Element(Ns.P15 + "parentCm") is null,
            $"The reply names a parent that is no longer in the file.{Environment.NewLine}{PackageAssert.Describe(package, part)}");
    }

    [Fact]
    public void RemovingACommentRemovesRepliesToItsRepliesFromTheFile()
    {
        var path = _workspace.PathFor("comment-thread-removed.pptx");

        using (var presentation = new Presentation())
        {
            var author = presentation.CommentAuthors.AddAuthor("Reviewer", "RV");
            var slide = presentation.Slides[0];
            var parent = author.Comments.AddComment("Parent", slide, new PointF(1, 1), When);
            var reply = author.Comments.AddComment("Reply", slide, new PointF(1, 2), When);
            reply.ParentComment = parent;
            author.Comments.AddComment("Reply to the reply", slide, new PointF(1, 3), When).ParentComment = reply;
            author.Comments.AddComment("Kept", slide, new PointF(1, 4), When);
            parent.Remove();
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);
        var part = CommentsPartOfSlide(package, 0);
        var texts = PackageAssert.Select(package, part, "/p:cmLst/p:cm/p:text").Select(t => t.Value).ToList();
        Assert.True(texts.SequenceEqual(["Kept"]),
            $"Removing a comment must remove every reply under it; the file holds: {string.Join(", ", texts)}." +
            $"{Environment.NewLine}{PackageAssert.Describe(package, part)}");
    }

    /// <summary>
    /// A new author takes the highest author id plus one, so removing the author with the highest id
    /// hands that id to the next author added, whose first comment is idx 1 again. A reply that still
    /// named a removed comment would then name the new author's comment.
    /// </summary>
    [Theory]
    [InlineData("ICommentAuthor.Remove")]
    [InlineData("ICommentAuthorCollection.RemoveAt")]
    [InlineData("ICommentAuthorCollection.Remove")]
    public void AReplyWhoseParentsAuthorIsRemovedIsNotAdoptedByTheNextAuthor(string removal)
    {
        var path = _workspace.PathFor($"comment-author-id-reused-{removal}.pptx");

        using (var presentation = new Presentation())
        {
            var alice = presentation.CommentAuthors.AddAuthor("Alice", "A");
            var bob = presentation.CommentAuthors.AddAuthor("Bob", "B");
            var slide = presentation.Slides[0];
            var parent = bob.Comments.AddComment("Bob's comment", slide, new PointF(1, 1), When);
            var reply = alice.Comments.AddComment("Alice's reply", slide, new PointF(1, 2), When);
            reply.ParentComment = parent;

            switch (removal)
            {
                case "ICommentAuthor.Remove": bob.Remove(); break;
                case "ICommentAuthorCollection.RemoveAt": presentation.CommentAuthors.RemoveAt(1); break;
                default: presentation.CommentAuthors.Remove(bob); break;
            }

            var carol = presentation.CommentAuthors.AddAuthor("Carol", "C");
            carol.Comments.AddComment("Carol's unrelated comment", slide, new PointF(1, 3), When);

            var readBack = reply.ParentComment;
            Assert.True(readBack is null,
                $"Bob's comment was removed, and Alice's reply now reads as a reply to '{readBack?.Text}' by '{readBack?.Author.Name}'.");
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);
        var part = CommentsPartOfSlide(package, 0);
        var replyElement = CommentWithText(package, part, "Alice's reply");
        Assert.True(ThreadingInfoOrNull(replyElement)?.Element(Ns.P15 + "parentCm") is null,
            $"Bob's comment was removed, and Alice's reply still names a parent, which is now Carol's comment." +
            $"{Environment.NewLine}{PackageAssert.Describe(package, part)}");
    }

    [Fact]
    public void AReplyNamingACommentMissingFromAnOpenedDeckIsNotAdoptedByTheNextAuthor()
    {
        var source = _workspace.PathFor("comment-dangling-source.pptx");
        var path = _workspace.PathFor("comment-dangling-resaved.pptx");

        using (var presentation = new Presentation())
        {
            var alice = presentation.CommentAuthors.AddAuthor("Alice", "A");
            var bob = presentation.CommentAuthors.AddAuthor("Bob", "B");
            var slide = presentation.Slides[0];
            var parent = bob.Comments.AddComment("Bob's comment", slide, new PointF(1, 1), When);
            alice.Comments.AddComment("Alice's reply", slide, new PointF(1, 2), When).ParentComment = parent;
            presentation.Save(source, SaveFormat.Pptx);
        }

        // Take Bob and his comment out of the file, leaving Alice's reply naming authorId="1" idx="1".
        string commentsPart;
        using (var package = PptxPackage.Open(source))
            commentsPart = CommentsPartOfSlide(package, 0);
        RewriteXmlPart(source, commentsPart, root =>
            root.Elements(Ns.P + "cm").Where(cm => (string?)cm.Attribute("authorId") == "1").Remove());
        RewriteXmlPart(source, "ppt/commentAuthors.xml", root =>
            root.Elements(Ns.P + "cmAuthor").Where(a => (string?)a.Attribute("id") == "1").Remove());

        using (var presentation = new Presentation(source))
        {
            var carol = presentation.CommentAuthors.AddAuthor("Carol", "C");
            carol.Comments.AddComment("Carol's unrelated comment", presentation.Slides[0], new PointF(1, 3), When);

            var reply = presentation.CommentAuthors[0].Comments.ToArray().Single();
            var readBack = reply.ParentComment;
            Assert.True(readBack is null,
                $"Alice's reply named a comment the file does not have, and now reads as a reply to '{readBack?.Text}' by '{readBack?.Author.Name}'.");
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var saved = PptxPackage.Open(path);
        var part = CommentsPartOfSlide(saved, 0);
        Assert.True(ThreadingInfoOrNull(CommentWithText(saved, part, "Alice's reply"))?.Element(Ns.P15 + "parentCm") is null,
            $"Alice's reply named a comment the file does not have, and now names Carol's comment." +
            $"{Environment.NewLine}{PackageAssert.Describe(saved, part)}");
    }

    [Fact]
    public void AReplyPowerPointWroteReadsBackWithItsParent()
    {
        using var presentation = new Presentation(TestFixtures.PowerPointCommentReply);

        var comments = presentation.CommentAuthors[0].Comments.ToArray();
        var parent = comments.Single(c => c.Text == "Parent comment");
        var reply = comments.Single(c => c.Text == "Reply to parent");

        Assert.Null(parent.ParentComment);
        Assert.True(reply.ParentComment?.Text == "Parent comment",
            $"PowerPoint shows 'Reply to parent' as a reply to 'Parent comment'; its ParentComment reads " +
            $"{(reply.ParentComment is null ? "null" : $"'{reply.ParentComment.Text}'")}.");
    }

    [Fact]
    public void AReplyPowerPointWroteByAnotherAuthorReadsBackWithItsParent()
    {
        using var presentation = new Presentation(TestFixtures.PowerPointCommentReplyTwoAuthors);

        var bob = presentation.CommentAuthors.ToArray().Single(a => a.Name == "Bob Ref");
        var reply = bob.Comments.ToArray().Single();

        Assert.True(reply.ParentComment?.Text == "Parent comment" && reply.ParentComment.Author.Name == "Alice Ref",
            $"PowerPoint shows Bob's comment as a reply to Alice's; its ParentComment reads " +
            $"{(reply.ParentComment is null ? "null" : $"'{reply.ParentComment.Text}' by '{reply.ParentComment.Author.Name}'")}.");
    }

    [Fact]
    public void ACommentThreadPowerPointWroteSurvivesALoadAndSaveUnchanged()
    {
        var path = _workspace.PathFor("comment-powerpoint-thread-resaved.pptx");

        using (var presentation = new Presentation(TestFixtures.PowerPointCommentReply))
        {
            // Reading the comments is what loads them into memory, and so into the next save.
            foreach (var comment in presentation.CommentAuthors[0].Comments.ToArray())
            {
                _ = comment.Position;
                _ = comment.ParentComment;
            }
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var original = PptxPackage.Open(TestFixtures.PowerPointCommentReply);
        using var saved = PptxPackage.Open(path);
        foreach (var (originalPart, savedPart) in new[]
                 {
                     (CommentsPartOfSlide(original, 0), CommentsPartOfSlide(saved, 0)),
                     ("ppt/commentAuthors.xml", "ppt/commentAuthors.xml"),
                 })
        {
            Assert.True(XNode.DeepEquals(original.Xml(originalPart).Root, saved.Xml(savedPart).Root),
                $"{savedPart} changed on a load and save that edited nothing." +
                $"{Environment.NewLine}{PackageAssert.Describe(original, originalPart)}" +
                $"{Environment.NewLine}{PackageAssert.Describe(saved, savedPart)}");
        }
    }

    [Fact]
    public void AReplyAddedToAnOpenedDeckNamesThePowerPointComment()
    {
        var path = _workspace.PathFor("comment-powerpoint-thread-extended.pptx");

        using (var presentation = new Presentation(TestFixtures.PowerPointCommentReply))
        {
            var author = presentation.CommentAuthors[0];
            var parent = author.Comments.ToArray().Single(c => c.Text == "Parent comment");
            var reply = author.Comments.AddComment("Second reply", presentation.Slides[0], new PointF(1, 1), When);
            reply.ParentComment = parent;
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);
        var part = CommentsPartOfSlide(package, 0);
        var added = CommentWithText(package, part, "Second reply");
        var parentCm = ThreadingInfo(package, part, added).Element(Ns.P15 + "parentCm");
        Assert.True(parentCm is not null && (string?)parentCm.Attribute("authorId") == "1" && (string?)parentCm.Attribute("idx") == "1",
            $"The new reply must name PowerPoint's comment, authorId=\"1\" idx=\"1\".{Environment.NewLine}{PackageAssert.Describe(package, part)}");

        // The parent's own extension is PowerPoint's, and is kept as PowerPoint wrote it.
        var parentExtensions = CommentWithText(package, part, "Parent comment")
            .Elements(Ns.P + "extLst").Elements(Ns.P + "ext").ToList();
        Assert.True(parentExtensions.Count == 1 &&
                    (string?)parentExtensions[0].Element(Ns.P15 + "threadingInfo")?.Attribute("timeZoneBias") == "-240" &&
                    !parentExtensions[0].Element(Ns.P15 + "threadingInfo")!.HasElements,
            $"The parent's threading extension was not kept as PowerPoint wrote it.{Environment.NewLine}{PackageAssert.Describe(package, part)}");

        var first = CommentWithText(package, part, "Reply to parent");
        var firstParent = ThreadingInfo(package, part, first).Element(Ns.P15 + "parentCm");
        Assert.True(firstParent is not null && (string?)firstParent.Attribute("idx") == "1",
            $"PowerPoint's own reply lost its parent.{Environment.NewLine}{PackageAssert.Describe(package, part)}");
    }

    [Fact]
    public void ReplyingToACommentWritesNoAttributeTheSchemaDoesNotDeclare()
    {
        var path = WriteDeckWithAReply("comment-reply.pptx");

        using var package = PptxPackage.Open(path);

        var commentParts = package.PartNames
            .Where(name => name.StartsWith("ppt/comments/", StringComparison.Ordinal))
            .ToList();
        Assert.True(commentParts.Count > 0,
            $"No comments part was written. Parts: {string.Join(", ", package.PartNames)}");

        var undeclared = new List<string>();
        foreach (var part in commentParts)
        {
            foreach (var element in package.Xml(part).Descendants())
            {
                foreach (var attribute in element.Attributes().Where(a => a.Name.LocalName == "parentCmId"))
                    undeclared.Add($"{part}: <{element.Name.LocalName} {attribute.Name.LocalName}=\"{attribute.Value}\">");
            }
        }

        Assert.True(undeclared.Count == 0,
            $"The comments part carries {undeclared.Count} 'parentCmId' attribute(s), which CT_Comment does not " +
            $"declare. Threading is expressed by the p15:threadingInfo extension, not by an attribute on <p:cm>." +
            $"{Environment.NewLine}{string.Join(Environment.NewLine, undeclared)}");
    }

    [Fact]
    public void ADeckWithACommentReplyProducesASchemaValidPackage()
    {
        var path = WriteDeckWithAReply("comment-reply-valid.pptx");

        SchemaValidation.HasNoSchemaErrors(path);
    }

    [Fact]
    public void RemovingASlideTakesTheCommentsAddedToItWithIt()
    {
        var path = _workspace.PathFor("comment-slide-removed.pptx");

        using (var presentation = new Presentation())
        {
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            var author = presentation.CommentAuthors.AddAuthor("Reviewer", "RV");
            author.Comments.AddComment("Kept on slide 1", presentation.Slides[0], new PointF(10, 10), new DateTime(2024, 1, 1, 9, 0, 0, DateTimeKind.Utc));
            author.Comments.AddComment("Removed with slide 2", presentation.Slides[1], new PointF(10, 10), new DateTime(2024, 1, 1, 9, 0, 0, DateTimeKind.Utc));
            presentation.Slides.RemoveAt(1);
            presentation.Save(path, SaveFormat.Pptx);
        }

        AssertRemovedCommentIsGone(path, kept: "Kept on slide 1", removed: "Removed with slide 2");
    }

    [Fact]
    public void RemovingASlideFromAnOpenedDeckWhoseCommentsWereReadTakesItsCommentsWithIt()
    {
        var source = _workspace.PathFor("comment-source.pptx");
        var path = _workspace.PathFor("comment-opened-slide-removed.pptx");

        using (var presentation = new Presentation())
        {
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            var author = presentation.CommentAuthors.AddAuthor("Reviewer", "RV");
            author.Comments.AddComment("Kept on slide 1", presentation.Slides[0], new PointF(10, 10), new DateTime(2024, 1, 1, 9, 0, 0, DateTimeKind.Utc));
            author.Comments.AddComment("Removed with slide 2", presentation.Slides[1], new PointF(10, 10), new DateTime(2024, 1, 1, 9, 0, 0, DateTimeKind.Utc));
            presentation.Save(source, SaveFormat.Pptx);
        }

        using (var presentation = new Presentation(source))
        {
            // Reading the comments is what loads them; a caller who only looks and then removes a
            // slide has not asked for anything to be kept.
            Assert.Equal(2, presentation.CommentAuthors[0].Comments.Count);
            presentation.Slides.RemoveAt(1);
            presentation.Save(path, SaveFormat.Pptx);
        }

        AssertRemovedCommentIsGone(path, kept: "Kept on slide 1", removed: "Removed with slide 2");
    }

    [Fact]
    public void RemovingTheOnlySlideWithACommentLeavesTheAuthorsInADeclaredPart()
    {
        var path = _workspace.PathFor("comment-only-slide-removed.pptx");

        using (var presentation = new Presentation())
        {
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            var author = presentation.CommentAuthors.AddAuthor("Reviewer", "RV");
            author.Comments.AddComment("Removed with slide 2", presentation.Slides[1], new PointF(10, 10), new DateTime(2024, 1, 1, 9, 0, 0, DateTimeKind.Utc));
            presentation.Slides.RemoveAt(1);
            presentation.Save(path, SaveFormat.Pptx);
        }

        AssertRemovedCommentIsGone(path, kept: null, removed: "Removed with slide 2");
        AssertAuthorsPartIsDeclared(path);
    }

    [Fact]
    public void AnAuthorWithNoCommentsIsWrittenToAPartThePackageDeclares()
    {
        var path = _workspace.PathFor("comment-author-only.pptx");

        using (var presentation = new Presentation())
        {
            presentation.CommentAuthors.AddAuthor("Reviewer", "RV");
            presentation.Save(path, SaveFormat.Pptx);
        }

        AssertAuthorsPartIsDeclared(path);
    }

    [Fact]
    public void ReadingTheCommentAuthorsOfADeckWithoutCommentsAddsNoPart()
    {
        var path = _workspace.PathFor("comment-authors-read.pptx");

        using (var presentation = new Presentation())
        {
            Assert.Equal(0, presentation.CommentAuthors.Count);
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);
        Assert.False(package.Contains("ppt/commentAuthors.xml"),
            $"Only reading the comment authors added a part. Parts: {string.Join(", ", package.PartNames)}");
    }

    /// <summary>
    /// Asserts that the package keeps the comment on the surviving slide and carries no trace of the
    /// removed slide's: no part holding its text, no relationships part for the slide that is gone,
    /// and no content-type declaration for a part that is not there.
    /// </summary>
    private static void AssertRemovedCommentIsGone(string path, string? kept, string removed)
    {
        using var package = PptxPackage.Open(path);

        var holding = package.XmlPartNames.Where(name => package.Text(name).Contains(removed, StringComparison.Ordinal)).ToList();
        Assert.True(holding.Count == 0,
            $"The slide was removed and its comment text is still in the package, in: {string.Join(", ", holding)}" +
            $"{Environment.NewLine}Parts: {string.Join(", ", package.PartNames)}");

        if (kept is not null)
        {
            Assert.True(package.XmlPartNames.Any(name => package.Text(name).Contains(kept, StringComparison.Ordinal)),
                $"The comment on the slide that was kept is missing. Parts: {string.Join(", ", package.PartNames)}");
        }

        PackageAssert.EveryRelationshipsPartBelongsToAPart(package);
        PackageAssert.NoOverrideNamesAMissingPart(package);
        PackageAssert.EveryInternalRelationshipTargetExists(package);
        PackageAssert.EveryPartHasItsRequiredContentType(package);
    }

    /// <summary>
    /// Asserts that <c>ppt/commentAuthors.xml</c> is in the package, has its content type, and is
    /// related from the presentation part — a part without both is one no reader can find.
    /// </summary>
    private static void AssertAuthorsPartIsDeclared(string path)
    {
        using var package = PptxPackage.Open(path);

        Assert.True(package.Contains("ppt/commentAuthors.xml"),
            $"No comment authors part was written. Parts: {string.Join(", ", package.PartNames)}");

        PackageAssert.EveryPartHasItsRequiredContentType(package);

        var related = package.Relationships("ppt/presentation.xml")
            .Any(r => r.Type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors" &&
                      PptxPackage.ResolveTarget("ppt/presentation.xml", r.Target) == "ppt/commentAuthors.xml");
        Assert.True(related,
            $"ppt/commentAuthors.xml is not related from ppt/presentation.xml." +
            $"{Environment.NewLine}{PackageAssert.Describe(package, "ppt/_rels/presentation.xml.rels")}");
    }

    /// <summary>
    /// Replaces an XML part of the package at <paramref name="path"/> with an edited copy, to build a
    /// file this library would not write itself.
    /// </summary>
    private static void RewriteXmlPart(string path, string partName, Action<XElement> edit)
    {
        using var zip = System.IO.Compression.ZipFile.Open(path, System.IO.Compression.ZipArchiveMode.Update);
        var entry = zip.GetEntry(partName)!;
        XDocument doc;
        using (var stream = entry.Open())
            doc = XDocument.Load(stream);
        edit(doc.Root!);
        entry.Delete();
        using (var stream = zip.CreateEntry(partName).Open())
            doc.Save(stream);
    }

    private string WriteDeckWithAReply(string fileName)
    {
        var path = _workspace.PathFor(fileName);

        using var presentation = new Presentation();
        var author = presentation.CommentAuthors.AddAuthor("Reviewer", "RV");
        var parent = author.Comments.AddComment("Please check this", presentation.Slides[0], new PointF(10, 10), new DateTime(2024, 1, 1, 9, 0, 0, DateTimeKind.Utc));
        var reply = author.Comments.AddComment("Checked", presentation.Slides[0], new PointF(20, 20), new DateTime(2024, 1, 1, 10, 0, 0, DateTimeKind.Utc));
        reply.ParentComment = parent;
        presentation.Save(path, SaveFormat.Pptx);

        return path;
    }

    /// <summary>
    /// Returns the comments part of the slide at <paramref name="slideIndex"/> in slide-list order,
    /// found the way a reader finds it: through the slide's relationships.
    /// </summary>
    internal static string CommentsPartOfSlide(PptxPackage package, int slideIndex)
    {
        const string presentationPart = "ppt/presentation.xml";
        var slideIds = PackageAssert.Select(package, presentationPart, "/p:presentation/p:sldIdLst/p:sldId");
        Assert.True(slideIndex < slideIds.Count, $"The deck has {slideIds.Count} slide(s); slide {slideIndex} was asked for.");

        var relationshipId = (string?)slideIds[slideIndex].Attribute(Ns.R + "id");
        var slidePart = PptxPackage.ResolveTarget(presentationPart,
            package.Relationships(presentationPart).Single(r => r.Id == relationshipId).Target);

        var comments = package.Relationships(slidePart).Where(r => r.Type == CommentsRelationshipType).ToList();
        Assert.True(comments.Count == 1,
            $"{slidePart} has {comments.Count} comments relationship(s), expected 1." +
            $"{Environment.NewLine}{PackageAssert.Describe(package, PptxPackage.RelationshipsPartNameFor(slidePart))}");
        return PptxPackage.ResolveTarget(slidePart, comments[0].Target);
    }

    private static XElement CommentWithText(PptxPackage package, string part, string text)
    {
        var matches = PackageAssert.Select(package, part, "/p:cmLst/p:cm")
            .Where(cm => cm.Element(Ns.P + "text")?.Value == text)
            .ToList();
        Assert.True(matches.Count == 1,
            $"Expected one comment reading '{text}', found {matches.Count}.{Environment.NewLine}{PackageAssert.Describe(package, part)}");
        return matches[0];
    }

    private static XElement ThreadingInfo(PptxPackage package, string part, XElement comment)
    {
        var info = ThreadingInfoOrNull(comment);
        Assert.True(info is not null,
            $"The comment '{comment.Element(Ns.P + "text")?.Value}' has no p15:threadingInfo in a " +
            $"p:ext with uri {ThreadingExtensionUri}.{Environment.NewLine}{PackageAssert.Describe(package, part)}");
        return info!;
    }

    private static XElement? ThreadingInfoOrNull(XElement comment) =>
        comment.Elements(Ns.P + "extLst").Elements(Ns.P + "ext")
            .Where(ext => (string?)ext.Attribute("uri") == ThreadingExtensionUri)
            .Elements(Ns.P15 + "threadingInfo")
            .SingleOrDefault();
}
