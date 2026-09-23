using Aspose.Slides.Foss.ConformanceTests.Harness;
using Aspose.Slides.Foss.Drawing;
using Aspose.Slides.Foss.Export;

namespace Aspose.Slides.Foss.ConformanceTests;

/// <summary>
/// A comment reply is either written in a form a consumer understands or it is not written at all.
/// An attribute the schema does not declare is the worst of the three outcomes: it makes the package
/// invalid, no reader renders the thread, and the API reports success.
/// </summary>
public sealed class CommentConformanceTests : IDisposable
{
    private readonly TestWorkspace _workspace = new();

    public void Dispose() => _workspace.Dispose();

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
            $"declare. Threading is expressed by a threadedComments part, not by an attribute on <p:cm>." +
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
}
