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
