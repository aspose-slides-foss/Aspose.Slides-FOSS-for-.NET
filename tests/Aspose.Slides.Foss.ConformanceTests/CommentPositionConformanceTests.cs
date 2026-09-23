using System.Globalization;
using Aspose.Slides.Foss.ConformanceTests.Harness;
using Aspose.Slides.Foss.Drawing;
using Aspose.Slides.Foss.Export;

namespace Aspose.Slides.Foss.ConformanceTests;

/// <summary>
/// A comment must be placed where the caller put it, in the unit PowerPoint reads.
/// </summary>
/// <remarks>
/// The schema types <c>p:cm/p:pos</c> as a point in EMU, but PowerPoint writes and reads it in
/// eighths of a point, 576 to the inch: a comment PowerPoint places 10 pt from the top-left corner
/// of the slide is written <c>&lt;p:pos x="80" y="80"/&gt;</c>. A position written in EMU is 1587.5
/// times too large, and PowerPoint shows the comment far outside the slide. A comment's position in
/// this library's API is in centimetres from the top-left corner of the slide.
/// </remarks>
public sealed class CommentPositionConformanceTests : IDisposable
{
    /// <summary>PowerPoint's comment-position units per centimetre: 576 per inch.</summary>
    private const double UnitsPerCm = 576 / 2.54;

    /// <summary>EMU per comment-position unit: 12,700 EMU per point, eight units per point.</summary>
    private const double EmuPerUnit = 12700 / 8.0;

    private static readonly DateTime When = new(2026, 1, 2, 3, 4, 5, DateTimeKind.Utc);

    private readonly TestWorkspace _workspace = new();

    public void Dispose() => _workspace.Dispose();

    [Fact]
    public void ACommentIsWrittenInTheUnitPowerPointReads()
    {
        // One inch across and two down: PowerPoint writes 72 pt and 144 pt as 576 and 1152.
        var path = WriteComment("comment-one-inch.pptx", new PointF(2.54f, 5.08f));

        using var package = PptxPackage.Open(path);
        var part = CommentConformanceTests.CommentsPartOfSlide(package, 0);
        PackageAssert.ElementWithAttributes(package, part, "/p:cmLst/p:cm/p:pos", ("x", "576"), ("y", "1152"));
    }

    [Theory]
    [InlineData(0f, 0f)]
    [InlineData(1f, 1f)]
    [InlineData(2f, 3f)]
    [InlineData(12.7f, 9.525f)]
    [InlineData(25.4f, 19.05f)]
    public void ACommentPlacedOnTheSlideLiesOnTheSlide(float xCm, float yCm)
    {
        var path = WriteComment($"comment-{xCm.ToString(CultureInfo.InvariantCulture)}-{yCm.ToString(CultureInfo.InvariantCulture)}.pptx",
            new PointF(xCm, yCm));

        using var package = PptxPackage.Open(path);
        var size = PackageAssert.SingleElement(package, "ppt/presentation.xml", "/p:presentation/p:sldSz");
        var width = long.Parse((string)size.Attribute("cx")!, CultureInfo.InvariantCulture) / EmuPerUnit;
        var height = long.Parse((string)size.Attribute("cy")!, CultureInfo.InvariantCulture) / EmuPerUnit;

        var pos = PackageAssert.SingleElement(package, CommentConformanceTests.CommentsPartOfSlide(package, 0), "/p:cmLst/p:cm/p:pos");
        var x = long.Parse((string)pos.Attribute("x")!, CultureInfo.InvariantCulture);
        var y = long.Parse((string)pos.Attribute("y")!, CultureInfo.InvariantCulture);

        Assert.True(x == (long)Math.Round(xCm * UnitsPerCm) && y == (long)Math.Round(yCm * UnitsPerCm),
            $"({xCm} cm, {yCm} cm) was written as <p:pos x=\"{x}\" y=\"{y}\"/>; PowerPoint's unit gives " +
            $"x=\"{Math.Round(xCm * UnitsPerCm)}\" y=\"{Math.Round(yCm * UnitsPerCm)}\".");
        Assert.True(x >= 0 && x <= Math.Round(width) && y >= 0 && y <= Math.Round(height),
            $"({xCm} cm, {yCm} cm) was written as <p:pos x=\"{x}\" y=\"{y}\"/>, outside a slide " +
            $"{width} by {height} units.");
    }

    [Fact]
    public void AnInsertedCommentIsWrittenInTheUnitPowerPointReads()
    {
        var path = _workspace.PathFor("comment-inserted.pptx");
        using (var presentation = new Presentation())
        {
            var author = presentation.CommentAuthors.AddAuthor("Reviewer", "RV");
            author.Comments.InsertComment(0, "Here", presentation.Slides[0], new PointF(5.08f, 2.54f), When);
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);
        var part = CommentConformanceTests.CommentsPartOfSlide(package, 0);
        PackageAssert.ElementWithAttributes(package, part, "/p:cmLst/p:cm/p:pos", ("x", "1152"), ("y", "576"));
    }

    [Fact]
    public void AMovedCommentIsWrittenInTheUnitPowerPointReads()
    {
        var path = _workspace.PathFor("comment-moved.pptx");
        using (var presentation = new Presentation())
        {
            var author = presentation.CommentAuthors.AddAuthor("Reviewer", "RV");
            var comment = author.Comments.AddComment("Here", presentation.Slides[0], new PointF(1f, 1f), When);
            comment.Position = new PointF(2.54f, 7.62f);
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);
        var part = CommentConformanceTests.CommentsPartOfSlide(package, 0);
        PackageAssert.ElementWithAttributes(package, part, "/p:cmLst/p:cm/p:pos", ("x", "576"), ("y", "1728"));
    }

    /// <summary>
    /// A reader test: the expected value is PowerPoint's own <c>x="80" y="80"</c> for a comment it
    /// placed at 10 pt, which is 10 / 72 × 2.54 cm.
    /// </summary>
    [Fact]
    public void ACommentPowerPointPlacedReadsBackWherePowerPointPutIt()
    {
        using var presentation = new Presentation(TestFixtures.PowerPointCommentReply);

        var parent = presentation.CommentAuthors[0].Comments.ToArray().Single(c => c.Text == "Parent comment");
        var expected = 10 / 72.0 * 2.54;

        Assert.True(Math.Abs(parent.Position.X - expected) < 0.001 && Math.Abs(parent.Position.Y - expected) < 0.001,
            $"A comment PowerPoint placed at 10 pt (<p:pos x=\"80\" y=\"80\"/>) reads back as " +
            $"({parent.Position.X} cm, {parent.Position.Y} cm); expected ({expected:F4} cm, {expected:F4} cm).");
    }

    [Fact]
    public void ACommentPowerPointPlacedKeepsItsPositionWhenItIsSetBackAndSaved()
    {
        var path = _workspace.PathFor("comment-powerpoint-position-resaved.pptx");
        using (var presentation = new Presentation(TestFixtures.PowerPointCommentReply))
        {
            foreach (var comment in presentation.CommentAuthors[0].Comments.ToArray())
                comment.Position = comment.Position;
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);
        var part = CommentConformanceTests.CommentsPartOfSlide(package, 0);
        var written = PackageAssert.Select(package, part, "/p:cmLst/p:cm/p:pos")
            .Select(p => $"{(string?)p.Attribute("x")},{(string?)p.Attribute("y")}")
            .ToList();

        Assert.True(written.SequenceEqual(["80,80", "80,176"]),
            $"PowerPoint wrote the two comments at 80,80 and 80,176; after reading each position and " +
            $"setting it back they were written at {string.Join(" and ", written)}." +
            $"{Environment.NewLine}{PackageAssert.Describe(package, part)}");
    }

    private string WriteComment(string fileName, PointF position)
    {
        var path = _workspace.PathFor(fileName);
        using var presentation = new Presentation();
        var author = presentation.CommentAuthors.AddAuthor("Reviewer", "RV");
        author.Comments.AddComment("Here", presentation.Slides[0], position, When);
        presentation.Save(path, SaveFormat.Pptx);
        return path;
    }
}
