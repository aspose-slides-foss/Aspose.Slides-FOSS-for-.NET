using Aspose.Slides.Foss.ConformanceTests.Harness;
using Aspose.Slides.Foss.Export;

namespace Aspose.Slides.Foss.ConformanceTests;

/// <summary>
/// Text beyond a single run has to go through <c>Paragraphs</c> and <c>Portions</c>. A paragraph
/// that never reaches <c>&lt;a:txBody&gt;</c> is lost in the quietest way there is: the call
/// returns, the save succeeds, PowerPoint opens the file, and the text is simply not in it.
/// </summary>
public sealed class TextFrameConformanceTests : IDisposable
{
    private readonly TestWorkspace _workspace = new();

    public void Dispose() => _workspace.Dispose();

    [Fact]
    public void AnAddedParagraphIsWrittenWithItsPortions()
    {
        var path = _workspace.PathFor("paragraph-add.pptx");

        using (var presentation = new Presentation())
        {
            var shape = presentation.Slides[0].Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 300, 100);
            var textFrame = shape.TextFrame!;
            textFrame.Text = "First paragraph";

            var second = new Paragraph();
            second.Portions.Add(new Portion("Second-A "));
            second.Portions.Add(new Portion("Second-B"));
            textFrame.Paragraphs.Add(second);

            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);

        AssertParagraphTexts(package, ["First paragraph", "Second-A Second-B"]);

        var portions = PackageAssert.Select(package, "ppt/slides/slide1.xml", "//p:txBody/a:p[2]/a:r/a:t");
        Assert.True(portions.Count == 2,
            $"The added paragraph carries {portions.Count} portion(s), expected the two it was given." +
            $"{Environment.NewLine}{PackageAssert.Describe(package, "ppt/slides/slide1.xml")}");
        Assert.Equal("Second-A ", portions[0].Value);
        Assert.Equal("Second-B", portions[1].Value);
    }

    [Fact]
    public void AnAddedParagraphCarryingOnlyTextIsWrittenAsARun()
    {
        var path = _workspace.PathFor("paragraph-add-text.pptx");

        using (var presentation = new Presentation())
        {
            var shape = presentation.Slides[0].Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 300, 100);
            var textFrame = shape.TextFrame!;
            textFrame.Text = "First paragraph";
            textFrame.Paragraphs.Add(new Paragraph { Text = "Second paragraph" });

            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);

        AssertParagraphTexts(package, ["First paragraph", "Second paragraph"]);
    }

    [Fact]
    public void AnInsertedParagraphIsWrittenAtTheIndexItWasGiven()
    {
        var path = _workspace.PathFor("paragraph-insert.pptx");

        using (var presentation = new Presentation())
        {
            var shape = presentation.Slides[0].Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 300, 100);
            var textFrame = shape.TextFrame!;
            textFrame.Text = "Second paragraph";
            textFrame.Paragraphs.Insert(0, new Paragraph { Text = "First paragraph" });

            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);

        AssertParagraphTexts(package, ["First paragraph", "Second paragraph"]);
    }

    [Fact]
    public void ADeckWithSeveralParagraphsIsSchemaValid()
    {
        var path = _workspace.PathFor("paragraph-valid.pptx");

        using (var presentation = new Presentation())
        {
            var shape = presentation.Slides[0].Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 300, 100);
            var textFrame = shape.TextFrame!;
            textFrame.Text = "First paragraph";

            var second = new Paragraph();
            second.Portions.Add(new Portion("Second-A "));
            second.Portions.Add(new Portion("Second-B"));
            textFrame.Paragraphs.Add(second);

            presentation.Save(path, SaveFormat.Pptx);
        }

        SchemaValidation.HasNoSchemaErrors(path);
    }

    /// <summary>
    /// Asserts that the first slide's text body carries exactly these paragraphs, each read as the
    /// concatenation of its runs — which is what any consumer of the file sees.
    /// </summary>
    private static void AssertParagraphTexts(PptxPackage package, string[] expected)
    {
        var paragraphs = PackageAssert.Select(package, "ppt/slides/slide1.xml", "//p:txBody/a:p");

        var written = paragraphs
            .Select(paragraph => string.Concat(
                paragraph.Elements(Ns.A + "r").Select(run => run.Element(Ns.A + "t")?.Value ?? string.Empty)))
            .ToList();

        Assert.True(written.SequenceEqual(expected),
            $"The text body carries {written.Count} paragraph(s) [{string.Join(" | ", written)}], " +
            $"expected {expected.Length} [{string.Join(" | ", expected)}]." +
            $"{Environment.NewLine}{PackageAssert.Describe(package, "ppt/slides/slide1.xml")}");
    }
}
