using Aspose.Slides.Foss.ConformanceTests.Harness;
using Aspose.Slides.Foss.Export;

namespace Aspose.Slides.Foss.ConformanceTests;

/// <summary>
/// <c>docProps/app.xml</c> is what a file browser, a search indexer and PowerPoint's own info pane
/// read to describe a deck without opening it. A package that ships an extended-properties part
/// saying nothing about its own contents is worse than one that ships none: the part is declared and
/// the numbers in it are wrong.
/// </summary>
public sealed class DocumentPropertiesConformanceTests : IDisposable
{
    private readonly TestWorkspace _workspace = new();

    public void Dispose() => _workspace.Dispose();

    [Fact]
    public void ExtendedPropertiesReportTheNumberOfSlidesInTheDeck()
    {
        var path = _workspace.PathFor("app-properties.pptx");

        using (var presentation = new Presentation())
        {
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);

        var slideCount = PackageAssert.Select(package, "ppt/presentation.xml", "//p:sldIdLst/p:sldId").Count;
        Assert.Equal(3, slideCount);

        var declared = package.Xml("docProps/app.xml").Root?.Element(Ns.Ep + "Slides")?.Value;

        Assert.True(declared == slideCount.ToString(),
            $"The deck registers {slideCount} slides and docProps/app.xml declares " +
            $"{(declared is null ? "no <Slides> element at all" : $"'{declared}'")}." +
            $"{Environment.NewLine}{PackageAssert.Describe(package, "docProps/app.xml")}");
    }

    /// <summary>
    /// CT_Properties is a sequence, not a bag: an element in the wrong position is a schema error
    /// even when every element present is a legal one. Properties are set in whatever order the
    /// caller sets them, so the part has to impose the order itself on the way out.
    /// </summary>
    [Fact]
    public void ExtendedPropertiesProduceASchemaValidPackage()
    {
        var path = _workspace.PathFor("app-properties-valid.pptx");

        using (var presentation = new Presentation())
        {
            // Deliberately set in an order the schema does not declare them in.
            presentation.DocumentProperties.HyperlinkBase = "https://example.com/";
            presentation.DocumentProperties.PresentationFormat = "On-screen Show (4:3)";
            presentation.DocumentProperties.Company = "Example";
            presentation.DocumentProperties.ApplicationTemplate = "Blank";
            presentation.Save(path, SaveFormat.Pptx);
        }

        SchemaValidation.HasNoSchemaErrors(path);
    }

    [Fact]
    public void ExtendedPropertiesAreUpdatedWhenSlidesAreRemoved()
    {
        var path = _workspace.PathFor("app-properties-after-removal.pptx");

        using (var presentation = new Presentation())
        {
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Slides.RemoveAt(2);
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);

        var slideCount = PackageAssert.Select(package, "ppt/presentation.xml", "//p:sldIdLst/p:sldId").Count;
        var declared = package.Xml("docProps/app.xml").Root?.Element(Ns.Ep + "Slides")?.Value;

        Assert.True(declared == slideCount.ToString(),
            $"The deck registers {slideCount} slides and docProps/app.xml declares " +
            $"{(declared is null ? "no <Slides> element at all" : $"'{declared}'")}." +
            $"{Environment.NewLine}{PackageAssert.Describe(package, "docProps/app.xml")}");
    }
}
