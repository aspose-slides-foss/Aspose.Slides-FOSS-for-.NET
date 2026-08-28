using Aspose.Slides.Foss.ConformanceTests.Harness;
using Aspose.Slides.Foss.Export;

namespace Aspose.Slides.Foss.ConformanceTests;

/// <summary>
/// Sections live in a PowerPoint extension on <c>ppt/presentation.xml</c>. An API that accepts a
/// section, reports it back from memory and writes nothing is indistinguishable from a working one
/// until the file is reopened somewhere else.
/// </summary>
public sealed class SectionConformanceTests : IDisposable
{
    private readonly TestWorkspace _workspace = new();

    public void Dispose() => _workspace.Dispose();

    [Fact]
    public void AnAddedSectionIsWrittenToThePresentationPart()
    {
        var path = _workspace.PathFor("sections.pptx");

        using (var presentation = new Presentation())
        {
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Sections.AddSection("Intro", presentation.Slides[0]);
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);

        var sections = PackageAssert.Select(package, "ppt/presentation.xml", "//p14:sectionLst/p14:section");

        Assert.True(sections.Count == 1,
            $"One section was added and the package declares {sections.Count}." +
            $"{Environment.NewLine}{PackageAssert.Describe(package, "ppt/presentation.xml")}");

        Assert.Equal("Intro", (string?)sections[0].Attribute("name"));
    }

    [Fact]
    public void ASectionRegistersTheSlidesItContains()
    {
        var path = _workspace.PathFor("sections-slides.pptx");

        using (var presentation = new Presentation())
        {
            presentation.Sections.AddSection("Intro", presentation.Slides[0]);
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);

        var sectionSlideIds = PackageAssert.Select(
            package, "ppt/presentation.xml", "//p14:sectionLst/p14:section/p14:sldIdLst/p14:sldId");

        Assert.True(sectionSlideIds.Count == 1,
            $"The section should list the one slide it starts from; the package lists {sectionSlideIds.Count}." +
            $"{Environment.NewLine}{PackageAssert.Describe(package, "ppt/presentation.xml")}");
    }
}
