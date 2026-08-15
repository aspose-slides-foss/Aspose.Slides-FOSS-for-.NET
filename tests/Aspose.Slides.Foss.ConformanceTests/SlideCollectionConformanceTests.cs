using Aspose.Slides.Foss.ConformanceTests.Harness;
using Aspose.Slides.Foss.Export;

namespace Aspose.Slides.Foss.ConformanceTests;

/// <summary>
/// A slide exists in the package when four things agree: the part, the relationship from
/// <c>ppt/presentation.xml</c>, the entry in <c>&lt;p:sldIdLst&gt;</c> and the content-type
/// <c>Override</c>. Adding and removing slides has to keep all four in step — the in-memory slide
/// count is not evidence of any of it.
/// </summary>
public sealed class SlideCollectionConformanceTests : IDisposable
{
    private readonly TestWorkspace _workspace = new();

    public void Dispose() => _workspace.Dispose();

    [Fact]
    public void SavingASubsetOfSlidesWritesOnlyTheSlidesThatWereAskedFor()
    {
        var path = _workspace.PathFor("subset.pptx");

        using (var presentation = new Presentation())
        {
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Save(path, [0], SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);
        var slideIds = PackageAssert.Select(package, "ppt/presentation.xml", "//p:sldIdLst/p:sldId");

        Assert.True(slideIds.Count == 1,
            $"One slide was requested and the saved package registers {slideIds.Count}:{Environment.NewLine}" +
            string.Join(Environment.NewLine, slideIds));
    }

    [Fact]
    public void SavingASubsetToAStreamWritesOnlyTheSlidesThatWereAskedFor()
    {
        var path = _workspace.PathFor("subset-stream.pptx");

        using (var presentation = new Presentation())
        {
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            using var stream = File.Create(path);
            presentation.Save(stream, [0], SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);
        var slideIds = PackageAssert.Select(package, "ppt/presentation.xml", "//p:sldIdLst/p:sldId");

        Assert.True(slideIds.Count == 1,
            $"One slide was requested and the saved package registers {slideIds.Count}:{Environment.NewLine}" +
            string.Join(Environment.NewLine, slideIds));
    }

    [Fact]
    public void RemovingASlideRemovesItsPartAndItsRelationship()
    {
        var path = _workspace.PathFor("removed-slide.pptx");

        using (var presentation = new Presentation())
        {
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Slides.RemoveAt(1);
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);

        var registeredIds = PackageAssert
            .Select(package, "ppt/presentation.xml", "//p:sldIdLst/p:sldId")
            .Select(element => PackageAssert.AttributeValue(element, "r:id"))
            .Where(id => id is not null)
            .ToHashSet(StringComparer.Ordinal);

        var slideRelationships = package.Relationships("ppt/presentation.xml")
            .Where(relationship => relationship.Type == Ns.SlideRelationshipType)
            .ToList();

        var orphanRelationships = slideRelationships
            .Where(relationship => !registeredIds.Contains(relationship.Id))
            .Select(relationship => $"{relationship.Id} -> {relationship.Target}")
            .ToList();

        Assert.True(orphanRelationships.Count == 0,
            $"The removed slide left {orphanRelationships.Count} slide relationship(s) that no <p:sldId> refers to: " +
            string.Join(", ", orphanRelationships));

        var referencedParts = slideRelationships
            .Where(relationship => registeredIds.Contains(relationship.Id))
            .Select(relationship => PptxPackage.ResolveTarget("ppt/presentation.xml", relationship.Target))
            .ToHashSet(StringComparer.Ordinal);

        var orphanParts = package.PartNames
            .Where(name => name.StartsWith("ppt/slides/slide", StringComparison.Ordinal)
                        && name.EndsWith(".xml", StringComparison.Ordinal)
                        && !referencedParts.Contains(name))
            .ToList();

        Assert.True(orphanParts.Count == 0,
            $"The removed slide's part(s) are still in the package and nothing refers to them: " +
            string.Join(", ", orphanParts));
    }
}
