using System.Xml.Linq;
using Aspose.Slides.Foss.ConformanceTests.Harness;
using Aspose.Slides.Foss.Export;

namespace Aspose.Slides.Foss.ConformanceTests;

/// <summary>
/// Sections live in a PowerPoint extension on <c>ppt/presentation.xml</c>. An API that accepts a
/// section, reports it back from memory and writes nothing is indistinguishable from a working one
/// until the file is reopened somewhere else.
/// </summary>
/// <remarks>
/// A section names its slides by the numeric id of their <c>&lt;p:sldId&gt;</c>, not by part or
/// relationship, so every change to the slide list is also a change to the section list. The cases
/// below that open a file use decks PowerPoint wrote, and where PowerPoint was asked to make the same
/// edit, the expected sections are the ones it saved.
/// </remarks>
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

    // ── Removing slides ──────────────────────────────────────────

    [Fact]
    public void RemovingTheFirstSlideOfASectionLeavesTheEarlierSectionAlone()
    {
        var path = _workspace.PathFor("remove-section-start.pptx");

        using (var presentation = DeckWithSlides("A", "B", "C"))
        {
            presentation.Sections.AddSection("Intro", presentation.Slides[0]);
            presentation.Sections.AddSection("Body", presentation.Slides[1]);
            presentation.Slides.RemoveAt(1);
            presentation.Save(path, SaveFormat.Pptx);
        }

        // PowerPoint, deleting the same slide from the same deck, saves Intro [A] and Body [C].
        AssertSections(path, "Intro[Slide A] Body[Slide C]");
    }

    [Fact]
    public void RemovingTheOnlySlideOfASectionKeepsTheSectionEmptyInItsPlace()
    {
        var path = _workspace.PathFor("remove-section-only-slide.pptx");

        using (var presentation = DeckWithSlides("A", "B", "C"))
        {
            presentation.Sections.AddSection("Intro", presentation.Slides[0]);
            presentation.Sections.AddSection("Middle", presentation.Slides[1]);
            presentation.Sections.AddSection("Body", presentation.Slides[2]);
            presentation.Slides.RemoveAt(1);
            presentation.Save(path, SaveFormat.Pptx);
        }

        AssertSections(path, "Intro[Slide A] Middle[] Body[Slide C]");
    }

    [Fact]
    public void RemovingEverySlideOfTheLastSectionKeepsTheSectionEmpty()
    {
        var path = _workspace.PathFor("remove-section-all-slides.pptx");

        using (var presentation = DeckWithSlides("A", "B", "C"))
        {
            presentation.Sections.AddSection("Intro", presentation.Slides[0]);
            presentation.Sections.AddSection("Body", presentation.Slides[1]);
            presentation.Slides.RemoveAt(2);
            presentation.Slides.RemoveAt(1);
            presentation.Save(path, SaveFormat.Pptx);
        }

        // PowerPoint, deleting slides 3 and 2 of the same deck, keeps Body as an empty section.
        AssertSections(path, "Intro[Slide A] Body[]");
    }

    // ── Editing a deck opened from a file ────────────────────────

    [Fact]
    public void AddingASlideToAnOpenedDeckPutsItInTheLastSection()
    {
        var path = _workspace.PathFor("opened-add.pptx");

        using (var presentation = new Presentation(TestFixtures.PowerPointSections))
        {
            Label(presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]), "New");
            presentation.Save(path, SaveFormat.Pptx);
        }

        // PowerPoint, adding a slide at the end of the same deck, puts it in Body.
        AssertSections(path, "Intro[Slide A] Body[Slide B, Slide C, New]");
    }

    [Fact]
    public void InsertingASlideAtTheStartOfAnOpenedDeckKeepsEverySlideInItsSection()
    {
        var path = _workspace.PathFor("opened-insert.pptx");

        using (var presentation = new Presentation(TestFixtures.PowerPointSections))
        {
            Label(presentation.Slides.InsertEmptySlide(0, presentation.LayoutSlides[0]), "New");
            presentation.Save(path, SaveFormat.Pptx);
        }

        AssertSections(path, "Intro[New, Slide A] Body[Slide B, Slide C]");
    }

    [Fact]
    public void RemovingASlideFromAnOpenedDeckGivesTheSectionsPowerPointGives()
    {
        var path = _workspace.PathFor("opened-remove.pptx");

        using (var presentation = new Presentation(TestFixtures.PowerPointSections))
        {
            presentation.Slides.RemoveAt(1);
            presentation.Save(path, SaveFormat.Pptx);
        }

        using (var expected = PptxPackage.Open(TestFixtures.PowerPointSectionsWithIdGap))
            Assert.Equal("Intro[Slide A] Body[Slide C]", DescribeSections(expected));

        AssertSections(path, "Intro[Slide A] Body[Slide C]");
    }

    [Fact]
    public void AddingASlideToAnOpenedDeckWhoseSlideIdsHaveAGapKeepsItsSections()
    {
        var path = _workspace.PathFor("opened-gap-add.pptx");

        using (var presentation = new Presentation(TestFixtures.PowerPointSectionsWithIdGap))
        {
            Label(presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]), "New");
            presentation.Save(path, SaveFormat.Pptx);
        }

        AssertSections(path, "Intro[Slide A] Body[Slide C, New]");
    }

    [Fact]
    public void CloningASlideIntoAnOpenedDeckPutsTheCloneInTheLastSection()
    {
        var path = _workspace.PathFor("opened-clone.pptx");

        using (var presentation = new Presentation(TestFixtures.PowerPointSections))
        {
            presentation.Slides.AddClone(presentation.Slides[0]);
            presentation.Save(path, SaveFormat.Pptx);
        }

        AssertSections(path, "Intro[Slide A] Body[Slide B, Slide C, Slide A]");
    }

    [Fact]
    public void AnEmptySectionInAnOpenedDeckStaysWhereItWasWhenTheDeckIsEdited()
    {
        var withEmptySection = _workspace.PathFor("empty-middle-section.pptx");
        var path = _workspace.PathFor("empty-middle-section-edited.pptx");

        using (var presentation = DeckWithSlides("A", "B", "C"))
        {
            presentation.Sections.AddSection("Intro", presentation.Slides[0]);
            presentation.Sections.AddSection("Middle", presentation.Slides[1]);
            presentation.Sections.AddSection("Body", presentation.Slides[2]);
            presentation.Slides.RemoveAt(1);
            presentation.Save(withEmptySection, SaveFormat.Pptx);
        }

        AssertSections(withEmptySection, "Intro[Slide A] Middle[] Body[Slide C]");

        using (var presentation = new Presentation(withEmptySection))
        {
            Label(presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]), "New");
            presentation.Save(path, SaveFormat.Pptx);
        }

        AssertSections(path, "Intro[Slide A] Middle[] Body[Slide C, New]");
    }

    [Fact]
    public void AnOpenedDeckWithSectionsThatIsEditedIsStillSchemaValid()
    {
        var path = _workspace.PathFor("opened-edited-valid.pptx");

        using (var presentation = new Presentation(TestFixtures.PowerPointSections))
        {
            presentation.Slides.RemoveAt(1);
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Save(path, SaveFormat.Pptx);
        }

        using (var package = PptxPackage.Open(path))
            PackageAssert.SectionsPartitionTheSlideList(package);

        SchemaValidation.HasNoSchemaErrors(path);
    }

    // ── Saving a subset of the slides ────────────────────────────

    [Fact]
    public void SavingASubsetOfADeckWithSectionsListsOnlyTheSlidesItWrites()
    {
        var path = _workspace.PathFor("subset-sections.pptx");

        using (var presentation = DeckWithSlides("A", "B", "C"))
        {
            presentation.Sections.AddSection("Intro", presentation.Slides[0]);
            presentation.Sections.AddSection("Body", presentation.Slides[1]);
            presentation.Save(path, [0], SaveFormat.Pptx);
        }

        // PowerPoint, deleting slides 3 and 2 of the same deck, keeps Body as an empty section.
        AssertSections(path, "Intro[Slide A] Body[]");
    }

    [Fact]
    public void SavingASubsetOfAnOpenedDeckWithSectionsListsOnlyTheSlidesItWrites()
    {
        var path = _workspace.PathFor("opened-subset-sections.pptx");

        using (var presentation = new Presentation(TestFixtures.PowerPointSections))
            presentation.Save(path, [0, 2], SaveFormat.Pptx);

        AssertSections(path, "Intro[Slide A] Body[Slide C]");
    }

    // ── Helpers ──────────────────────────────────────────────────

    /// <summary>
    /// A new deck whose slides each carry one rectangle reading <c>Slide {name}</c>, so the package
    /// can say which slide ended up in which section without asking the library.
    /// </summary>
    private static Presentation DeckWithSlides(params string[] names)
    {
        var presentation = new Presentation();
        for (var i = 1; i < names.Length; i++)
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);

        for (var i = 0; i < names.Length; i++)
            Label(presentation.Slides[i], $"Slide {names[i]}");

        return presentation;
    }

    private static void Label(ISlide slide, string text) =>
        slide.Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 400, 80).AddTextFrame(text);

    /// <summary>
    /// Asserts, on the written package, that the sections partition the slide list and hold the
    /// expected slides.
    /// </summary>
    private static void AssertSections(string path, string expected)
    {
        using var package = PptxPackage.Open(path);

        PackageAssert.SectionsPartitionTheSlideList(package);

        var actual = DescribeSections(package);
        Assert.True(expected == actual,
            $"expected sections: {expected}{Environment.NewLine}written sections:  {actual}" +
            $"{Environment.NewLine}{PackageAssert.Describe(package, "ppt/presentation.xml")}");
    }

    /// <summary>
    /// Describes the section list as <c>Name[slide text, …]</c>, following each
    /// <c>&lt;p14:sldId&gt;</c> through <c>&lt;p:sldIdLst&gt;</c> and the presentation relationships
    /// to the slide part it names. A section id that names no slide reads as <c>?id</c>.
    /// </summary>
    private static string DescribeSections(PptxPackage package)
    {
        const string presentation = "ppt/presentation.xml";

        var relationships = package.Relationships(presentation).ToDictionary(r => r.Id, r => r.Target);
        var slideText = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var sldId in PackageAssert.Select(package, presentation, "/p:presentation/p:sldIdLst/p:sldId"))
        {
            var id = (string?)sldId.Attribute("id") ?? string.Empty;
            var relId = PackageAssert.AttributeValue(sldId, "r:id") ?? string.Empty;
            if (!relationships.TryGetValue(relId, out var target))
                continue;

            var part = PptxPackage.ResolveTarget(presentation, target);
            slideText[id] = string.Concat(package.Xml(part).Descendants(Ns.A + "t").Select(t => t.Value));
        }

        var sections = PackageAssert.Select(package, presentation, "//p14:sectionLst/p14:section")
            .Select(section =>
            {
                var slides = section.Elements(Ns.P14 + "sldIdLst").Elements(Ns.P14 + "sldId")
                    .Select(e => (string?)e.Attribute("id") ?? string.Empty)
                    .Select(id => slideText.TryGetValue(id, out var text) ? text : $"?{id}");
                return $"{(string?)section.Attribute("name")}[{string.Join(", ", slides)}]";
            });

        return string.Join(" ", sections);
    }
}
