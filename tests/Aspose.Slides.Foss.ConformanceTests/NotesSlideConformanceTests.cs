using Aspose.Slides.Foss.ConformanceTests.Harness;
using Aspose.Slides.Foss.Export;

namespace Aspose.Slides.Foss.ConformanceTests;

/// <summary>
/// ECMA-376 Part 1 §13.3.5 gives a notes slide exactly one implicit relationship to a notes master.
/// PowerPoint reads the notes text back without it, so the omission is latent — until a stricter
/// consumer refuses the package or a notes layout has nothing to inherit from.
/// </summary>
public sealed class NotesSlideConformanceTests : IDisposable
{
    private readonly TestWorkspace _workspace = new();

    public void Dispose() => _workspace.Dispose();

    [Fact]
    public void ANotesSlideReferencesANotesMasterThatIsInThePackage()
    {
        var path = WriteDeckWithNotes("notes-master.pptx");

        using var package = PptxPackage.Open(path);

        const string notesSlide = "ppt/notesSlides/notesSlide1.xml";
        Assert.True(package.Contains(notesSlide),
            $"No notes slide was written. Parts: {string.Join(", ", package.PartNames)}");

        var toMaster = package.Relationships(notesSlide)
            .Where(relationship => relationship.Type == Ns.NotesMasterRelationshipType)
            .ToList();

        Assert.True(toMaster.Count == 1,
            $"The notes slide declares {toMaster.Count} notesMaster relationship(s), expected exactly one." +
            $"{Environment.NewLine}{PackageAssert.Describe(package, PptxPackage.RelationshipsPartNameFor(notesSlide))}");

        var target = PptxPackage.ResolveTarget(notesSlide, toMaster[0].Target);
        Assert.True(package.Contains(target),
            $"The notes slide points at '{target}', which is not in the package. " +
            $"Parts: {string.Join(", ", package.PartNames)}");
    }

    [Fact]
    public void ANotesMasterIsRegisteredInThePresentationPart()
    {
        var path = WriteDeckWithNotes("notes-master-idlst.pptx");

        using var package = PptxPackage.Open(path);

        var registered = PackageAssert.Select(
            package, "ppt/presentation.xml", "//p:notesMasterIdLst/p:notesMasterId");

        Assert.True(registered.Count == 1,
            $"ppt/presentation.xml registers {registered.Count} notes master(s), expected exactly one." +
            $"{Environment.NewLine}{PackageAssert.Describe(package, "ppt/presentation.xml")}");
    }

    /// <summary>
    /// CT_Presentation is a sequence, so a notes master list has to sit between the slide master
    /// list and the slide list. Every slide mutation rewrites the slide list, and putting it back in
    /// the wrong place makes the presentation part invalid while every element in it is legal.
    /// </summary>
    [Fact]
    public void ChangingTheSlideListKeepsThePresentationPartInSchemaOrder()
    {
        var path = _workspace.PathFor("notes-then-slide-change.pptx");

        using (var presentation = new Presentation())
        {
            var notes = presentation.Slides[0].NotesSlideManager.AddNotesSlide();
            notes.NotesTextFrame!.Text = "Speaker notes";
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Save(path, SaveFormat.Pptx);
        }

        SchemaValidation.HasNoSchemaErrors(path);
    }

    /// <summary>
    /// A theme part belongs to exactly one master. PowerPoint refuses a package whose notes master
    /// and slide master relate to the same theme part — it validates clean against the schema, so
    /// only PowerPoint itself reports it, and only by refusing to open the file at all.
    /// </summary>
    [Fact]
    public void NoThemePartIsSharedBetweenTwoMasters()
    {
        var path = WriteDeckWithNotes("notes-master-theme.pptx");

        using var package = PptxPackage.Open(path);

        AssertEachMasterHasItsOwnTheme(package);
    }

    /// <summary>
    /// The same rule on the path the library is most used for: open an existing deck, add speaker
    /// notes, save. The notes master is created here too, so it needs its own theme here too.
    /// </summary>
    [Fact]
    public void AddingNotesToAnExistingDeckLeavesEveryMasterItsOwnTheme()
    {
        var path = _workspace.PathFor("loaded-notes.pptx");

        using (var presentation = new Presentation(TestFixtures.SimpleDeck))
        {
            var notes = presentation.Slides[0].NotesSlideManager.AddNotesSlide();
            notes.NotesTextFrame!.Text = "Speaker notes";
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);

        AssertEachMasterHasItsOwnTheme(package);
        SchemaValidation.HasNoSchemaErrors(path);
    }

    /// <summary>
    /// Asserts that every master part in the package relates to a theme part, and that no two
    /// masters relate to the same one.
    /// </summary>
    private static void AssertEachMasterHasItsOwnTheme(PptxPackage package)
    {
        var masters = package.ContentPartNames
            .Where(IsMasterPart)
            .ToList();

        Assert.True(masters.Count > 0,
            $"The package has no master parts at all. Parts: {string.Join(", ", package.PartNames)}");

        var themeOfMaster = new Dictionary<string, string>(StringComparer.Ordinal);

        foreach (var master in masters)
        {
            var themes = package.Relationships(master)
                .Where(relationship => relationship.Type == Ns.ThemeRelationshipType)
                .Select(relationship => PptxPackage.ResolveTarget(master, relationship.Target))
                .ToList();

            Assert.True(themes.Count == 1,
                $"{master} relates to {themes.Count} theme part(s), expected exactly one." +
                $"{Environment.NewLine}{PackageAssert.Describe(package, PptxPackage.RelationshipsPartNameFor(master))}");

            themeOfMaster[master] = themes[0];
        }

        var shared = themeOfMaster
            .GroupBy(pair => pair.Value, StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Count() > 1)
            .Select(group => $"'{group.Key}' is related from {string.Join(" and ", group.Select(pair => pair.Key))}")
            .ToList();

        Assert.True(shared.Count == 0,
            $"{shared.Count} theme part(s) are shared between masters:{Environment.NewLine}" +
            string.Join(Environment.NewLine, shared));
    }

    private static bool IsMasterPart(string partName) =>
        (partName.StartsWith("ppt/slideMasters/", StringComparison.Ordinal)
            || partName.StartsWith("ppt/notesMasters/", StringComparison.Ordinal)
            || partName.StartsWith("ppt/handoutMasters/", StringComparison.Ordinal))
        && partName.EndsWith(".xml", StringComparison.Ordinal);

    private string WriteDeckWithNotes(string fileName)
    {
        var path = _workspace.PathFor(fileName);

        using var presentation = new Presentation();
        var notes = presentation.Slides[0].NotesSlideManager.AddNotesSlide();
        notes.NotesTextFrame!.Text = "Speaker notes";
        presentation.Save(path, SaveFormat.Pptx);

        return path;
    }
}
