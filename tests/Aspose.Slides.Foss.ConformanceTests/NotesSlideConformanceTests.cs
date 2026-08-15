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
