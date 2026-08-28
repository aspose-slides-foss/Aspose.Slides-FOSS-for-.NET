using Aspose.Slides.Foss.ConformanceTests.Harness;
using Aspose.Slides.Foss.Export;

namespace Aspose.Slides.Foss.ConformanceTests;

/// <summary>
/// Per ISO/IEC 29500-2 §10.1.2 a part's content type is its identity. A slide part that falls through
/// to <c>&lt;Default Extension="xml"/&gt;</c> is an <c>application/xml</c> part as far as any reader is
/// concerned — PowerPoint guesses and opens it anyway, while the Open XML SDK and python-pptx refuse
/// the package outright. That makes this the failure a user meets in their own toolchain rather than
/// in PowerPoint.
/// </summary>
public sealed class ContentTypeConformanceTests : IDisposable
{
    private readonly TestWorkspace _workspace = new();

    public void Dispose() => _workspace.Dispose();

    [Fact]
    public void AnAddedSlideIsDeclaredAsASlide()
    {
        var path = _workspace.PathFor("added-slide.pptx");

        using (var presentation = new Presentation())
        {
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);
        PackageAssert.EveryPartHasItsRequiredContentType(package);
    }

    [Fact]
    public void AnInsertedSlideIsDeclaredAsASlide()
    {
        var path = _workspace.PathFor("inserted-slide.pptx");

        using (var presentation = new Presentation())
        {
            presentation.Slides.InsertEmptySlide(0, presentation.LayoutSlides[0]);
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);
        PackageAssert.EveryPartHasItsRequiredContentType(package);
    }

    [Fact]
    public void AClonedSlideIsDeclaredAsASlide()
    {
        var path = _workspace.PathFor("cloned-slide.pptx");

        using (var presentation = new Presentation())
        {
            presentation.Slides.AddClone(presentation.Slides[0]);
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);
        PackageAssert.EveryPartHasItsRequiredContentType(package);
    }

    [Fact]
    public void ANotesSlideIsDeclaredAsANotesSlide()
    {
        var path = _workspace.PathFor("speaker-notes.pptx");

        using (var presentation = new Presentation())
        {
            var notes = presentation.Slides[0].NotesSlideManager.AddNotesSlide();
            notes.NotesTextFrame!.Text = "Speaker notes";
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);
        PackageAssert.EveryPartHasItsRequiredContentType(package);
    }

    [Fact]
    public void AModifiedPackageIsAcceptedByAStrictReader()
    {
        var path = _workspace.PathFor("strict-reader.pptx");

        using (var presentation = new Presentation())
        {
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Save(path, SaveFormat.Pptx);
        }

        SchemaValidation.HasNoSchemaErrors(path);
    }
}
