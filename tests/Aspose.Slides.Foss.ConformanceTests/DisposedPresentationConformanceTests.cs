using Aspose.Slides.Foss.ConformanceTests.Harness;
using Aspose.Slides.Foss.Export;

namespace Aspose.Slides.Foss.ConformanceTests;

/// <summary>
/// Saving a presentation that has already been disposed must fail loudly and leave the target file
/// alone. Creating the file first and only then discovering there is nothing to write into it
/// replaces the user's document with an empty one — and PowerPoint opens a 0-byte file as an empty
/// deck rather than reporting damage, so nothing tells them.
/// </summary>
public sealed class DisposedPresentationConformanceTests : IDisposable
{
    private readonly TestWorkspace _workspace = new();

    public void Dispose() => _workspace.Dispose();

    [Fact]
    public void SavingAfterDisposeRaisesAndWritesNothing()
    {
        var path = _workspace.PathFor("after-dispose.pptx");

        var presentation = new Presentation();
        presentation.Dispose();

        var thrown = Record.Exception(() => presentation.Save(path, SaveFormat.Pptx));

        Assert.True(thrown is ObjectDisposedException,
            $"Save after Dispose {(thrown is null ? "returned successfully" : $"threw {thrown.GetType().Name}")}, " +
            $"expected ObjectDisposedException. " +
            (File.Exists(path)
                ? $"It wrote a {new FileInfo(path).Length}-byte file to the target path."
                : "No file was written."));

        Assert.False(File.Exists(path),
            $"Save after Dispose created '{Path.GetFileName(path)}' " +
            $"({(File.Exists(path) ? new FileInfo(path).Length : 0)} bytes). " +
            "A failed save must not replace what is already at the target path.");
    }

    [Fact]
    public void SavingAfterDisposeDoesNotOverwriteAnExistingDeck()
    {
        var path = _workspace.PathFor("existing-deck.pptx");

        using (var original = new Presentation())
        {
            original.Slides.AddEmptySlide(original.LayoutSlides[0]);
            original.Save(path, SaveFormat.Pptx);
        }

        var lengthBefore = new FileInfo(path).Length;

        var presentation = new Presentation();
        presentation.Dispose();
        Record.Exception(() => presentation.Save(path, SaveFormat.Pptx));

        var lengthAfter = new FileInfo(path).Length;

        Assert.True(lengthBefore == lengthAfter,
            $"A save that could not succeed truncated an existing deck from {lengthBefore} to {lengthAfter} bytes.");
    }
}
