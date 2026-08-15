using Aspose.Slides.Foss.ConformanceTests.Harness;
using Aspose.Slides.Foss.Export;

namespace Aspose.Slides.Foss.ConformanceTests;

/// <summary>
/// Opening a deck this library did not write, changing something in it and saving it again must not
/// cost the caller anything else in the file.
/// <para>
/// A package holds parts no API here exposes — <c>ppt/presProps.xml</c>, <c>ppt/viewProps.xml</c>,
/// <c>ppt/tableStyles.xml</c>, a thumbnail, a second theme, the layouts a slide never uses. Nothing
/// in the object model would show their loss, and PowerPoint opens the result without complaint;
/// the deck simply comes back smaller than it went in. The only way to see it is to compare the part
/// names in the ZIP before and after, which is what these tests do.
/// </para>
/// </summary>
public sealed class RoundTripConformanceTests : IDisposable
{
    private readonly TestWorkspace _workspace = new();

    public void Dispose() => _workspace.Dispose();

    [Fact]
    public void SavingAPowerPointDeckUnchangedKeepsEveryPart()
    {
        var path = _workspace.PathFor("roundtrip-unchanged.pptx");

        using (var presentation = new Presentation(TestFixtures.PowerPointDeck))
        {
            presentation.Save(path, SaveFormat.Pptx);
        }

        AssertSamePartSet(TestFixtures.PowerPointDeck, path);
    }

    [Fact]
    public void AddingAShapeToAPowerPointDeckKeepsEveryOtherPart()
    {
        var path = _workspace.PathFor("roundtrip-mutated.pptx");

        using (var presentation = new Presentation(TestFixtures.PowerPointDeck))
        {
            presentation.Slides[0].Shapes!.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);
            presentation.Save(path, SaveFormat.Pptx);
        }

        AssertSamePartSet(TestFixtures.PowerPointDeck, path);

        using var package = PptxPackage.Open(path);

        // The deck's own two shapes, plus the one just added.
        var shapes = PackageAssert.Select(package, "ppt/slides/slide1.xml", "//p:spTree/p:sp");
        Assert.True(shapes.Count == 3,
            $"slide1 carries {shapes.Count} shapes, expected the deck's two plus the added one.");
    }

    [Fact]
    public void APowerPointDeckSurvivesTheChecksEveryProducedPackageFaces()
    {
        var path = _workspace.PathFor("roundtrip-checked.pptx");

        using (var presentation = new Presentation(TestFixtures.PowerPointDeck))
        {
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);
        PackageAssert.EveryPartResolvesAContentType(package);
        PackageAssert.EveryPartHasItsRequiredContentType(package);
        PackageAssert.NoOverrideNamesAMissingPart(package);
        PackageAssert.EveryInternalRelationshipTargetExists(package);
        PackageAssert.AllRelationshipReferencesResolve(package);
        SchemaValidation.HasNoSchemaErrors(path);
    }

    /// <summary>
    /// Asserts that the produced package holds exactly the parts the input did — none dropped and
    /// none invented.
    /// </summary>
    private static void AssertSamePartSet(string inputPath, string producedPath)
    {
        using var input = PptxPackage.Open(inputPath);
        using var produced = PptxPackage.Open(producedPath);

        var before = input.PartNames.ToHashSet(StringComparer.Ordinal);
        var after = produced.PartNames.ToHashSet(StringComparer.Ordinal);

        var lost = before.Except(after).Order(StringComparer.Ordinal).ToList();
        var added = after.Except(before).Order(StringComparer.Ordinal).ToList();

        Assert.True(lost.Count == 0 && added.Count == 0,
            $"The saved package does not hold the parts it was opened with " +
            $"({before.Count} in, {after.Count} out)." +
            $"{Environment.NewLine}lost:  {(lost.Count == 0 ? "none" : string.Join(", ", lost))}" +
            $"{Environment.NewLine}added: {(added.Count == 0 ? "none" : string.Join(", ", added))}");
    }
}
