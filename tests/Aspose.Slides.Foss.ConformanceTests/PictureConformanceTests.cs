using Aspose.Slides.Foss.ConformanceTests.Harness;
using Aspose.Slides.Foss.Export;

namespace Aspose.Slides.Foss.ConformanceTests;

/// <summary>
/// An embedded image needs three things in the package, and all three or none: the bytes, a
/// relationship from the slide to those bytes, and a content type for them. Any two of the three
/// produce a file that either will not open or shows a placeholder where the picture should be.
/// </summary>
public sealed class PictureConformanceTests : IDisposable
{
    private readonly TestWorkspace _workspace = new();

    public void Dispose() => _workspace.Dispose();

    [Fact]
    public void APictureFrameResolvesTheImageItPointsAt()
    {
        var path = _workspace.PathFor("picture-frame.pptx");

        using (var presentation = new Presentation())
        {
            var image = presentation.Images.AddImage(TestWorkspace.Png());
            presentation.Slides[0].Shapes!.AddPictureFrame(ShapeType.Rectangle, 50, 50, 100, 100, image);
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);
        PackageAssert.AllRelationshipReferencesResolve(package);
    }

    [Fact]
    public void APictureFrameBuiltFromAStreamResolvesTheImageItPointsAt()
    {
        var path = _workspace.PathFor("picture-frame-stream.pptx");

        using (var presentation = new Presentation())
        {
            using var stream = new MemoryStream(TestWorkspace.Png());
            var image = presentation.Images.AddImage(stream);
            presentation.Slides[0].Shapes!.AddPictureFrame(ShapeType.Rectangle, 50, 50, 100, 100, image);
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);
        PackageAssert.AllRelationshipReferencesResolve(package);
    }

    [Fact]
    public void AnEmbeddedImageDeclaresItsContentType()
    {
        var path = _workspace.PathFor("picture-content-type.pptx");

        using (var presentation = new Presentation())
        {
            var image = presentation.Images.AddImage(TestWorkspace.Png());
            presentation.Slides[0].Shapes!.AddPictureFrame(ShapeType.Rectangle, 50, 50, 100, 100, image);
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);

        var media = package.PartNames.Where(n => n.StartsWith("ppt/media/", StringComparison.Ordinal)).ToList();
        Assert.True(media.Count > 0, $"No media part was written. Parts: {string.Join(", ", package.PartNames)}");

        PackageAssert.EveryPartHasItsRequiredContentType(package);
    }

    [Fact]
    public void APictureFilledShapeCarriesARealRelationshipAndNoInternalMarker()
    {
        var path = _workspace.PathFor("picture-fill.pptx");

        using (var presentation = new Presentation())
        {
            var image = presentation.Images.AddImage(TestWorkspace.Png());
            var shape = presentation.Slides[0].Shapes!.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);
            shape.FillFormat!.FillType = FillType.Picture;
            shape.FillFormat.PictureFillFormat.Picture.Image = image;
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);
        PackageAssert.NoInternalMarkerAttributes(package, "ppt/slides/slide1.xml");
        PackageAssert.AllRelationshipReferencesResolve(package);
    }
}
