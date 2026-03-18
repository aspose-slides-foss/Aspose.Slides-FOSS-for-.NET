using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.IntegrationTests;

/// <summary>
/// Tests for ImageCollection and PictureFrame.
/// </summary>
public sealed class ImagesTests : IDisposable
{
    private readonly string _tempDir;

    public ImagesTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, recursive: true);
    }

    [Fact]
    public void TestAddImage()
    {
        // Adding an image increases collection count.
        using var pres = new Presentation();
        pres.Images.AddImage(TestHelpers.CreateTestPng(255, 0, 0));
        pres.Images.Count.Should().BeGreaterThanOrEqualTo(1);
    }

    [Fact]
    public void TestMultipleImages()
    {
        // Multiple images can be added and iterated.
        using var pres = new Presentation();
        var colors = new (byte r, byte g, byte b)[] { (255, 0, 0), (0, 255, 0), (0, 0, 255) };
        foreach (var (r, g, b) in colors)
        {
            pres.Images.AddImage(TestHelpers.CreateTestPng(r, g, b));
        }

        pres.Images.Count.Should().BeGreaterThanOrEqualTo(3);
        var imgs = pres.Images.AsIEnumerable.ToList();
        imgs.Count.Should().BeGreaterThanOrEqualTo(3);
    }

    [Fact]
    public void TestPictureFrame()
    {
        // Picture frame with image persists after save/reload.
        using var pres = new Presentation();
        var img = pres.Images.AddImage(TestHelpers.CreateTestPng(0, 0, 255));
        pres.Slides[0].Shapes!.AddPictureFrame(ShapeType.Rectangle, 50, 50, 100, 100, img);
        pres.Slides[0].Shapes!.Count.Should().BeGreaterThanOrEqualTo(1);

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        pres2.Slides[0].Shapes!.Count.Should().BeGreaterThanOrEqualTo(1);
    }

    [Fact]
    public void TestImageFromFile()
    {
        // Load an image from the test_data directory.
        var imgPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test_data", "lotus.png");

        using var pres = new Presentation();
        var imageData = File.ReadAllBytes(imgPath);
        var ppImg = pres.Images.AddImage(imageData);
        pres.Slides[0].Shapes!.AddPictureFrame(ShapeType.Rectangle, 50, 50, 200, 200, ppImg);
        pres.Slides[0].Shapes!.Count.Should().BeGreaterThanOrEqualTo(1);
    }
}
