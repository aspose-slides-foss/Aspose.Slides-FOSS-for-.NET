using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.IntegrationTests;

/// <summary>
/// Tests for ThreeDFormat: bevel, camera, light rig, depth.
/// </summary>
public sealed class ThreeDFormatTests : IDisposable
{
    private readonly string _tempDir;

    public ThreeDFormatTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, recursive: true);
    }

    /// <summary>Return a slide with all placeholder shapes removed.</summary>
    private static ISlide Clear(Presentation pres)
    {
        var slide = pres.Slides[0];
        slide.Shapes!.Clear();
        return slide;
    }

    [Fact]
    public void TestBevelTop()
    {
        // Bevel top type/width/height persist.
        using var pres = new Presentation();
        var slide = Clear(pres);
        var shape = slide.Shapes!.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var tdf = shape.ThreeDFormat;
        tdf.BevelTop.BevelType = BevelPresetType.Circle;
        tdf.BevelTop.Width = 10;
        tdf.BevelTop.Height = 5;

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var bt = pres2.Slides[0].Shapes![0].ThreeDFormat.BevelTop;
        bt.BevelType.Should().Be(BevelPresetType.Circle);
        bt.Width.Should().Be(10);
        bt.Height.Should().Be(5);
    }

    [Fact]
    public void TestCamera()
    {
        // Camera preset persists.
        using var pres = new Presentation();
        var slide = Clear(pres);
        var shape = slide.Shapes!.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        shape.ThreeDFormat.Camera.CameraType = CameraPresetType.PerspectiveAbove;

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var cam = pres2.Slides[0].Shapes![0].ThreeDFormat.Camera;
        cam.CameraType.Should().Be(CameraPresetType.PerspectiveAbove);
    }

    [Fact]
    public void TestLightRig()
    {
        // Light rig preset and direction persist.
        using var pres = new Presentation();
        var slide = Clear(pres);
        var shape = slide.Shapes!.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var lr = shape.ThreeDFormat.LightRig;
        lr.LightType = LightRigPresetType.Balanced;
        lr.Direction = LightingDirection.Top;

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var lr2 = pres2.Slides[0].Shapes![0].ThreeDFormat.LightRig;
        lr2.LightType.Should().Be(LightRigPresetType.Balanced);
        lr2.Direction.Should().Be(LightingDirection.Top);
    }

    [Fact]
    public void TestDepthAndMaterial()
    {
        // Extrusion depth and material persist.
        using var pres = new Presentation();
        var slide = Clear(pres);
        var shape = slide.Shapes!.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var tdf = shape.ThreeDFormat;
        tdf.Depth = 20;
        tdf.Material = MaterialPresetType.Metal;

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var tdf2 = pres2.Slides[0].Shapes![0].ThreeDFormat;
        tdf2.Depth.Should().Be(20);
        tdf2.Material.Should().Be(MaterialPresetType.Metal);
    }
}
