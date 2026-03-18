using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.IntegrationTests;

/// <summary>
/// Tests for ShapeCollection operations and shape frame properties.
/// </summary>
public sealed class ShapesTests : IDisposable
{
    private readonly string _tempDir;

    public ShapesTests()
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
    private static ISlide BlankSlide(Presentation pres)
    {
        var slide = pres.Slides[0];
        slide.Shapes!.Clear();
        return slide;
    }

    [Fact]
    public void TestAddAutoShape()
    {
        // add_auto_shape adds a rectangle with correct type.
        using var pres = new Presentation();
        var slide = BlankSlide(pres);
        var shape = slide.Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        slide.Shapes!.Count.Should().Be(1);
        shape.ShapeType.Should().Be(ShapeType.Rectangle);
    }

    [Fact]
    public void TestMultipleShapeTypes()
    {
        // Various ShapeType values are preserved.
        ShapeType[] types = [ShapeType.Rectangle, ShapeType.Ellipse, ShapeType.Triangle];
        using var pres = new Presentation();
        var slide = BlankSlide(pres);
        foreach (var st in types)
        {
            var s = slide.Shapes!.AddAutoShape(st, 10, 10, 100, 100);
            s.ShapeType.Should().Be(st);
        }
        slide.Shapes!.Count.Should().Be(3);
    }

    [Fact]
    public void TestInsertAutoShape()
    {
        // insert_auto_shape places a shape at the requested index.
        using var pres = new Presentation();
        var slide = BlankSlide(pres);
        slide.Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        slide.Shapes!.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);
        slide.Shapes!.InsertAutoShape(1, ShapeType.Triangle, 150, 200, 100, 100);
        slide.Shapes!.Count.Should().Be(3);
    }

    [Fact]
    public void TestRemoveShape()
    {
        // Removing a shape by reference decreases count.
        using var pres = new Presentation();
        var slide = BlankSlide(pres);
        var s = slide.Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        slide.Shapes!.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);
        slide.Shapes!.Count.Should().Be(2);
        slide.Shapes!.Remove(s);
        slide.Shapes!.Count.Should().Be(1);
    }

    [Fact]
    public void TestRemoveAt()
    {
        // remove_at removes by index.
        using var pres = new Presentation();
        var slide = BlankSlide(pres);
        slide.Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        slide.Shapes!.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);
        slide.Shapes!.RemoveAt(0);
        slide.Shapes!.Count.Should().Be(1);
    }

    [Fact]
    public void TestClearShapes()
    {
        // clear() empties the shape collection.
        using var pres = new Presentation();
        var slide = pres.Slides[0];
        slide.Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        slide.Shapes!.Clear();
        slide.Shapes!.Count.Should().Be(0);
    }

    [Fact]
    public void TestShapeFrameProperties()
    {
        // x, y, width, height, rotation persist after save/reload.
        using var pres = new Presentation();
        var slide = BlankSlide(pres);
        var shape = slide.Shapes!.AddAutoShape(ShapeType.Rectangle, 200, 200, 300, 250);
        shape.Rotation = 45;

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var s2 = pres2.Slides[0].Shapes![0];
        s2.X.Should().Be(200);
        s2.Y.Should().Be(200);
        s2.Width.Should().Be(300);
        s2.Height.Should().Be(250);
        s2.Rotation.Should().Be(45);
    }

    [Fact]
    public void TestReorderShapes()
    {
        // reorder() changes the z-order of shapes.
        using var pres = new Presentation();
        var slide = BlankSlide(pres);
        slide.Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        var ellipse = slide.Shapes!.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);
        slide.Shapes!.Reorder(0, ellipse);
        ((IGeometryShape)slide.Shapes![0]).ShapeType.Should().Be(ShapeType.Ellipse);
    }

    [Fact]
    public void TestIterateShapes()
    {
        // Shapes collection is iterable.
        using var pres = new Presentation();
        var slide = BlankSlide(pres);
        slide.Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        slide.Shapes!.AddAutoShape(ShapeType.Ellipse, 300, 50, 150, 150);
        var shapes = slide.Shapes!.ToList();
        shapes.Count.Should().Be(2);
    }

    [Fact]
    public void TestShapePersistsAfterReload()
    {
        // Shapes survive a save/reload cycle.
        using var pres = new Presentation();
        var slide = BlankSlide(pres);
        slide.Shapes!.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        pres2.Slides[0].Shapes!.Count.Should().BeGreaterThanOrEqualTo(1);
        ((IGeometryShape)pres2.Slides[0].Shapes![0]).ShapeType.Should().Be(ShapeType.Rectangle);
    }
}
