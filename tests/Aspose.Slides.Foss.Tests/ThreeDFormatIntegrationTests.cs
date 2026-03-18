using System.Text;
using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Internal;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Verifies that 3-D formatting properties persist through XML round-trip cycles.
/// </summary>
public sealed class ThreeDFormatIntegrationTests
{
    private static readonly string SlideXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
               xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:cSld>
            <p:spTree>
              <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
              <p:grpSpPr/>
            </p:spTree>
          </p:cSld>
        </p:sld>
        """;

    private static (SlidePart slidePart, ShapeCollection shapes) CreateSlideWithShapes()
    {
        var slidePart = new SlidePart();
        slidePart.InitInternal("ppt/slides/slide1.xml");
        slidePart.Element = XDocument.Parse(SlideXml).Root;

        var shapes = new ShapeCollection();
        shapes.InitInternal(slidePart, null);
        return (slidePart, shapes);
    }

    private static ShapeCollection RoundTripShapes(SlidePart slidePart)
    {
        var xml = slidePart.Element!.ToString();
        var reloadedRoot = XDocument.Parse(xml).Root!;

        var newSlidePart = new SlidePart();
        newSlidePart.InitInternal("ppt/slides/slide1.xml");
        newSlidePart.Element = reloadedRoot;

        var shapes = new ShapeCollection();
        shapes.InitInternal(newSlidePart, null);
        return shapes;
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void BevelTop_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);

        shape.ThreeDFormat.BevelTop.BevelType = BevelPresetType.Circle;
        shape.ThreeDFormat.BevelTop.Width = 10;
        shape.ThreeDFormat.BevelTop.Height = 5;

        var reloadedShapes = RoundTripShapes(slidePart);
        var reloadedShape = reloadedShapes[0];

        reloadedShape.ThreeDFormat.BevelTop.BevelType.Should().Be(BevelPresetType.Circle);
        reloadedShape.ThreeDFormat.BevelTop.Width.Should().BeApproximately(10f, 0.1f);
        reloadedShape.ThreeDFormat.BevelTop.Height.Should().BeApproximately(5f, 0.1f);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void Camera_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);

        shape.ThreeDFormat.Camera.CameraType = CameraPresetType.PerspectiveAbove;

        var reloadedShapes = RoundTripShapes(slidePart);
        var reloadedShape = reloadedShapes[0];

        reloadedShape.ThreeDFormat.Camera.CameraType.Should().Be(CameraPresetType.PerspectiveAbove);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void LightRig_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);

        shape.ThreeDFormat.LightRig.LightType = LightRigPresetType.Balanced;
        shape.ThreeDFormat.LightRig.Direction = LightingDirection.Top;

        var reloadedShapes = RoundTripShapes(slidePart);
        var reloadedShape = reloadedShapes[0];

        reloadedShape.ThreeDFormat.LightRig.LightType.Should().Be(LightRigPresetType.Balanced);
        reloadedShape.ThreeDFormat.LightRig.Direction.Should().Be(LightingDirection.Top);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void DepthAndMaterial_PersistAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);

        shape.ThreeDFormat.Depth = 20;
        shape.ThreeDFormat.Material = MaterialPresetType.Metal;

        var reloadedShapes = RoundTripShapes(slidePart);
        var reloadedShape = reloadedShapes[0];

        reloadedShape.ThreeDFormat.Depth.Should().BeApproximately(20f, 0.1f);
        reloadedShape.ThreeDFormat.Material.Should().Be(MaterialPresetType.Metal);
    }
}
