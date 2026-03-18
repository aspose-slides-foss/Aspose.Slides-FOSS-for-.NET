using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Internal;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests for <see cref="TextFrameFormat"/> properties, including round-trip persistence
/// </summary>
public sealed class TextFrameFormatTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

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

    private static ITextFrameFormat GetTextFrameFormat(IShapeCollection shapes, int index = 0)
    {
        var shape = (IAutoShape)shapes[index];
        return shape.TextFrame!.TextFrameFormat;
    }

    // --- Standalone (detached) property tests ---

    [Fact]
    public void DefaultMargins_ReturnExpectedValues()
    {
        var fmt = new TextFrameFormat();
        // Default left/right = 91440 EMU / 12700 = 7.2 pts
        fmt.MarginLeft.Should().BeApproximately(7.2f, 0.01f);
        fmt.MarginRight.Should().BeApproximately(7.2f, 0.01f);
        // Default top/bottom = 45720 EMU / 12700 = 3.6 pts
        fmt.MarginTop.Should().BeApproximately(3.6f, 0.01f);
        fmt.MarginBottom.Should().BeApproximately(3.6f, 0.01f);
    }

    [Fact]
    public void MarginLeft_SetAndGet()
    {
        var fmt = new TextFrameFormat();
        fmt.MarginLeft = 10f;
        fmt.MarginLeft.Should().BeApproximately(10f, 0.1f);
    }

    [Fact]
    public void WrapText_DefaultIsNotDefined()
    {
        var fmt = new TextFrameFormat();
        fmt.WrapText.Should().Be(NullableBool.NotDefined);
    }

    [Fact]
    public void WrapText_SetTrueAndFalse()
    {
        var fmt = new TextFrameFormat();
        fmt.WrapText = NullableBool.True;
        fmt.WrapText.Should().Be(NullableBool.True);

        fmt.WrapText = NullableBool.False;
        fmt.WrapText.Should().Be(NullableBool.False);

        fmt.WrapText = NullableBool.NotDefined;
        fmt.WrapText.Should().Be(NullableBool.NotDefined);
    }

    [Fact]
    public void AnchoringType_DefaultIsNotDefined()
    {
        var fmt = new TextFrameFormat();
        fmt.AnchoringType.Should().Be(TextAnchorType.NotDefined);
    }

    [Theory]
    [InlineData(TextAnchorType.Top)]
    [InlineData(TextAnchorType.Center)]
    [InlineData(TextAnchorType.Bottom)]
    [InlineData(TextAnchorType.Justified)]
    [InlineData(TextAnchorType.Distributed)]
    public void AnchoringType_SetAndGet(TextAnchorType expected)
    {
        var fmt = new TextFrameFormat();
        fmt.AnchoringType = expected;
        fmt.AnchoringType.Should().Be(expected);
    }

    [Fact]
    public void CenterText_SetAndGet()
    {
        var fmt = new TextFrameFormat();
        fmt.CenterText = NullableBool.True;
        fmt.CenterText.Should().Be(NullableBool.True);

        fmt.CenterText = NullableBool.False;
        fmt.CenterText.Should().Be(NullableBool.False);
    }

    [Theory]
    [InlineData(TextVerticalType.Horizontal)]
    [InlineData(TextVerticalType.Vertical)]
    [InlineData(TextVerticalType.Vertical270)]
    public void TextVerticalType_SetAndGet(TextVerticalType expected)
    {
        var fmt = new TextFrameFormat();
        fmt.TextVerticalType = expected;
        fmt.TextVerticalType.Should().Be(expected);
    }

    [Theory]
    [InlineData(TextAutofitType.None)]
    [InlineData(TextAutofitType.Normal)]
    [InlineData(TextAutofitType.Shape)]
    public void AutofitType_SetAndGet(TextAutofitType expected)
    {
        var fmt = new TextFrameFormat();
        fmt.AutofitType = expected;
        fmt.AutofitType.Should().Be(expected);
    }

    [Fact]
    public void ColumnCount_DefaultIsOne()
    {
        var fmt = new TextFrameFormat();
        fmt.ColumnCount.Should().Be(1);
    }

    [Fact]
    public void ColumnCount_SetAndGet()
    {
        var fmt = new TextFrameFormat();
        fmt.ColumnCount = 3;
        fmt.ColumnCount.Should().Be(3);
    }

    [Fact]
    public void ColumnSpacing_DefaultIsZero()
    {
        var fmt = new TextFrameFormat();
        fmt.ColumnSpacing.Should().Be(0f);
    }

    [Fact]
    public void ColumnSpacing_SetAndGet()
    {
        var fmt = new TextFrameFormat();
        fmt.ColumnSpacing = 15f;
        fmt.ColumnSpacing.Should().BeApproximately(15f, 0.1f);
    }

    [Fact]
    public void RotationAngle_SetAndGet()
    {
        var fmt = new TextFrameFormat();
        fmt.RotationAngle = 45f;
        fmt.RotationAngle.Should().BeApproximately(45f, 0.01f);
    }

    [Theory]
    [InlineData(TextShapeType.None)]
    [InlineData(TextShapeType.Plain)]
    [InlineData(TextShapeType.Wave1)]
    public void Transform_SetAndGet(TextShapeType expected)
    {
        var fmt = new TextFrameFormat();
        fmt.Transform = expected;
        fmt.Transform.Should().Be(expected);
    }

    [Fact]
    public void KeepTextFlat_SetAndGet()
    {
        var fmt = new TextFrameFormat();
        fmt.KeepTextFlat = true;
        fmt.KeepTextFlat.Should().BeTrue();
    }

    // --- Slide property tests ---

    [Fact]
    public void Slide_DefaultIsNull_WhenDetached()
    {
        var fmt = new TextFrameFormat();
        fmt.Slide.Should().BeNull();
    }

    [Fact]
    public void Slide_ReturnsParentSlide_WhenBoundToShape()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        var fmt = shape.TextFrame!.TextFrameFormat;
        // When bound via a ShapeCollection with a parent slide, Slide should reflect that.
        // In this test helper, parent slide is null, so we verify the property works without error.
        // The property returns whatever _parentSlide was set to via InitInternal.
        fmt.Slide.Should().BeNull();
    }

    [Fact]
    public void Slide_ReturnsParentSlide_WhenInitializedWithSlide()
    {
        var slidePart = new SlidePart();
        slidePart.InitInternal("ppt/slides/slide1.xml");
        slidePart.Element = XDocument.Parse(SlideXml).Root;

        var mockSlide = new Slide();
        var shapes = new ShapeCollection();
        shapes.InitInternal(slidePart, mockSlide);

        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        var fmt = shape.TextFrame!.TextFrameFormat;
        fmt.Slide.Should().BeSameAs(mockSlide);
    }

    // --- Round-trip integration tests (via shape TextFrame) ---

    [Fact]
    public void Margins_PersistAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        var fmt = shape.TextFrame!.TextFrameFormat;

        fmt.MarginLeft = 10f;
        fmt.MarginRight = 12f;
        fmt.MarginTop = 5f;
        fmt.MarginBottom = 8f;

        var reloadedFmt = GetTextFrameFormat(RoundTripShapes(slidePart));

        reloadedFmt.MarginLeft.Should().BeApproximately(10f, 0.1f);
        reloadedFmt.MarginRight.Should().BeApproximately(12f, 0.1f);
        reloadedFmt.MarginTop.Should().BeApproximately(5f, 0.1f);
        reloadedFmt.MarginBottom.Should().BeApproximately(8f, 0.1f);
    }

    [Fact]
    public void WrapText_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        shape.TextFrame!.TextFrameFormat.WrapText = NullableBool.True;

        var reloadedFmt = GetTextFrameFormat(RoundTripShapes(slidePart));
        reloadedFmt.WrapText.Should().Be(NullableBool.True);
    }

    [Fact]
    public void AnchoringType_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        shape.TextFrame!.TextFrameFormat.AnchoringType = TextAnchorType.Center;

        var reloadedFmt = GetTextFrameFormat(RoundTripShapes(slidePart));
        reloadedFmt.AnchoringType.Should().Be(TextAnchorType.Center);
    }

    [Fact]
    public void TextVerticalType_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        shape.TextFrame!.TextFrameFormat.TextVerticalType = TextVerticalType.Vertical;

        var reloadedFmt = GetTextFrameFormat(RoundTripShapes(slidePart));
        reloadedFmt.TextVerticalType.Should().Be(TextVerticalType.Vertical);
    }

    [Fact]
    public void AutofitType_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        shape.TextFrame!.TextFrameFormat.AutofitType = TextAutofitType.Shape;

        var reloadedFmt = GetTextFrameFormat(RoundTripShapes(slidePart));
        reloadedFmt.AutofitType.Should().Be(TextAutofitType.Shape);
    }

    [Fact]
    public void RotationAngle_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        shape.TextFrame!.TextFrameFormat.RotationAngle = 90f;

        var reloadedFmt = GetTextFrameFormat(RoundTripShapes(slidePart));
        reloadedFmt.RotationAngle.Should().BeApproximately(90f, 0.01f);
    }

    [Fact]
    public void Transform_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 100);
        shape.TextFrame!.TextFrameFormat.Transform = TextShapeType.Wave1;

        var reloadedFmt = GetTextFrameFormat(RoundTripShapes(slidePart));
        reloadedFmt.Transform.Should().Be(TextShapeType.Wave1);
    }


    [Fact]
    public void ThreeDFormat_BevelTop_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var tdf = shape.TextFrame!.TextFrameFormat.ThreeDFormat;

        tdf.BevelTop.BevelType = BevelPresetType.Circle;
        tdf.BevelTop.Width = 10;
        tdf.BevelTop.Height = 5;

        var reloadedFmt = GetTextFrameFormat(RoundTripShapes(slidePart));
        var bt = reloadedFmt.ThreeDFormat.BevelTop;

        bt.BevelType.Should().Be(BevelPresetType.Circle);
        bt.Width.Should().BeApproximately(10f, 0.1f);
        bt.Height.Should().BeApproximately(5f, 0.1f);
    }

    [Fact]
    public void ThreeDFormat_Camera_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        shape.TextFrame!.TextFrameFormat.ThreeDFormat.Camera.CameraType = CameraPresetType.PerspectiveAbove;

        var reloadedFmt = GetTextFrameFormat(RoundTripShapes(slidePart));
        reloadedFmt.ThreeDFormat.Camera.CameraType.Should().Be(CameraPresetType.PerspectiveAbove);
    }

    [Fact]
    public void ThreeDFormat_LightRig_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var lr = shape.TextFrame!.TextFrameFormat.ThreeDFormat.LightRig;
        lr.LightType = LightRigPresetType.Balanced;
        lr.Direction = LightingDirection.Top;

        var reloadedFmt = GetTextFrameFormat(RoundTripShapes(slidePart));
        var lr2 = reloadedFmt.ThreeDFormat.LightRig;
        lr2.LightType.Should().Be(LightRigPresetType.Balanced);
        lr2.Direction.Should().Be(LightingDirection.Top);
    }

    [Fact]
    public void ThreeDFormat_DepthAndMaterial_PersistAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var shape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var tdf = shape.TextFrame!.TextFrameFormat.ThreeDFormat;
        tdf.Depth = 20;
        tdf.Material = MaterialPresetType.Metal;

        var reloadedFmt = GetTextFrameFormat(RoundTripShapes(slidePart));
        reloadedFmt.ThreeDFormat.Depth.Should().BeApproximately(20f, 0.1f);
        reloadedFmt.ThreeDFormat.Material.Should().Be(MaterialPresetType.Metal);
    }
}
