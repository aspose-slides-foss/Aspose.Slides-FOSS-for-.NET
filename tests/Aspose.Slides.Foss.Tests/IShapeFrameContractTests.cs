using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests the IShapeFrame contract via the ShapeFrame implementation.
/// </summary>
public sealed class IShapeFrameContractTests
{
    private static ShapeFrame CreateFrame(
        float x = 200, float y = 200,
        float width = 300, float height = 250,
        NullableBool flipH = NullableBool.False,
        NullableBool flipV = NullableBool.False,
        float rotation = 0) =>
        new ShapeFrame(x, y, width, height, flipH, flipV, rotation);

    /// <summary>
    /// x, y, width, height persist after construction.
    /// </summary>
    [Fact]
    public void Constructor_XYWidthHeightArePreserved()
    {
        var frame = CreateFrame(x: 200, y: 200, width: 300, height: 250);

        frame.X.Should().Be(200);
        frame.Y.Should().Be(200);
        frame.Width.Should().Be(300);
        frame.Height.Should().Be(250);
    }

    /// <summary>
    /// Rotation persists after construction.
    /// </summary>
    [Fact]
    public void Constructor_RotationIsPreserved()
    {
        var frame = CreateFrame(rotation: 45);

        frame.Rotation.Should().Be(45);
    }

    /// <summary>
    /// All frame properties persist together.
    /// </summary>
    [Fact]
    public void Constructor_AllPropertiesPersistTogether()
    {
        var frame = new ShapeFrame(200, 200, 300, 250, NullableBool.False, NullableBool.False, 45);

        frame.X.Should().Be(200);
        frame.Y.Should().Be(200);
        frame.Width.Should().Be(300);
        frame.Height.Should().Be(250);
        frame.Rotation.Should().Be(45);
    }

    /// <summary>
    /// ShapeFrame implements IShapeFrame interface.
    /// </summary>
    [Fact]
    public void ShapeFrame_ImplementsIShapeFrame()
    {
        var frame = CreateFrame();

        frame.Should().BeAssignableTo<IShapeFrame>();
    }

    /// <summary>
    /// CenterX is computed as X + Width / 2.
    /// </summary>
    [Fact]
    public void CenterX_IsComputedFromXAndWidth()
    {
        var frame = CreateFrame(x: 100, width: 200);

        frame.CenterX.Should().Be(200); // 100 + 200/2
    }

    /// <summary>
    /// CenterY is computed as Y + Height / 2.
    /// </summary>
    [Fact]
    public void CenterY_IsComputedFromYAndHeight()
    {
        var frame = CreateFrame(y: 100, height: 200);

        frame.CenterY.Should().Be(200); // 100 + 200/2
    }

    /// <summary>
    /// CenterX and CenterY for the standard test frame (200, 200, 300, 250).
    /// </summary>
    [Fact]
    public void Center_MatchesExpectedForStandardFrame()
    {
        var frame = CreateFrame(x: 200, y: 200, width: 300, height: 250);

        frame.CenterX.Should().Be(350); // 200 + 300/2
        frame.CenterY.Should().Be(325); // 200 + 250/2
    }

    /// <summary>
    /// FlipH is preserved after construction.
    /// </summary>
    [Theory]
    [InlineData(NullableBool.False)]
    [InlineData(NullableBool.True)]
    [InlineData(NullableBool.NotDefined)]
    public void FlipH_IsPreserved(NullableBool flipH)
    {
        var frame = CreateFrame(flipH: flipH);

        frame.FlipH.Should().Be(flipH);
    }

    /// <summary>
    /// FlipV is preserved after construction.
    /// </summary>
    [Theory]
    [InlineData(NullableBool.False)]
    [InlineData(NullableBool.True)]
    [InlineData(NullableBool.NotDefined)]
    public void FlipV_IsPreserved(NullableBool flipV)
    {
        var frame = CreateFrame(flipV: flipV);

        frame.FlipV.Should().Be(flipV);
    }

    /// <summary>
    /// Rectangle returns a tuple of (X, Y, Width, Height).
    /// </summary>
    [Fact]
    public void Rectangle_ReturnsTupleOfCoordinates()
    {
        var frame = CreateFrame(x: 200, y: 200, width: 300, height: 250);

        var rect = frame.Rectangle;
        rect.Should().NotBeNull();
        rect.Should().Be((200f, 200f, 300f, 250f));
    }

    /// <summary>
    /// CloneT creates a deep copy with the same property values.
    /// </summary>
    [Fact]
    public void CloneT_CreatesDeepCopyWithSameValues()
    {
        var frame = new ShapeFrame(200, 200, 300, 250, NullableBool.True, NullableBool.False, 45);

        var clone = frame.CloneT();

        clone.X.Should().Be(frame.X);
        clone.Y.Should().Be(frame.Y);
        clone.Width.Should().Be(frame.Width);
        clone.Height.Should().Be(frame.Height);
        clone.FlipH.Should().Be(frame.FlipH);
        clone.FlipV.Should().Be(frame.FlipV);
        clone.Rotation.Should().Be(frame.Rotation);
    }

    /// <summary>
    /// CloneT returns a different instance.
    /// </summary>
    [Fact]
    public void CloneT_ReturnsDifferentInstance()
    {
        var frame = CreateFrame();

        var clone = frame.CloneT();

        clone.Should().NotBeSameAs(frame);
    }

    /// <summary>
    /// CloneT preserves computed properties (CenterX, CenterY).
    /// </summary>
    [Fact]
    public void CloneT_PreservesComputedProperties()
    {
        var frame = CreateFrame(x: 100, y: 50, width: 400, height: 300);

        var clone = frame.CloneT();

        clone.CenterX.Should().Be(frame.CenterX);
        clone.CenterY.Should().Be(frame.CenterY);
    }

    /// <summary>
    /// Zero-dimension frame is valid.
    /// </summary>
    [Fact]
    public void Constructor_ZeroDimensionsAreValid()
    {
        var frame = CreateFrame(x: 0, y: 0, width: 0, height: 0, rotation: 0);

        frame.X.Should().Be(0);
        frame.Y.Should().Be(0);
        frame.Width.Should().Be(0);
        frame.Height.Should().Be(0);
        frame.CenterX.Should().Be(0);
        frame.CenterY.Should().Be(0);
    }

    /// <summary>
    /// A connector-like frame with positive width or height is valid.
    /// </summary>
    [Fact]
    public void Constructor_PositiveWidthOrHeightIsValid()
    {
        var frame = CreateFrame(x: 0, y: 0, width: 350, height: 0);

        (frame.Width > 0 || frame.Height > 0).Should().BeTrue();
    }

    /// <summary>
    /// Negative rotation values are supported.
    /// </summary>
    [Fact]
    public void Constructor_NegativeRotationIsSupported()
    {
        var frame = CreateFrame(rotation: -30);

        frame.Rotation.Should().Be(-30);
    }

    /// <summary>
    /// Clone (ICloneable) creates a deep copy with the same property values.
    /// </summary>
    [Fact]
    public void Clone_CreatesDeepCopyWithSameValues()
    {
        var frame = new ShapeFrame(200, 200, 300, 250, NullableBool.True, NullableBool.False, 45);

        var clone = (ShapeFrame)frame.Clone();

        clone.X.Should().Be(frame.X);
        clone.Y.Should().Be(frame.Y);
        clone.Width.Should().Be(frame.Width);
        clone.Height.Should().Be(frame.Height);
        clone.FlipH.Should().Be(frame.FlipH);
        clone.FlipV.Should().Be(frame.FlipV);
        clone.Rotation.Should().Be(frame.Rotation);
    }

    /// <summary>
    /// Clone returns a different instance.
    /// </summary>
    [Fact]
    public void Clone_ReturnsDifferentInstance()
    {
        var frame = CreateFrame();

        var clone = frame.Clone();

        clone.Should().NotBeSameAs(frame);
    }

    /// <summary>
    /// Equals returns true for frames with identical field values.
    /// </summary>
    [Fact]
    public void Equals_TrueForIdenticalFrames()
    {
        var a = new ShapeFrame(100, 200, 300, 400, NullableBool.True, NullableBool.False, 45);
        var b = new ShapeFrame(100, 200, 300, 400, NullableBool.True, NullableBool.False, 45);

        a.Equals(b).Should().BeTrue();
    }

    /// <summary>
    /// Equals returns false when any field differs.
    /// </summary>
    [Fact]
    public void Equals_FalseWhenFieldsDiffer()
    {
        var frame = new ShapeFrame(100, 200, 300, 400, NullableBool.True, NullableBool.False, 45);

        frame.Equals(new ShapeFrame(999, 200, 300, 400, NullableBool.True, NullableBool.False, 45)).Should().BeFalse();
        frame.Equals(new ShapeFrame(100, 999, 300, 400, NullableBool.True, NullableBool.False, 45)).Should().BeFalse();
        frame.Equals(new ShapeFrame(100, 200, 999, 400, NullableBool.True, NullableBool.False, 45)).Should().BeFalse();
        frame.Equals(new ShapeFrame(100, 200, 300, 999, NullableBool.True, NullableBool.False, 45)).Should().BeFalse();
        frame.Equals(new ShapeFrame(100, 200, 300, 400, NullableBool.False, NullableBool.False, 45)).Should().BeFalse();
        frame.Equals(new ShapeFrame(100, 200, 300, 400, NullableBool.True, NullableBool.True, 45)).Should().BeFalse();
        frame.Equals(new ShapeFrame(100, 200, 300, 400, NullableBool.True, NullableBool.False, 90)).Should().BeFalse();
    }

    /// <summary>
    /// Equals returns false for null and non-ShapeFrame objects.
    /// </summary>
    [Fact]
    public void Equals_FalseForNullAndOtherTypes()
    {
        var frame = CreateFrame();

        frame!.Equals(null).Should().BeFalse();
        frame!.Equals((object)"not a frame").Should().BeFalse();
    }

    /// <summary>
    /// GetHashCode is consistent for equal frames.
    /// </summary>
    [Fact]
    public void GetHashCode_ConsistentForEqualFrames()
    {
        var a = new ShapeFrame(100, 200, 300, 400, NullableBool.True, NullableBool.False, 45);
        var b = new ShapeFrame(100, 200, 300, 400, NullableBool.True, NullableBool.False, 45);

        a.GetHashCode().Should().Be(b.GetHashCode());
    }

    /// <summary>
    /// A clone is equal to the original.
    /// </summary>
    [Fact]
    public void Equals_CloneIsEqualToOriginal()
    {
        var frame = new ShapeFrame(50, 100, 200, 150, NullableBool.True, NullableBool.True, 30);

        var clone = (ShapeFrame)frame.Clone();

        frame.Equals(clone).Should().BeTrue();
    }

    /// <summary>
    /// Various rotation angles via Theory.
    /// </summary>
    [Theory]
    [InlineData(0)]
    [InlineData(45)]
    [InlineData(90)]
    [InlineData(180)]
    [InlineData(270)]
    [InlineData(360)]
    [InlineData(-45)]
    public void Constructor_VariousRotationAngles(float rotation)
    {
        var frame = CreateFrame(rotation: rotation);

        frame.Rotation.Should().Be(rotation);
    }

    /// <summary>
    /// Frame properties with typical shape coordinates.
    /// </summary>
    [Theory]
    [InlineData(50, 50, 200, 100)]
    [InlineData(200, 200, 300, 250)]
    [InlineData(0, 0, 1, 1)]
    [InlineData(100, 100, 80, 80)]
    public void Constructor_VariousCoordinates(float x, float y, float width, float height)
    {
        var frame = CreateFrame(x: x, y: y, width: width, height: height);

        frame.X.Should().Be(x);
        frame.Y.Should().Be(y);
        frame.Width.Should().Be(width);
        frame.Height.Should().Be(height);
    }
}
