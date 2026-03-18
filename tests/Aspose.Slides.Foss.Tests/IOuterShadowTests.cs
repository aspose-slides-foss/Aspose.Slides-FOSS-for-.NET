using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Effects;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests for the OuterShadow class backed by XML elements.
/// </summary>
public sealed class IOuterShadowTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static OuterShadow CreateShadow(string attributes = "")
    {
        var xml = $"<a:outerShdw xmlns:a=\"{ANs}\" {attributes}/>";
        var element = XElement.Parse(xml);
        return new OuterShadow(element);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void BlurRadius_SetAndReadBack()
    {
        var shadow = CreateShadow();
        shadow.BlurRadius = 10;
        shadow.BlurRadius.Should().Be(10);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void Direction_SetAndReadBack()
    {
        var shadow = CreateShadow();
        shadow.Direction = 315;
        shadow.Direction.Should().Be(315);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void Distance_SetAndReadBack()
    {
        var shadow = CreateShadow();
        shadow.Distance = 8;
        shadow.Distance.Should().Be(8);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void MultipleProperties_SetTogether()
    {
        var shadow = CreateShadow();
        shadow.BlurRadius = 10;
        shadow.Direction = 315;
        shadow.Distance = 8;

        shadow.BlurRadius.Should().Be(10);
        shadow.Direction.Should().Be(315);
        shadow.Distance.Should().Be(8);
    }

    [Fact]
    public void OuterShadow_IsImageTransformOperation()
    {
        var shadow = CreateShadow();
        shadow.Should().BeAssignableTo<IImageTransformOperation>();
    }

    [Fact]
    public void AsIImageTransformOperation_ReturnsSelf()
    {
        var shadow = CreateShadow();
        shadow.AsIImageTransformOperation.Should().BeSameAs(shadow);
    }

    [Fact]
    public void BlurRadius_DefaultIsZero()
    {
        var shadow = CreateShadow();
        shadow.BlurRadius.Should().Be(0);
    }

    [Fact]
    public void Direction_DefaultIsZero()
    {
        var shadow = CreateShadow();
        shadow.Direction.Should().Be(0);
    }

    [Fact]
    public void Distance_DefaultIsZero()
    {
        var shadow = CreateShadow();
        shadow.Distance.Should().Be(0);
    }

    [Fact]
    public void SkewHorizontal_DefaultIsZero()
    {
        var shadow = CreateShadow();
        shadow.SkewHorizontal.Should().Be(0);
    }

    [Fact]
    public void SkewVertical_DefaultIsZero()
    {
        var shadow = CreateShadow();
        shadow.SkewVertical.Should().Be(0);
    }

    [Fact]
    public void ScaleHorizontal_DefaultIs100()
    {
        var shadow = CreateShadow();
        shadow.ScaleHorizontal.Should().Be(100);
    }

    [Fact]
    public void ScaleVertical_DefaultIs100()
    {
        var shadow = CreateShadow();
        shadow.ScaleVertical.Should().Be(100);
    }

    [Fact]
    public void RectangleAlign_DefaultIsBottom()
    {
        var shadow = CreateShadow();
        shadow.RectangleAlign.Should().Be(RectangleAlignment.Bottom);
    }

    [Fact]
    public void RectangleAlign_SetAndReadBack()
    {
        var shadow = CreateShadow();
        shadow.RectangleAlign = RectangleAlignment.Center;
        shadow.RectangleAlign.Should().Be(RectangleAlignment.Center);
    }

    [Fact]
    public void RectangleAlign_ReadsOoxmlAbbreviations()
    {
        var shadow = CreateShadow("algn=\"tl\"");
        shadow.RectangleAlign.Should().Be(RectangleAlignment.TopLeft);
    }

    [Fact]
    public void RectangleAlign_NotDefined_RemovesAttribute()
    {
        var shadow = CreateShadow("algn=\"ctr\"");
        shadow.RectangleAlign = RectangleAlignment.NotDefined;
        shadow.RectangleAlign.Should().Be(RectangleAlignment.Bottom); // no attribute -> default
    }

    [Fact]
    public void SkewHorizontal_SetAndReadBack()
    {
        var shadow = CreateShadow();
        shadow.SkewHorizontal = 45.5;
        shadow.SkewHorizontal.Should().BeApproximately(45.5, 0.01);
    }

    [Fact]
    public void SkewVertical_SetAndReadBack()
    {
        var shadow = CreateShadow();
        shadow.SkewVertical = -30.0;
        shadow.SkewVertical.Should().BeApproximately(-30.0, 0.01);
    }

    [Fact]
    public void RotateShadowWithShape_DefaultIsTrue()
    {
        var shadow = CreateShadow();
        shadow.RotateShadowWithShape.Should().BeTrue();
    }

    [Fact]
    public void RotateShadowWithShape_SetFalse()
    {
        var shadow = CreateShadow();
        shadow.RotateShadowWithShape = false;
        shadow.RotateShadowWithShape.Should().BeFalse();
    }

    [Fact]
    public void RotateShadowWithShape_SetTrue()
    {
        var shadow = CreateShadow();
        shadow.RotateShadowWithShape = false;
        shadow.RotateShadowWithShape = true;
        shadow.RotateShadowWithShape.Should().BeTrue();
    }

    [Fact]
    public void ScaleHorizontal_SetAndReadBack()
    {
        var shadow = CreateShadow();
        shadow.ScaleHorizontal = 50;
        shadow.ScaleHorizontal.Should().Be(50);
    }

    [Fact]
    public void ScaleVertical_SetAndReadBack()
    {
        var shadow = CreateShadow();
        shadow.ScaleVertical = 200;
        shadow.ScaleVertical.Should().Be(200);
    }

    /// <summary>
    /// Verifies that skew and scale use distinct XML attributes (kx/ky vs sx/sy).
    /// </summary>
    [Fact]
    public void SkewAndScale_UseDistinctAttributes()
    {
        var shadow = CreateShadow();
        shadow.SkewHorizontal = 10;
        shadow.ScaleHorizontal = 50;

        // They should not interfere with each other
        shadow.SkewHorizontal.Should().BeApproximately(10, 0.01);
        shadow.ScaleHorizontal.Should().Be(50);
    }
}
