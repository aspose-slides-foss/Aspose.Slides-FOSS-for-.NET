using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Effects;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests for the Reflection class backed by XML elements.
/// </summary>
public sealed class IReflectionTests
{
    private static Reflection CreateReflection(params (string name, string value)[] attributes)
    {
        var element = new XElement("reflection");
        foreach (var (name, value) in attributes)
            element.SetAttributeValue(name, value);
        return new Reflection(element);
    }

    private static Reflection CreateEmptyReflection() => new Reflection(new XElement("reflection"));

    /// <summary>
    /// StartPosAlpha reads from stPos attribute with /1000 conversion.
    /// </summary>
    [Fact]
    public void StartPosAlpha_ReadsFromStPos()
    {
        var r = CreateReflection(("stPos", "50000"));
        r.StartPosAlpha.Should().Be(50f);
    }

    /// <summary>
    /// StartPosAlpha defaults to 0 when attribute is missing.
    /// </summary>
    [Fact]
    public void StartPosAlpha_DefaultsToZero()
    {
        var r = CreateEmptyReflection();
        r.StartPosAlpha.Should().Be(0f);
    }

    /// <summary>
    /// StartPosAlpha setter writes stPos attribute with *1000 conversion.
    /// </summary>
    [Fact]
    public void StartPosAlpha_SetRoundTrips()
    {
        var r = CreateEmptyReflection();
        r.StartPosAlpha = 75f;
        r.StartPosAlpha.Should().Be(75f);
    }

    /// <summary>
    /// EndPosAlpha reads from endPos attribute and defaults to 100.
    /// </summary>
    [Fact]
    public void EndPosAlpha_DefaultsTo100()
    {
        var r = CreateEmptyReflection();
        r.EndPosAlpha.Should().Be(100f);
    }

    [Fact]
    public void EndPosAlpha_ReadsFromEndPos()
    {
        var r = CreateReflection(("endPos", "50000"));
        r.EndPosAlpha.Should().Be(50f);
    }

    [Fact]
    public void EndPosAlpha_SetRoundTrips()
    {
        var r = CreateEmptyReflection();
        r.EndPosAlpha = 80f;
        r.EndPosAlpha.Should().Be(80f);
    }

    /// <summary>
    /// FadeDirection defaults to 90 and uses 60000ths of a degree.
    /// </summary>
    [Fact]
    public void FadeDirection_DefaultsTo90()
    {
        var r = CreateEmptyReflection();
        r.FadeDirection.Should().Be(90f);
    }

    [Fact]
    public void FadeDirection_ReadsFromFadeDir()
    {
        var r = CreateReflection(("fadeDir", "2700000"));
        r.FadeDirection.Should().Be(45f);
    }

    [Fact]
    public void FadeDirection_SetRoundTrips()
    {
        var r = CreateEmptyReflection();
        r.FadeDirection = 45f;
        r.FadeDirection.Should().Be(45f);
    }

    /// <summary>
    /// StartReflectionOpacity reads from stA and defaults to 100.
    /// </summary>
    [Fact]
    public void StartReflectionOpacity_DefaultsTo100()
    {
        var r = CreateEmptyReflection();
        r.StartReflectionOpacity.Should().Be(100f);
    }

    [Fact]
    public void StartReflectionOpacity_ReadsFromStA()
    {
        var r = CreateReflection(("stA", "50000"));
        r.StartReflectionOpacity.Should().Be(50f);
    }

    [Fact]
    public void StartReflectionOpacity_SetRoundTrips()
    {
        var r = CreateEmptyReflection();
        r.StartReflectionOpacity = 75f;
        r.StartReflectionOpacity.Should().Be(75f);
    }

    /// <summary>
    /// EndReflectionOpacity reads from endA and defaults to 0.
    /// </summary>
    [Fact]
    public void EndReflectionOpacity_DefaultsToZero()
    {
        var r = CreateEmptyReflection();
        r.EndReflectionOpacity.Should().Be(0f);
    }

    [Fact]
    public void EndReflectionOpacity_ReadsFromEndA()
    {
        var r = CreateReflection(("endA", "30000"));
        r.EndReflectionOpacity.Should().Be(30f);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void BlurRadius_SetAndReadBack()
    {
        var r = CreateEmptyReflection();
        r.BlurRadius = 10;
        r.BlurRadius.Should().Be(10);
    }

    [Fact]
    public void BlurRadius_DefaultsToZero()
    {
        var r = CreateEmptyReflection();
        r.BlurRadius.Should().Be(0);
    }

    [Fact]
    public void Direction_SetAndReadBack()
    {
        var r = CreateEmptyReflection();
        r.Direction = 315;
        r.Direction.Should().Be(315);
    }

    [Fact]
    public void Direction_DefaultsToZero()
    {
        var r = CreateEmptyReflection();
        r.Direction.Should().Be(0);
    }

    [Fact]
    public void Distance_SetAndReadBack()
    {
        var r = CreateEmptyReflection();
        r.Distance = 8;
        r.Distance.Should().Be(8);
    }

    [Fact]
    public void Distance_DefaultsToZero()
    {
        var r = CreateEmptyReflection();
        r.Distance.Should().Be(0);
    }

    /// <summary>
    /// RectangleAlign maps OOXML short codes and defaults to Bottom.
    /// </summary>
    [Fact]
    public void RectangleAlign_DefaultsToBottom()
    {
        var r = CreateEmptyReflection();
        r.RectangleAlign.Should().Be(RectangleAlignment.Bottom);
    }

    [Theory]
    [InlineData("tl", RectangleAlignment.TopLeft)]
    [InlineData("t", RectangleAlignment.Top)]
    [InlineData("tr", RectangleAlignment.TopRight)]
    [InlineData("l", RectangleAlignment.Left)]
    [InlineData("ctr", RectangleAlignment.Center)]
    [InlineData("r", RectangleAlignment.Right)]
    [InlineData("bl", RectangleAlignment.BottomLeft)]
    [InlineData("b", RectangleAlignment.Bottom)]
    [InlineData("br", RectangleAlignment.BottomRight)]
    public void RectangleAlign_MapsOoxmlCodes(string ooxmlCode, RectangleAlignment expected)
    {
        var r = CreateReflection(("algn", ooxmlCode));
        r.RectangleAlign.Should().Be(expected);
    }

    [Fact]
    public void RectangleAlign_SetCenter_RoundTrips()
    {
        var r = CreateEmptyReflection();
        r.RectangleAlign = RectangleAlignment.Center;
        r.RectangleAlign.Should().Be(RectangleAlignment.Center);
    }

    [Fact]
    public void RectangleAlign_SetNotDefined_RemovesAttribute()
    {
        var r = CreateReflection(("algn", "ctr"));
        r.RectangleAlign = RectangleAlignment.NotDefined;
        r.RectangleAlign.Should().Be(RectangleAlignment.Bottom); // falls back to default
    }

    [Fact]
    public void RectangleAlign_UnknownOoxmlCode_ReturnsNotDefined()
    {
        var r = CreateReflection(("algn", "xyz"));
        r.RectangleAlign.Should().Be(RectangleAlignment.NotDefined);
    }

    /// <summary>
    /// SkewHorizontal reads kx, SkewVertical reads ky (angle in 60000ths of degree).
    /// </summary>
    [Fact]
    public void SkewHorizontal_ReadsFromKx()
    {
        var r = CreateReflection(("kx", "2700000"));
        r.SkewHorizontal.Should().Be(45d);
    }

    [Fact]
    public void SkewHorizontal_SetRoundTrips()
    {
        var r = CreateEmptyReflection();
        r.SkewHorizontal = 30d;
        r.SkewHorizontal.Should().Be(30d);
    }

    [Fact]
    public void SkewVertical_ReadsFromKy()
    {
        var r = CreateReflection(("ky", "1800000"));
        r.SkewVertical.Should().Be(30d);
    }

    [Fact]
    public void SkewVertical_SetRoundTrips()
    {
        var r = CreateEmptyReflection();
        r.SkewVertical = -15d;
        r.SkewVertical.Should().Be(-15d);
    }

    /// <summary>
    /// RotateShadowWithShape uses "1"/"0" in OOXML and defaults to true.
    /// </summary>
    [Fact]
    public void RotateShadowWithShape_DefaultsToTrue()
    {
        var r = CreateEmptyReflection();
        r.RotateShadowWithShape.Should().BeTrue();
    }

    [Fact]
    public void RotateShadowWithShape_ReadsOoxmlValues()
    {
        var r1 = CreateReflection(("rotWithShape", "1"));
        r1.RotateShadowWithShape.Should().BeTrue();

        var r0 = CreateReflection(("rotWithShape", "0"));
        r0.RotateShadowWithShape.Should().BeFalse();
    }

    [Fact]
    public void RotateShadowWithShape_SetRoundTrips()
    {
        var r = CreateEmptyReflection();
        r.RotateShadowWithShape = false;
        r.RotateShadowWithShape.Should().BeFalse();
        r.RotateShadowWithShape = true;
        r.RotateShadowWithShape.Should().BeTrue();
    }

    /// <summary>
    /// ScaleHorizontal reads sx, ScaleVertical reads sy (percentage * 1000).
    /// </summary>
    [Fact]
    public void ScaleHorizontal_ReadsFromSx()
    {
        var r = CreateReflection(("sx", "100000"));
        r.ScaleHorizontal.Should().Be(100d);
    }

    [Fact]
    public void ScaleHorizontal_SetRoundTrips()
    {
        var r = CreateEmptyReflection();
        r.ScaleHorizontal = 100d;
        r.ScaleHorizontal.Should().Be(100d);
    }

    [Fact]
    public void ScaleVertical_ReadsFromSy()
    {
        var r = CreateReflection(("sy", "50000"));
        r.ScaleVertical.Should().Be(50d);
    }

    [Fact]
    public void ScaleVertical_SetRoundTrips()
    {
        var r = CreateEmptyReflection();
        r.ScaleVertical = -50d;
        r.ScaleVertical.Should().Be(-50d);
    }

    /// <summary>
    /// AsIImageTransformOperation returns the same instance.
    /// </summary>
    [Fact]
    public void AsIImageTransformOperation_ReturnsSelf()
    {
        var r = CreateEmptyReflection();
        r.AsIImageTransformOperation.Should().BeSameAs(r);
    }

    /// <summary>
    /// Reflection implements IImageTransformOperation.
    /// </summary>
    [Fact]
    public void Reflection_IsImageTransformOperation()
    {
        IReflection reflection = CreateEmptyReflection();
        reflection.Should().BeAssignableTo<IImageTransformOperation>();
    }

    /// <summary>
    /// Multiple properties can be set together and persist correctly.
    /// </summary>
    [Fact]
    public void MultipleProperties_RoundTrip()
    {
        var r = CreateEmptyReflection();
        r.BlurRadius = 10;
        r.Direction = 315;
        r.Distance = 8;
        r.StartPosAlpha = 50;
        r.EndPosAlpha = 80;

        r.BlurRadius.Should().Be(10);
        r.Direction.Should().Be(315);
        r.Distance.Should().Be(8);
        r.StartPosAlpha.Should().Be(50);
        r.EndPosAlpha.Should().Be(80);
    }

    /// <summary>
    /// Parameterless constructor creates a Reflection with default values.
    /// </summary>
    [Fact]
    public void DefaultConstructor_HasNullElement()
    {
        var r = new Reflection();
        // With no backing element, getters return defaults
        r.StartPosAlpha.Should().Be(0f);
        r.EndPosAlpha.Should().Be(100f);
        r.FadeDirection.Should().Be(90f);
        r.StartReflectionOpacity.Should().Be(100f);
        r.EndReflectionOpacity.Should().Be(0f);
        r.BlurRadius.Should().Be(0d);
        r.Direction.Should().Be(0f);
        r.Distance.Should().Be(0d);
        r.RectangleAlign.Should().Be(RectangleAlignment.Bottom);
        r.SkewHorizontal.Should().Be(0d);
        r.SkewVertical.Should().Be(0d);
        r.RotateShadowWithShape.Should().BeTrue();
        r.ScaleHorizontal.Should().Be(0d);
        r.ScaleVertical.Should().Be(0d);
    }
}
