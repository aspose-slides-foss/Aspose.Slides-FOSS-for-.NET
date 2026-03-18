using System.Xml.Linq;
using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Unit tests for the LightRig class.
/// </summary>
public sealed class LightRigTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static LightRig CreateLightRig(XElement? scene3d = null)
    {
        scene3d ??= new XElement(ANs + "scene3d");
        var spPr = new XElement(ANs + "spPr", scene3d);
        var tdf = new ThreeDFormat();
        tdf.InitInternal(spPr, null);
        return (LightRig)tdf.LightRig;
    }

    [Fact]
    public void LightType_DefaultIsNotDefined()
    {
        var lr = CreateLightRig();
        lr.LightType.Should().Be(LightRigPresetType.NotDefined);
    }

    [Fact]
    public void Direction_DefaultIsNotDefined()
    {
        var lr = CreateLightRig();
        lr.Direction.Should().Be(LightingDirection.NotDefined);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void LightType_SetBalanced_ReadBackBalanced()
    {
        var lr = CreateLightRig();
        lr.LightType = LightRigPresetType.Balanced;
        lr.LightType.Should().Be(LightRigPresetType.Balanced);
    }

    [Fact]
    public void Direction_SetTop_ReadBackTop()
    {
        var lr = CreateLightRig();
        lr.Direction = LightingDirection.Top;
        lr.Direction.Should().Be(LightingDirection.Top);
    }

    [Theory]
    [InlineData(LightRigPresetType.Balanced)]
    [InlineData(LightRigPresetType.BrightRoom)]
    [InlineData(LightRigPresetType.Harsh)]
    [InlineData(LightRigPresetType.Soft)]
    [InlineData(LightRigPresetType.ThreePt)]
    [InlineData(LightRigPresetType.TwoPt)]
    [InlineData(LightRigPresetType.Flood)]
    [InlineData(LightRigPresetType.Morning)]
    public void LightType_SetAndReadBack(LightRigPresetType preset)
    {
        var lr = CreateLightRig();
        lr.LightType = preset;
        lr.LightType.Should().Be(preset);
    }

    [Theory]
    [InlineData(LightingDirection.TopLeft)]
    [InlineData(LightingDirection.Top)]
    [InlineData(LightingDirection.TopRight)]
    [InlineData(LightingDirection.Right)]
    [InlineData(LightingDirection.BottomRight)]
    [InlineData(LightingDirection.Bottom)]
    [InlineData(LightingDirection.BottomLeft)]
    [InlineData(LightingDirection.Left)]
    public void Direction_SetAndReadBack(LightingDirection direction)
    {
        var lr = CreateLightRig();
        lr.Direction = direction;
        lr.Direction.Should().Be(direction);
    }

    [Fact]
    public void LightType_SetNotDefined_RemovesRigAttribute()
    {
        var lr = CreateLightRig();
        lr.LightType = LightRigPresetType.Balanced;
        lr.LightType.Should().Be(LightRigPresetType.Balanced);

        lr.LightType = LightRigPresetType.NotDefined;
        lr.LightType.Should().Be(LightRigPresetType.NotDefined);
    }

    [Fact]
    public void Direction_SetNotDefined_RemovesDirAttribute()
    {
        var lr = CreateLightRig();
        lr.Direction = LightingDirection.Top;
        lr.Direction.Should().Be(LightingDirection.Top);

        lr.Direction = LightingDirection.NotDefined;
        lr.Direction.Should().Be(LightingDirection.NotDefined);
    }

    [Fact]
    public void SetRotation_GetRotation_RoundTrip()
    {
        var lr = CreateLightRig();
        lr.SetRotation(45f, 90f, 180f);
        var rot = lr.GetRotation();
        rot.Should().HaveCount(3);
        rot[0].Should().BeApproximately(45f, 0.01f);
        rot[1].Should().BeApproximately(90f, 0.01f);
        rot[2].Should().BeApproximately(180f, 0.01f);
    }

    [Fact]
    public void GetRotation_DefaultIsAllZeros()
    {
        var lr = CreateLightRig();
        var rot = lr.GetRotation();
        rot.Should().Equal([0f, 0f, 0f]);
    }

    [Fact]
    public void SetRotation_OverwritesPreviousValues()
    {
        var lr = CreateLightRig();
        lr.SetRotation(10f, 20f, 30f);
        lr.SetRotation(45f, 60f, 90f);
        var rot = lr.GetRotation();
        rot[0].Should().BeApproximately(45f, 0.01f);
        rot[1].Should().BeApproximately(60f, 0.01f);
        rot[2].Should().BeApproximately(90f, 0.01f);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void PresetAndDirection_PersistViaThreeDFormat()
    {
        var spPr = XElement.Parse("<spPr xmlns='http://schemas.openxmlformats.org/drawingml/2006/main'/>");
        var tdf = new ThreeDFormat();
        tdf.InitInternal(spPr, null);

        tdf.LightRig.LightType = LightRigPresetType.Balanced;
        tdf.LightRig.Direction = LightingDirection.Top;

        // Re-read from same XML to verify persistence
        var tdf2 = new ThreeDFormat();
        tdf2.InitInternal(spPr, null);
        tdf2.LightRig.LightType.Should().Be(LightRigPresetType.Balanced);
        tdf2.LightRig.Direction.Should().Be(LightingDirection.Top);
    }

    [Fact]
    public void SetRotation_ZeroValues()
    {
        var lr = CreateLightRig();
        lr.SetRotation(0f, 0f, 0f);
        var rot = lr.GetRotation();
        rot.Should().Equal([0f, 0f, 0f]);
    }
}
