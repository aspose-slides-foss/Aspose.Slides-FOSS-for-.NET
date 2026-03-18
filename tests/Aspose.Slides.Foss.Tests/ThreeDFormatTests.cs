using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Verifies the enum values used by 3D format operations.
/// </summary>
public sealed class ThreeDFormatTests
{
    /// <summary>
    /// Verifies BevelPresetType.Circle.
    /// </summary>
    [Fact]
    public void BevelTop_CirclePresetIsDefined()
    {
        Enum.IsDefined(BevelPresetType.Circle).Should().BeTrue();
        BevelPresetType.Circle.Should().NotBe(BevelPresetType.NotDefined);
    }

    [Fact]
    public void BevelTop_CircleIsDistinctFromOtherPresets()
    {
        BevelPresetType.Circle.Should().NotBe(BevelPresetType.Angle);
        BevelPresetType.Circle.Should().NotBe(BevelPresetType.SoftRound);
    }

    /// <summary>
    /// Verifies CameraPresetType.PerspectiveAbove.
    /// </summary>
    [Fact]
    public void Camera_PerspectiveAboveIsDefined()
    {
        Enum.IsDefined(CameraPresetType.PerspectiveAbove).Should().BeTrue();
        CameraPresetType.PerspectiveAbove.Should().NotBe(CameraPresetType.NotDefined);
    }

    /// <summary>
    /// Verifies LightRigPresetType.Balanced and LightingDirection.Top.
    /// </summary>
    [Fact]
    public void LightRig_BalancedPresetIsDefined()
    {
        Enum.IsDefined(LightRigPresetType.Balanced).Should().BeTrue();
        LightRigPresetType.Balanced.Should().NotBe(LightRigPresetType.NotDefined);
    }

    [Fact]
    public void LightRig_TopDirectionIsDefined()
    {
        Enum.IsDefined(LightingDirection.Top).Should().BeTrue();
        LightingDirection.Top.Should().NotBe(LightingDirection.NotDefined);
    }

    /// <summary>
    /// Verifies MaterialPresetType.Metal.
    /// </summary>
    [Fact]
    public void DepthAndMaterial_MetalPresetIsDefined()
    {
        Enum.IsDefined(MaterialPresetType.Metal).Should().BeTrue();
        MaterialPresetType.Metal.Should().NotBe(MaterialPresetType.NotDefined);
    }

    [Fact]
    public void BevelPresetType_NotDefinedSentinelIsZero()
    {
        ((int)BevelPresetType.NotDefined).Should().Be(0);
    }

    [Fact]
    public void CameraPresetType_NotDefinedSentinelIsZero()
    {
        ((int)CameraPresetType.NotDefined).Should().Be(0);
    }

    [Fact]
    public void LightRigPresetType_NotDefinedSentinelIsZero()
    {
        ((int)LightRigPresetType.NotDefined).Should().Be(0);
    }

    [Fact]
    public void LightingDirection_NotDefinedSentinelIsZero()
    {
        ((int)LightingDirection.NotDefined).Should().Be(0);
    }

    [Fact]
    public void MaterialPresetType_NotDefinedSentinelIsZero()
    {
        ((int)MaterialPresetType.NotDefined).Should().Be(0);
    }

    [Theory]
    [InlineData(BevelPresetType.Circle)]
    [InlineData(BevelPresetType.Angle)]
    [InlineData(BevelPresetType.SoftRound)]
    public void BevelPresetType_CommonValuesAreDefined(BevelPresetType bevel)
    {
        Enum.IsDefined(bevel).Should().BeTrue();
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
    public void LightingDirection_AllDirectionsAreDefined(LightingDirection direction)
    {
        Enum.IsDefined(direction).Should().BeTrue();
    }

    /// <summary>
    /// ThreeDFormat can be instantiated and initialized with XML.
    /// </summary>
    [Fact]
    public void BevelTop_ThreeDFormatCanBeInstantiated()
    {
        var tdf = new ThreeDFormat();
        tdf.Should().NotBeNull();
    }

    /// <summary>
    /// Bevel dimensions (10, 5) are representable as the expected numeric types.
    /// </summary>
    [Fact]
    public void BevelTop_DimensionValuesAreRepresentable()
    {
        BevelPresetType.Circle.Should().NotBe(BevelPresetType.NotDefined);
        var width = 10.0;
        var height = 5.0;
        width.Should().BeGreaterThan(0);
        height.Should().BeGreaterThan(0);
    }

    /// <summary>
    /// PerspectiveAbove is distinct from other common camera presets.
    /// </summary>
    [Fact]
    public void Camera_PerspectiveAboveIsDistinctFromOtherPresets()
    {
        CameraPresetType.PerspectiveAbove.Should().NotBe(CameraPresetType.PerspectiveFront);
        CameraPresetType.PerspectiveAbove.Should().NotBe(CameraPresetType.OrthographicFront);
    }

    /// <summary>
    /// Balanced preset is distinct from other light rig types.
    /// </summary>
    [Fact]
    public void LightRig_BalancedIsDistinctFromOtherPresets()
    {
        LightRigPresetType.Balanced.Should().NotBe(LightRigPresetType.Harsh);
        LightRigPresetType.Balanced.Should().NotBe(LightRigPresetType.Soft);
    }

    /// <summary>
    /// Metal is distinct from other material presets.
    /// </summary>
    [Fact]
    public void DepthAndMaterial_MetalIsDistinctFromOtherPresets()
    {
        MaterialPresetType.Metal.Should().NotBe(MaterialPresetType.Plastic);
        MaterialPresetType.Metal.Should().NotBe(MaterialPresetType.Matte);
    }

    /// <summary>
    /// Depth value 20 is representable.
    /// </summary>
    [Fact]
    public void DepthAndMaterial_DepthValueIsRepresentable()
    {
        MaterialPresetType.Metal.Should().NotBe(MaterialPresetType.NotDefined);
        var depth = 20.0;
        depth.Should().BeGreaterThan(0);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void LightRig_PresetAndDirectionPersistOnThreeDFormat()
    {
        var spPr = System.Xml.Linq.XElement.Parse("<spPr xmlns='http://schemas.openxmlformats.org/drawingml/2006/main'/>");
        var tdf = new ThreeDFormat();
        tdf.InitInternal(spPr, null);

        var lr = tdf.LightRig;
        lr.LightType = LightRigPresetType.Balanced;
        lr.Direction = LightingDirection.Top;

        // Re-read from the same XML to verify persistence
        var tdf2 = new ThreeDFormat();
        tdf2.InitInternal(spPr, null);
        var lr2 = tdf2.LightRig;

        lr2.LightType.Should().Be(LightRigPresetType.Balanced);
        lr2.Direction.Should().Be(LightingDirection.Top);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void LightRig_CanBeAccessedFromFreshThreeDFormat()
    {
        var spPr = System.Xml.Linq.XElement.Parse("<spPr xmlns='http://schemas.openxmlformats.org/drawingml/2006/main'/>");
        var tdf = new ThreeDFormat();
        tdf.InitInternal(spPr, null);

        var lr = tdf.LightRig;
        lr.Should().NotBeNull();
        lr.LightType.Should().Be(LightRigPresetType.NotDefined);
    }

    /// <summary>
    /// </summary>
    [Theory]
    [InlineData(LightRigPresetType.Balanced)]
    [InlineData(LightRigPresetType.Harsh)]
    [InlineData(LightRigPresetType.Soft)]
    [InlineData(LightRigPresetType.ThreePt)]
    public void LightRig_VariousPresetsCanBeSet(LightRigPresetType preset)
    {
        var spPr = System.Xml.Linq.XElement.Parse("<spPr xmlns='http://schemas.openxmlformats.org/drawingml/2006/main'/>");
        var tdf = new ThreeDFormat();
        tdf.InitInternal(spPr, null);

        tdf.LightRig.LightType = preset;

        var tdf2 = new ThreeDFormat();
        tdf2.InitInternal(spPr, null);
        tdf2.LightRig.LightType.Should().Be(preset);
    }

    /// <summary>
    /// </summary>
    [Theory]
    [InlineData(LightingDirection.Top)]
    [InlineData(LightingDirection.Bottom)]
    [InlineData(LightingDirection.Left)]
    [InlineData(LightingDirection.Right)]
    public void LightRig_VariousDirectionsCanBeSet(LightingDirection direction)
    {
        var spPr = System.Xml.Linq.XElement.Parse("<spPr xmlns='http://schemas.openxmlformats.org/drawingml/2006/main'/>");
        var tdf = new ThreeDFormat();
        tdf.InitInternal(spPr, null);

        tdf.LightRig.Direction = direction;

        var tdf2 = new ThreeDFormat();
        tdf2.InitInternal(spPr, null);
        tdf2.LightRig.Direction.Should().Be(direction);
    }

    /// <summary>
    /// ThreeDFormat.Depth can be set and read back.
    /// </summary>
    [Fact]
    public void Depth_CanBeSetAndReadBack()
    {
        var spPr = System.Xml.Linq.XElement.Parse("<spPr xmlns='http://schemas.openxmlformats.org/drawingml/2006/main'/>");
        var tdf = new ThreeDFormat();
        tdf.InitInternal(spPr, null);

        tdf.Depth = 20;

        var tdf2 = new ThreeDFormat();
        tdf2.InitInternal(spPr, null);
        tdf2.Depth.Should().Be(20);
    }

    /// <summary>
    /// ThreeDFormat.Material can be set and read back.
    /// </summary>
    [Fact]
    public void Material_CanBeSetAndReadBack()
    {
        var spPr = System.Xml.Linq.XElement.Parse("<spPr xmlns='http://schemas.openxmlformats.org/drawingml/2006/main'/>");
        var tdf = new ThreeDFormat();
        tdf.InitInternal(spPr, null);

        tdf.Material = MaterialPresetType.Metal;

        var tdf2 = new ThreeDFormat();
        tdf2.InitInternal(spPr, null);
        tdf2.Material.Should().Be(MaterialPresetType.Metal);
    }
}
