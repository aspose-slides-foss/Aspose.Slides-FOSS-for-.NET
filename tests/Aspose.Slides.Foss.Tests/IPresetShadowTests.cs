using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Effects;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests for the IPresetShadow interface contract.
/// </summary>
public sealed class IPresetShadowTests
{
    /// <summary>
    /// Stub implementation of IPresetShadow for testing the interface contract.
    /// </summary>
    private sealed class PresetShadowStub : IPresetShadow
    {
        public float Direction { get; set; }
        public double Distance { get; set; }
        public IColorFormat ShadowColor { get; } = null!;
        public PresetShadowType Preset { get; set; }
        public IImageTransformOperation AsIImageTransformOperation => this;
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void Direction_CanBeSetAndReadBack()
    {
        IPresetShadow shadow = new PresetShadowStub();
        shadow.Direction = 315;
        shadow.Direction.Should().Be(315);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void Distance_CanBeSetAndReadBack()
    {
        IPresetShadow shadow = new PresetShadowStub();
        shadow.Distance = 8;
        shadow.Distance.Should().Be(8);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void MultipleProperties_CanBeSetTogether()
    {
        IPresetShadow shadow = new PresetShadowStub();
        shadow.Direction = 315;
        shadow.Distance = 8;
        shadow.Preset = PresetShadowType.TopLeftDropShadow;

        shadow.Direction.Should().Be(315);
        shadow.Distance.Should().Be(8);
        shadow.Preset.Should().Be(PresetShadowType.TopLeftDropShadow);
    }

    /// <summary>
    /// IPresetShadow extends IImageTransformOperation.
    /// </summary>
    [Fact]
    public void PresetShadow_IsImageTransformOperation()
    {
        IPresetShadow shadow = new PresetShadowStub();
        shadow.Should().BeAssignableTo<IImageTransformOperation>();
    }

    /// <summary>
    /// AsIImageTransformOperation returns the same instance.
    /// </summary>
    [Fact]
    public void AsIImageTransformOperation_ReturnsSelf()
    {
        IPresetShadow shadow = new PresetShadowStub();
        shadow.AsIImageTransformOperation.Should().BeSameAs(shadow);
    }

    /// <summary>
    /// Default numeric values are zero.
    /// </summary>
    [Fact]
    public void DefaultValues_AreZero()
    {
        IPresetShadow shadow = new PresetShadowStub();
        shadow.Direction.Should().Be(0);
        shadow.Distance.Should().Be(0);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void Preset_CanBeSetAndReadBack()
    {
        IPresetShadow shadow = new PresetShadowStub();
        shadow.Preset = PresetShadowType.BottomRightDropShadow;
        shadow.Preset.Should().Be(PresetShadowType.BottomRightDropShadow);
    }

    /// <summary>
    /// Default preset value is the first enum member.
    /// </summary>
    [Fact]
    public void DefaultPreset_IsFirstEnumValue()
    {
        IPresetShadow shadow = new PresetShadowStub();
        shadow.Preset.Should().Be(PresetShadowType.TopLeftDropShadow);
    }
}
