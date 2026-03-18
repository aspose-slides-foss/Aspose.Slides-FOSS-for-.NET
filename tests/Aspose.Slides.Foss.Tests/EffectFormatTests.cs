using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using Aspose.Slides.Foss.Effects;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Verifies effect format operations: shadow, glow, soft edge, blur, enable/disable.
/// </summary>
public sealed class EffectFormatTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static EffectFormat CreateEffectFormat()
    {
        var parent = new XElement(ANs + "spPr");
        var ef = new EffectFormat();
        ef.InitInternal(parent, parentSlide: null);
        return ef;
    }

    // ── test_outer_shadow ──

    [Fact]
    public void OuterShadow_EffectFormatCanBeInstantiated()
    {
        var ef = new EffectFormat();
        ef.Should().NotBeNull();
    }

    [Fact]
    public void OuterShadow_OuterShadowCanBeInstantiated()
    {
        var shadow = new OuterShadow();
        shadow.Should().NotBeNull();
    }

    [Fact]
    public void OuterShadow_SemiTransparentBlackColorHasCorrectComponents()
    {
        var c = Color.FromArgb(128, 0, 0, 0);
        c.A.Should().Be(128);
        c.R.Should().Be(0);
        c.G.Should().Be(0);
        c.B.Should().Be(0);
    }

    [Fact]
    public void OuterShadow_EnableCreatesEffect()
    {
        var ef = CreateEffectFormat();
        ef.EnableOuterShadowEffect();

        ef.OuterShadowEffect.Should().NotBeNull("outer shadow should exist after enabling");
    }

    [Fact]
    public void OuterShadow_PropertyPersistsInXml()
    {
        var ef = CreateEffectFormat();
        ef.EnableOuterShadowEffect();

        // Re-read from same XML backing store — simulates persistence.
        ef.OuterShadowEffect.Should().NotBeNull();
        ef.IsNoEffects.Should().BeFalse();
    }

    // ── test_glow ──

    [Fact]
    public void Glow_GlowCanBeInstantiated()
    {
        var glow = new Glow();
        glow.Should().NotBeNull();
    }

    [Fact]
    public void Glow_GoldColorExists()
    {
        Color.Gold.Should().NotBe(Color.Empty);
    }

    [Fact]
    public void Glow_GoldColorHasCorrectComponents()
    {
        Color.Gold.R.Should().Be(255);
        Color.Gold.G.Should().Be(215);
        Color.Gold.B.Should().Be(0);
    }

    [Fact]
    public void Glow_EnableCreatesEffect()
    {
        var ef = CreateEffectFormat();
        ef.EnableGlowEffect();

        ef.GlowEffect.Should().NotBeNull("glow effect should exist after enabling");
    }

    [Fact]
    public void Glow_PropertyPersistsInXml()
    {
        var ef = CreateEffectFormat();
        ef.EnableGlowEffect();

        ef.GlowEffect.Should().NotBeNull();
        ef.IsNoEffects.Should().BeFalse();
    }

    // ── test_soft_edge ──

    [Fact]
    public void SoftEdge_SoftEdgeCanBeInstantiated()
    {
        var softEdge = new SoftEdge();
        softEdge.Should().NotBeNull();
    }

    [Fact]
    public void SoftEdge_EnableCreatesEffect()
    {
        var ef = CreateEffectFormat();
        ef.EnableSoftEdgeEffect();

        ef.SoftEdgeEffect.Should().NotBeNull("soft edge should exist after enabling");
    }

    [Fact]
    public void SoftEdge_PropertyPersistsInXml()
    {
        var ef = CreateEffectFormat();
        ef.EnableSoftEdgeEffect();

        ef.SoftEdgeEffect.Should().NotBeNull();
        ef.IsNoEffects.Should().BeFalse();
    }

    // ── test_blur ──

    [Fact]
    public void Blur_BlurCanBeInstantiated()
    {
        var blur = new Blur();
        blur.Should().NotBeNull();
    }

    [Fact]
    public void Blur_SetBlurEffectCreatesEffect()
    {
        var ef = CreateEffectFormat();
        ef.SetBlurEffect(8, true);

        ef.BlurEffect.Should().NotBeNull("blur effect should exist after SetBlurEffect");
    }

    [Fact]
    public void Blur_SetBlurEffectPersistsRadiusInXml()
    {
        var parent = new XElement(ANs + "spPr");
        var ef = new EffectFormat();
        ef.InitInternal(parent, parentSlide: null);
        ef.SetBlurEffect(8, true);

        // Verify the XML has the correct radius (8 points * 12700 EMU/point = 101600).
        var blurEl = parent.Element(ANs + "effectLst")?.Element(ANs + "blur");
        blurEl.Should().NotBeNull();
        blurEl!.Attribute("rad")!.Value.Should().Be("101600");
        blurEl.Attribute("grow")!.Value.Should().Be("1");
    }

    [Fact]
    public void Blur_SetBlurEffectGrowFalseWritesZero()
    {
        var parent = new XElement(ANs + "spPr");
        var ef = new EffectFormat();
        ef.InitInternal(parent, parentSlide: null);
        ef.SetBlurEffect(8, false);

        var blurEl = parent.Element(ANs + "effectLst")?.Element(ANs + "blur");
        blurEl!.Attribute("grow")!.Value.Should().Be("0");
    }

    // ── test_enable_disable_effects ──

    [Fact]
    public void EnableDisableEffects_AllEffectTypesExist()
    {
        new EffectFormat().Should().NotBeNull();
        new OuterShadow().Should().NotBeNull();
        new Glow().Should().NotBeNull();
        new SoftEdge().Should().NotBeNull();
        new Blur().Should().NotBeNull();
    }

    [Fact]
    public void EnableDisableEffects_EffectFormatImplementsInterface()
    {
        var ef = new EffectFormat();
        ef.Should().BeAssignableTo<IEffectFormat>();
    }

    [Fact]
    public void EnableDisableEffects_OuterShadowImplementsInterface()
    {
        var shadow = new OuterShadow();
        shadow.Should().BeAssignableTo<IOuterShadow>();
    }

    [Fact]
    public void EnableDisableEffects_GlowImplementsInterface()
    {
        var glow = new Glow();
        glow.Should().BeAssignableTo<IGlow>();
    }

    [Fact]
    public void EnableDisableEffects_SoftEdgeImplementsInterface()
    {
        var softEdge = new SoftEdge();
        softEdge.Should().BeAssignableTo<ISoftEdge>();
    }

    [Fact]
    public void EnableDisableEffects_BlurImplementsInterface()
    {
        var blur = new Blur();
        blur.Should().BeAssignableTo<IBlur>();
    }

    [Fact]
    public void EnableDisableEffects_EnabledEffectsAreNotEmpty()
    {
        var ef = CreateEffectFormat();
        ef.EnableOuterShadowEffect();
        ef.EnableGlowEffect();

        ef.IsNoEffects.Should().BeFalse();
    }

    [Fact]
    public void EnableDisableEffects_DisablingAllEffectsRestoresEmpty()
    {
        var ef = CreateEffectFormat();
        ef.EnableOuterShadowEffect();
        ef.EnableGlowEffect();
        ef.IsNoEffects.Should().BeFalse();

        ef.DisableOuterShadowEffect();
        ef.DisableGlowEffect();
        ef.IsNoEffects.Should().BeTrue();
    }

    [Fact]
    public void EnableDisableEffects_StartsWithNoEffects()
    {
        var ef = CreateEffectFormat();
        ef.IsNoEffects.Should().BeTrue();
    }
}
