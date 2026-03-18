using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using FluentAssertions;

namespace Aspose.Slides.Foss.IntegrationTests;

/// <summary>
/// Tests for EffectFormat: shadow, glow, soft edge, blur, reflection.
/// </summary>
public sealed class EffectFormatTests : IDisposable
{
    private readonly string _tempDir;

    public EffectFormatTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, recursive: true);
    }

    [Fact]
    public void TestOuterShadow()
    {
        // Outer shadow properties persist after save/reload.
        using var pres = new Presentation();
        var slide = pres.Slides[0];
        slide.Shapes!.Clear();
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var ef = shape.EffectFormat;
        ef.EnableOuterShadowEffect();
        var shadow = ef.OuterShadowEffect!;
        shadow.BlurRadius = 10;
        shadow.Direction = 315;
        shadow.Distance = 8;
        shadow.ShadowColor.Color = Color.FromArgb(128, 0, 0, 0);

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var ef2 = pres2.Slides[0].Shapes![0].EffectFormat;
        var s2 = ef2.OuterShadowEffect;
        s2.Should().NotBeNull("outer_shadow_effect should not be None after reload");
        s2!.BlurRadius.Should().Be(10);
        s2.Direction.Should().Be(315);
        s2.Distance.Should().Be(8);
    }

    [Fact]
    public void TestGlow()
    {
        // Glow effect persists.
        using var pres = new Presentation();
        var slide = pres.Slides[0];
        slide.Shapes!.Clear();
        var shape = slide.Shapes.AddAutoShape(ShapeType.Ellipse, 100, 100, 200, 200);
        var ef = shape.EffectFormat;
        ef.EnableGlowEffect();
        ef.GlowEffect!.Radius = 15;
        ef.GlowEffect.Color.Color = Color.Gold;

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var g2 = pres2.Slides[0].Shapes![0].EffectFormat.GlowEffect;
        g2.Should().NotBeNull("glow_effect should not be None after reload");
        g2!.Radius.Should().Be(15);
    }

    [Fact]
    public void TestSoftEdge()
    {
        // Soft edge radius persists.
        using var pres = new Presentation();
        var slide = pres.Slides[0];
        slide.Shapes!.Clear();
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var ef = shape.EffectFormat;
        ef.EnableSoftEdgeEffect();
        ef.SoftEdgeEffect!.Radius = 10;

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var se2 = pres2.Slides[0].Shapes![0].EffectFormat.SoftEdgeEffect;
        se2.Should().NotBeNull("soft_edge_effect should not be None after reload");
        se2!.Radius.Should().Be(10);
    }

    [Fact]
    public void TestBlur()
    {
        // Blur effect persists.
        using var pres = new Presentation();
        var slide = pres.Slides[0];
        slide.Shapes!.Clear();
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var ef = shape.EffectFormat;
        ef.SetBlurEffect(8, true);

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var b2 = pres2.Slides[0].Shapes![0].EffectFormat.BlurEffect;
        b2.Should().NotBeNull("blur_effect should not be None after reload");
        b2!.Radius.Should().Be(8);
    }

    [Fact]
    public void TestEnableDisableEffects()
    {
        // Effects can be enabled then disabled.
        using var pres = new Presentation();
        var shape = pres.Slides[0].Shapes!.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var ef = shape.EffectFormat;
        ef.EnableOuterShadowEffect();
        ef.EnableGlowEffect();
        ef.IsNoEffects.Should().BeFalse();

        ef.DisableOuterShadowEffect();
        ef.DisableGlowEffect();
        ef.IsNoEffects.Should().BeTrue();
    }
}
