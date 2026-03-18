using Aspose.Slides.Foss.Effects;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents visual effect formatting properties for a shape.
/// </summary>
public interface IEffectFormat : IEffectParamSource
{
    /// <summary>
    /// Gets a value indicating whether no effects are applied.
    /// </summary>
    bool IsNoEffects { get; }

    /// <summary>Gets or sets the blur effect.</summary>
    IBlur? BlurEffect { get; set; }

    /// <summary>Gets or sets the fill overlay effect.</summary>
    IFillOverlay? FillOverlayEffect { get; set; }

    /// <summary>Gets or sets the glow effect.</summary>
    IGlow? GlowEffect { get; set; }

    /// <summary>Gets or sets the inner shadow effect.</summary>
    IInnerShadow? InnerShadowEffect { get; set; }

    /// <summary>Gets or sets the outer shadow effect.</summary>
    IOuterShadow? OuterShadowEffect { get; set; }

    /// <summary>Gets or sets the preset shadow effect.</summary>
    IPresetShadow? PresetShadowEffect { get; set; }

    /// <summary>Gets or sets the reflection effect.</summary>
    IReflection? ReflectionEffect { get; set; }

    /// <summary>Gets or sets the soft edge effect.</summary>
    ISoftEdge? SoftEdgeEffect { get; set; }

    /// <summary>
    /// Returns this instance as <see cref="IEffectParamSource"/>.
    /// </summary>
    IEffectParamSource AsIEffectParamSource { get; }

    /// <summary>Enables the blur effect.</summary>
    void EnableBlurEffect();

    /// <summary>Disables the blur effect.</summary>
    void DisableBlurEffect();

    /// <summary>Enables the fill overlay effect.</summary>
    void EnableFillOverlayEffect();

    /// <summary>Disables the fill overlay effect.</summary>
    void DisableFillOverlayEffect();

    /// <summary>Enables the glow effect.</summary>
    void EnableGlowEffect();

    /// <summary>Disables the glow effect.</summary>
    void DisableGlowEffect();

    /// <summary>Enables the inner shadow effect.</summary>
    void EnableInnerShadowEffect();

    /// <summary>Disables the inner shadow effect.</summary>
    void DisableInnerShadowEffect();

    /// <summary>Enables the outer shadow effect.</summary>
    void EnableOuterShadowEffect();

    /// <summary>Disables the outer shadow effect.</summary>
    void DisableOuterShadowEffect();

    /// <summary>Enables the preset shadow effect.</summary>
    void EnablePresetShadowEffect();

    /// <summary>Disables the preset shadow effect.</summary>
    void DisablePresetShadowEffect();

    /// <summary>Enables the reflection effect.</summary>
    void EnableReflectionEffect();

    /// <summary>Disables the reflection effect.</summary>
    void DisableReflectionEffect();

    /// <summary>Enables the soft edge effect.</summary>
    void EnableSoftEdgeEffect();

    /// <summary>Disables the soft edge effect.</summary>
    void DisableSoftEdgeEffect();

    /// <summary>Sets blur effect parameters.</summary>
    void SetBlurEffect(float radius, bool grow);
}
