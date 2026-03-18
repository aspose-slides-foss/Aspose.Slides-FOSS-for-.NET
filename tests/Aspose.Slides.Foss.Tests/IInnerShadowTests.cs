using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Effects;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests for the IInnerShadow interface contract.
/// these unit tests verify the IInnerShadow interface shape and property behavior.
/// </summary>
public sealed class IInnerShadowTests
{
    /// <summary>
    /// Stub implementation of IInnerShadow for testing the interface contract.
    /// </summary>
    private sealed class InnerShadowStub : IInnerShadow
    {
        public double BlurRadius { get; set; }
        public float Direction { get; set; }
        public double Distance { get; set; }
        public IColorFormat ShadowColor { get; } = null!;
        public IImageTransformOperation AsIImageTransformOperation => this;
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void BlurRadius_CanBeSetAndReadBack()
    {
        IInnerShadow shadow = new InnerShadowStub();
        shadow.BlurRadius = 10;
        shadow.BlurRadius.Should().Be(10);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void Direction_CanBeSetAndReadBack()
    {
        IInnerShadow shadow = new InnerShadowStub();
        shadow.Direction = 315;
        shadow.Direction.Should().Be(315);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void Distance_CanBeSetAndReadBack()
    {
        IInnerShadow shadow = new InnerShadowStub();
        shadow.Distance = 8;
        shadow.Distance.Should().Be(8);
    }

    /// <summary>
    /// IInnerShadow extends IImageTransformOperation.
    /// </summary>
    [Fact]
    public void InnerShadow_IsImageTransformOperation()
    {
        IInnerShadow shadow = new InnerShadowStub();
        shadow.Should().BeAssignableTo<IImageTransformOperation>();
    }

    /// <summary>
    /// AsIImageTransformOperation returns the same instance.
    /// </summary>
    [Fact]
    public void AsIImageTransformOperation_ReturnsSelf()
    {
        IInnerShadow shadow = new InnerShadowStub();
        shadow.AsIImageTransformOperation.Should().BeSameAs(shadow);
    }

    /// <summary>
    /// Default values are zero for numeric properties.
    /// </summary>
    [Fact]
    public void DefaultValues_AreZero()
    {
        IInnerShadow shadow = new InnerShadowStub();
        shadow.BlurRadius.Should().Be(0);
        shadow.Direction.Should().Be(0);
        shadow.Distance.Should().Be(0);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void MultipleProperties_CanBeSetTogether()
    {
        IInnerShadow shadow = new InnerShadowStub();
        shadow.BlurRadius = 10;
        shadow.Direction = 315;
        shadow.Distance = 8;

        shadow.BlurRadius.Should().Be(10);
        shadow.Direction.Should().Be(315);
        shadow.Distance.Should().Be(8);
    }
}
