using System.Xml.Linq;
using Aspose.Slides.Foss.Effects;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Unit tests for <see cref="Blur"/> covering property behavior,
/// XML attribute mapping, and the <see cref="IBlur"/> contract.
/// </summary>
public sealed class BlurTests
{
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static Blur CreateFromXml(string? rad = null, string? grow = null)
    {
        var el = new XElement(A + "blur");
        if (rad is not null) el.SetAttributeValue("rad", rad);
        if (grow is not null) el.SetAttributeValue("grow", grow);
        return new Blur(el);
    }

    /// <summary>
    /// Radius converts EMU attribute to points (1 point = 12700 EMU).
    /// </summary>
    [Theory]
    [InlineData("101600", 8.0f)]   // 8 * 12700
    [InlineData("63500", 5.0f)]    // 5 * 12700
    [InlineData("0", 0f)]
    public void Radius_ReadsEmuAndConvertsToPoints(string radEmu, float expectedPoints)
    {
        var blur = CreateFromXml(rad: radEmu);
        blur.Radius.Should().Be(expectedPoints);
    }

    /// <summary>
    /// Setting Radius writes the correct EMU value to the XML attribute.
    /// </summary>
    [Fact]
    public void Radius_Set_WritesEmuAttribute()
    {
        var blur = CreateFromXml();
        blur.Radius = 8.0f;

        blur.Element!.Attribute("rad")!.Value.Should().Be("101600");
    }

    /// <summary>
    /// Radius defaults to 0 when the rad attribute is absent.
    /// </summary>
    [Fact]
    public void Radius_DefaultsToZero_WhenAttributeAbsent()
    {
        var blur = CreateFromXml();
        blur.Radius.Should().Be(0f);
    }

    /// <summary>
    /// Grow returns true when the attribute is "1".
    /// </summary>
    [Fact]
    public void Grow_ReturnsTrue_WhenAttributeIsOne()
    {
        var blur = CreateFromXml(grow: "1");
        blur.Grow.Should().BeTrue();
    }

    /// <summary>
    /// Grow returns false when the attribute is "0".
    /// </summary>
    [Fact]
    public void Grow_ReturnsFalse_WhenAttributeIsZero()
    {
        var blur = CreateFromXml(grow: "0");
        blur.Grow.Should().BeFalse();
    }

    /// <summary>
    /// Grow defaults to true when the attribute is absent.
    /// </summary>
    [Fact]
    public void Grow_DefaultsToTrue_WhenAttributeAbsent()
    {
        var blur = CreateFromXml();
        blur.Grow.Should().BeTrue();
    }

    /// <summary>
    /// Setting Grow writes the correct attribute value.
    /// </summary>
    [Theory]
    [InlineData(true, "1")]
    [InlineData(false, "0")]
    public void Grow_Set_WritesCorrectAttribute(bool value, string expected)
    {
        var blur = CreateFromXml();
        blur.Grow = value;

        blur.Element!.Attribute("grow")!.Value.Should().Be(expected);
    }

    /// <summary>
    /// AsIImageTransformOperation returns the instance itself.
    /// </summary>
    [Fact]
    public void AsIImageTransformOperation_ReturnsSelf()
    {
        var blur = CreateFromXml();
        blur.AsIImageTransformOperation.Should().BeSameAs(blur);
    }

    /// <summary>
    /// Blur implements IBlur and IImageTransformOperation.
    /// </summary>
    [Fact]
    public void Blur_ImplementsExpectedInterfaces()
    {
        var blur = CreateFromXml();
        blur.Should().BeAssignableTo<IBlur>();
        blur.Should().BeAssignableTo<IImageTransformOperation>();
    }

    /// <summary>
    /// Default constructor creates a Blur with null element; property access does not throw.
    /// </summary>
    [Fact]
    public void DefaultConstructor_PropertiesDoNotThrow()
    {
        var blur = new Blur();
        blur.Radius.Should().Be(0f);
        blur.Grow.Should().BeTrue();
        blur.Slide.Should().BeNull();
        blur.AsIPresentationComponent.Should().BeSameAs(blur);
    }
}
