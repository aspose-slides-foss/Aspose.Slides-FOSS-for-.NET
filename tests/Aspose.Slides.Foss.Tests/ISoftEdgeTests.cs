using System.Xml.Linq;
using Aspose.Slides.Foss.Effects;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Unit tests for <see cref="SoftEdge"/> covering property behavior,
/// XML attribute mapping, and the <see cref="ISoftEdge"/> contract.
/// </summary>
public sealed class ISoftEdgeTests
{
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static SoftEdge CreateFromXml(string? rad = null)
    {
        var el = new XElement(A + "softEdge");
        if (rad is not null) el.SetAttributeValue("rad", rad);
        return new SoftEdge(el);
    }

    /// <summary>
    /// Radius converts EMU attribute to points (1 point = 12700 EMU).
    /// </summary>
    [Theory]
    [InlineData("127000", 10.0f)]  // 10 * 12700
    [InlineData("63500", 5.0f)]    // 5 * 12700
    [InlineData("0", 0f)]
    public void Radius_ReadsEmuAndConvertsToPoints(string radEmu, float expectedPoints)
    {
        var softEdge = CreateFromXml(rad: radEmu);
        softEdge.Radius.Should().Be(expectedPoints);
    }

    /// <summary>
    /// Setting Radius writes the correct EMU value to the XML attribute.
    /// </summary>
    [Fact]
    public void Radius_Set_WritesEmuAttribute()
    {
        var softEdge = CreateFromXml();
        softEdge.Radius = 10.0f;

        softEdge.Element!.Attribute("rad")!.Value.Should().Be("127000");
    }

    /// <summary>
    /// Radius defaults to 0 when the rad attribute is absent.
    /// </summary>
    [Fact]
    public void Radius_DefaultsToZero_WhenAttributeAbsent()
    {
        var softEdge = CreateFromXml();
        softEdge.Radius.Should().Be(0f);
    }

    /// <summary>
    /// AsIImageTransformOperation returns the instance itself.
    /// </summary>
    [Fact]
    public void AsIImageTransformOperation_ReturnsSelf()
    {
        var softEdge = CreateFromXml();
        softEdge.AsIImageTransformOperation.Should().BeSameAs(softEdge);
    }

    /// <summary>
    /// SoftEdge implements ISoftEdge and IImageTransformOperation.
    /// </summary>
    [Fact]
    public void SoftEdge_ImplementsExpectedInterfaces()
    {
        var softEdge = CreateFromXml();
        softEdge.Should().BeAssignableTo<ISoftEdge>();
        softEdge.Should().BeAssignableTo<IImageTransformOperation>();
    }

    /// <summary>
    /// Default constructor creates a SoftEdge with null element; property access does not throw.
    /// </summary>
    [Fact]
    public void DefaultConstructor_PropertiesDoNotThrow()
    {
        var softEdge = new SoftEdge();
        softEdge.Radius.Should().Be(0f);
    }
}
