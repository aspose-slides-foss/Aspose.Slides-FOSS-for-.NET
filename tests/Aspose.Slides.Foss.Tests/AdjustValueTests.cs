using System.Xml.Linq;
using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests AdjustValue Name, RawValue, and AngleValue properties.
/// </summary>
public sealed class AdjustValueTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// <summary>
    /// Creates an AdjustValue backed by a guide-definition element.
    /// </summary>
    private static AdjustValue CreateAdjustValue(string name, int rawValue)
    {
        var gd = new XElement(ANs + "gd",
            new XAttribute("name", name),
            new XAttribute("fmla", $"val {rawValue}"));

        var av = new AdjustValue();
        av.InitInternal(gd, slidePart: null);
        return av;
    }

    /// <summary>
    /// Name is accessible and matches the XML attribute.
    /// </summary>
    [Fact]
    public void Name_ReturnsValueFromXml()
    {
        var av = CreateAdjustValue("adj1", 30000);

        av.Name.Should().Be("adj1");
    }

    /// <summary>
    /// Name is not null even for a valid guide element.
    /// </summary>
    [Fact]
    public void Name_IsNotNull()
    {
        var av = CreateAdjustValue("adj1", 30000);

        av.Name.Should().NotBeNull();
    }

    /// <summary>
    /// RawValue is readable and returns the integer from the formula.
    /// </summary>
    [Fact]
    public void RawValue_ReturnsIntegerFromFormula()
    {
        var av = CreateAdjustValue("adj1", 30000);

        av.RawValue.Should().Be(30000);
    }

    /// <summary>
    /// RawValue can be updated in-place and reflects the new value.
    /// </summary>
    [Fact]
    public void RawValue_CanBeUpdated()
    {
        var av = CreateAdjustValue("adj1", 10000);

        av.RawValue = 30000;

        av.RawValue.Should().Be(30000);
    }

    /// <summary>
    /// AngleValue is derived from RawValue divided by 60000.
    /// </summary>
    [Fact]
    public void AngleValue_IsDerivedFromRawValue()
    {
        var av = CreateAdjustValue("adj1", 30000);

        av.AngleValue.Should().Be(30000 / 60000f);
    }

    /// <summary>
    /// AngleValue of 1.0 corresponds to RawValue of 60000.
    /// </summary>
    [Fact]
    public void AngleValue_OneCorrespondsToRawValue60000()
    {
        var av = CreateAdjustValue("adj1", 60000);

        av.AngleValue.Should().Be(1.0f);
    }

    /// <summary>
    /// AngleValue setter updates RawValue correctly.
    /// </summary>
    [Fact]
    public void AngleValue_SetterUpdatesRawValue()
    {
        var av = CreateAdjustValue("adj1", 0);

        av.AngleValue = 0.5f;

        av.RawValue.Should().Be(30000);
    }

    /// <summary>
    /// Default-constructed AdjustValue without Init returns defaults.
    /// </summary>
    [Fact]
    public void DefaultConstruction_ReturnsDefaults()
    {
        var av = new AdjustValue();

        av.Name.Should().BeEmpty();
        av.RawValue.Should().Be(0);
        av.AngleValue.Should().Be(0f);
    }

    /// <summary>
    /// RawValue persists through the underlying XML element.
    /// </summary>
    [Fact]
    public void RawValue_PersistsThroughXml()
    {
        var gd = new XElement(ANs + "gd",
            new XAttribute("name", "adj1"),
            new XAttribute("fmla", "val 10000"));

        var av = new AdjustValue();
        av.InitInternal(gd, slidePart: null);
        av.RawValue = 30000;

        // Re-read from the same XML element via a new AdjustValue
        var av2 = new AdjustValue();
        av2.InitInternal(gd, slidePart: null);
        av2.RawValue.Should().Be(30000);
    }

    /// <summary>
    /// Various raw values produce correct angle values.
    /// </summary>
    [Theory]
    [InlineData(0, 0f)]
    [InlineData(30000, 0.5f)]
    [InlineData(60000, 1.0f)]
    [InlineData(90000, 1.5f)]
    public void AngleValue_CorrespondsToRawValue(int rawValue, float expectedAngle)
    {
        var av = CreateAdjustValue("adj1", rawValue);

        av.AngleValue.Should().Be(expectedAngle);
    }
}
