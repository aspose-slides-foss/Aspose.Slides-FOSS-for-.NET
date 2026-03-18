using System.Xml.Linq;
using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests ShapeBevel: construction, BevelType, Width, Height.
/// </summary>
public sealed class ShapeBevelTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static ShapeBevel CreateBevel(string? preset = null, int? width = null, int? height = null)
    {
        var element = new XElement(ANs + "bevelT");
        if (preset is not null)
            element.SetAttributeValue("prst", preset);
        if (width is not null)
            element.SetAttributeValue("w", width);
        if (height is not null)
            element.SetAttributeValue("h", height);

        var bevel = new ShapeBevel();
        bevel.InitInternal(element, slidePart: null, parentSlide: null);
        return bevel;
    }

    /// <summary>
    /// ShapeBevel can be instantiated with default constructor.
    /// </summary>
    [Fact]
    public void Init_DefaultConstructor_CreatesInstance()
    {
        var bevel = new ShapeBevel();
        bevel.Should().NotBeNull();
    }

    /// <summary>
    /// ShapeBevel can be instantiated with isTopBevel parameter.
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Init_WithIsTopBevel_CreatesInstance(bool isTop)
    {
        var bevel = new ShapeBevel(isTop);
        bevel.Should().NotBeNull();
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void BevelType_CircleIsReadFromXml()
    {
        var bevel = CreateBevel(preset: "circle");

        bevel.BevelType.Should().Be(BevelPresetType.Circle);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void BevelType_CanBeSet()
    {
        var bevel = CreateBevel();

        bevel.BevelType = BevelPresetType.Circle;

        bevel.BevelType.Should().Be(BevelPresetType.Circle);
    }

    /// <summary>
    /// BevelType defaults to NotDefined when no preset attribute exists.
    /// </summary>
    [Fact]
    public void BevelType_DefaultsToNotDefined()
    {
        var bevel = CreateBevel();

        bevel.BevelType.Should().Be(BevelPresetType.NotDefined);
    }

    /// <summary>
    /// Setting BevelType to NotDefined removes the attribute.
    /// </summary>
    [Fact]
    public void BevelType_SetToNotDefined_RemovesAttribute()
    {
        var bevel = CreateBevel(preset: "circle");
        bevel.BevelType.Should().Be(BevelPresetType.Circle);

        bevel.BevelType = BevelPresetType.NotDefined;

        bevel.BevelType.Should().Be(BevelPresetType.NotDefined);
    }

    /// <summary>
    /// Various bevel presets are correctly read from XML.
    /// </summary>
    [Theory]
    [InlineData("circle", BevelPresetType.Circle)]
    [InlineData("angle", BevelPresetType.Angle)]
    [InlineData("softRound", BevelPresetType.SoftRound)]
    [InlineData("convex", BevelPresetType.Convex)]
    [InlineData("hardEdge", BevelPresetType.HardEdge)]
    public void BevelType_MatchesXmlPreset(string preset, BevelPresetType expected)
    {
        var bevel = CreateBevel(preset: preset);

        bevel.BevelType.Should().Be(expected);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void Width_IsReadFromXml()
    {
        var bevel = CreateBevel(width: 127000); // 10 * 12700

        bevel.Width.Should().Be(10f);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void Width_CanBeSet()
    {
        var bevel = CreateBevel();

        bevel.Width = 10;

        bevel.Width.Should().Be(10f);
    }

    /// <summary>
    /// Width defaults to zero when no attribute exists.
    /// </summary>
    [Fact]
    public void Width_DefaultsToZero()
    {
        var bevel = CreateBevel();

        bevel.Width.Should().Be(0f);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void Height_IsReadFromXml()
    {
        var bevel = CreateBevel(height: 63500); // 5 * 12700

        bevel.Height.Should().Be(5f);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void Height_CanBeSet()
    {
        var bevel = CreateBevel();

        bevel.Height = 5;

        bevel.Height.Should().Be(5f);
    }

    /// <summary>
    /// Height defaults to zero when no attribute exists.
    /// </summary>
    [Fact]
    public void Height_DefaultsToZero()
    {
        var bevel = CreateBevel();

        bevel.Height.Should().Be(0f);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void BevelTop_AllPropertiesPersistTogether()
    {
        var bevel = CreateBevel(preset: "circle", width: 127000, height: 63500);

        bevel.BevelType.Should().Be(BevelPresetType.Circle);
        bevel.Width.Should().Be(10f);
        bevel.Height.Should().Be(5f);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void BevelTop_SetAllProperties_ReadBack()
    {
        var bevel = CreateBevel();

        bevel.BevelType = BevelPresetType.Circle;
        bevel.Width = 10;
        bevel.Height = 5;

        bevel.BevelType.Should().Be(BevelPresetType.Circle);
        bevel.Width.Should().Be(10f);
        bevel.Height.Should().Be(5f);
    }

    /// <summary>
    /// Properties return defaults when initialized with null element.
    /// </summary>
    [Fact]
    public void NullElement_PropertiesReturnDefaults()
    {
        var bevel = new ShapeBevel();
        bevel.InitInternal(null!, slidePart: null, parentSlide: null);

        bevel.BevelType.Should().Be(BevelPresetType.NotDefined);
        bevel.Width.Should().Be(0f);
        bevel.Height.Should().Be(0f);
    }
}
