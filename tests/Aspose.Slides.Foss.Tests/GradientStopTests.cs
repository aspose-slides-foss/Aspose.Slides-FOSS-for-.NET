using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests for <see cref="GradientStop"/>: Position and Color properties.
/// </summary>
public sealed class GradientStopTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// <summary>
    /// Creates a GradientStop backed by an &lt;a:gs&gt; element.
    /// </summary>
    private static GradientStop CreateStop(string? pos = null)
    {
        var gs = new XElement(ANs + "gs");
        if (pos is not null)
            gs.SetAttributeValue("pos", pos);
        var stop = new GradientStop();
        stop.InitInternal(gs, slidePart: null!, parentSlide: null);
        return stop;
    }

    // --- Position getter ---

    [Fact]
    public void Position_NoPosAttribute_ReturnsZero()
    {
        var stop = CreateStop();
        stop.Position.Should().Be(0f);
    }

    [Fact]
    public void Position_ZeroPos_ReturnsZero()
    {
        var stop = CreateStop("0");
        stop.Position.Should().Be(0f);
    }

    [Fact]
    public void Position_100000_ReturnsOne()
    {
        var stop = CreateStop("100000");
        stop.Position.Should().Be(1.0f);
    }

    [Fact]
    public void Position_50000_ReturnsHalf()
    {
        var stop = CreateStop("50000");
        stop.Position.Should().Be(0.5f);
    }

    [Theory]
    [InlineData("0", 0f)]
    [InlineData("25000", 0.25f)]
    [InlineData("50000", 0.5f)]
    [InlineData("75000", 0.75f)]
    [InlineData("100000", 1.0f)]
    public void Position_VariousXmlValues_ConvertCorrectly(string xmlPos, float expected)
    {
        var stop = CreateStop(xmlPos);
        stop.Position.Should().Be(expected);
    }

    // --- Position setter ---

    [Fact]
    public void Position_SetZero_PersistsToXml()
    {
        var stop = CreateStop();
        stop.Position = 0f;
        stop.Position.Should().Be(0f);
    }

    [Fact]
    public void Position_SetOne_PersistsToXml()
    {
        var stop = CreateStop();
        stop.Position = 1.0f;
        stop.Position.Should().Be(1.0f);
    }

    [Fact]
    public void Position_SetHalf_PersistsToXml()
    {
        var stop = CreateStop();
        stop.Position = 0.5f;
        stop.Position.Should().Be(0.5f);
    }

    [Fact]
    public void Position_SetAndRead_RoundTrips()
    {
        var stop = CreateStop();
        stop.Position = 0.73f;
        stop.Position.Should().Be(0.73f);
    }

    [Fact]
    public void Position_OverwritesExistingValue()
    {
        var stop = CreateStop("25000");
        stop.Position.Should().Be(0.25f);

        stop.Position = 0.75f;
        stop.Position.Should().Be(0.75f);
    }

    // --- Color ---

    [Fact]
    public void Color_ReturnsIColorFormat()
    {
        var stop = CreateStop();
        stop.Color.Should().NotBeNull();
        stop.Color.Should().BeAssignableTo<IColorFormat>();
    }

    [Fact]
    public void Color_WithSrgbClrChild_ReturnsCorrectColor()
    {
        var gs = new XElement(ANs + "gs",
            new XAttribute("pos", "50000"),
            new XElement(ANs + "srgbClr", new XAttribute("val", "FF0000")));
        var stop = new GradientStop();
        stop.InitInternal(gs, slidePart: null!, parentSlide: null);

        var colorFormat = stop.Color;
        colorFormat.Should().NotBeNull();
        colorFormat.R.Should().Be(255);
        colorFormat.G.Should().Be(0);
        colorFormat.B.Should().Be(0);
    }

    // --- Interface ---

    [Fact]
    public void GradientStop_ImplementsIGradientStop()
    {
        var stop = CreateStop();
        stop.Should().BeAssignableTo<IGradientStop>();
    }

    // --- Round-trip via collection (integration with GradientStopCollection) ---

    [Fact]
    public void Position_ViaCollection_RoundTrips()
    {
        var gsLst = new XElement(ANs + "gsLst");
        var coll = new GradientStopCollection();
        coll.InitInternal(gsLst, slidePart: null, parentSlide: null);

        coll.Add(0.0f, Color.Blue);
        coll.Add(1.0f, Color.Red);

        coll[0].Position.Should().Be(0.0f);
        coll[1].Position.Should().Be(1.0f);
    }

    [Fact]
    public void Position_SetViaCollectionStop_Persists()
    {
        var gsLst = new XElement(ANs + "gsLst");
        var coll = new GradientStopCollection();
        coll.InitInternal(gsLst, slidePart: null, parentSlide: null);

        var stop = (GradientStop)coll.Add(0.0f, Color.Blue);
        stop.Position = 0.33f;

        coll[0].Position.Should().BeApproximately(0.33f, 0.001f);
    }
}
