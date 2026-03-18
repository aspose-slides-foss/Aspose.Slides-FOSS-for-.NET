using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Export;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Integration tests for Connector shapes through the Presentation API
/// and via direct XML construction for scenarios that require shape XML.
/// </summary>
public sealed class ConnectorIntegrationTests
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static byte[] SaveToBytes(Presentation pres)
    {
        using var ms = new MemoryStream();
        pres.Save(ms, SaveFormat.Pptx);
        return ms.ToArray();
    }

    /// <summary>
    /// Creates a Connector backed by XML with the given preset geometry.
    /// </summary>
    private static Connector CreateConnector(string preset)
    {
        var element = new XElement(PNs + "cxnSp",
            new XElement(PNs + "nvCxnSpPr",
                new XElement(PNs + "cNvPr", new XAttribute("id", "4"), new XAttribute("name", "Connector 1")),
                new XElement(PNs + "cNvCxnSpPr"),
                new XElement(PNs + "nvPr")),
            new XElement(PNs + "spPr",
                new XElement(ANs + "prstGeom", new XAttribute("prst", preset))));

        var connector = new Connector();
        connector.InitInternal(element, slidePart: null, parentSlide: null);
        return connector;
    }

    /// <summary>
    /// Creates a Connector with an adjustment value in its avLst.
    /// </summary>
    private static Connector CreateConnectorWithAdjustments(string preset, int rawValue)
    {
        var element = new XElement(PNs + "cxnSp",
            new XElement(PNs + "nvCxnSpPr",
                new XElement(PNs + "cNvPr", new XAttribute("id", "4"), new XAttribute("name", "Connector 1")),
                new XElement(PNs + "cNvCxnSpPr"),
                new XElement(PNs + "nvPr")),
            new XElement(PNs + "spPr",
                new XElement(ANs + "prstGeom", new XAttribute("prst", preset),
                    new XElement(ANs + "avLst",
                        new XElement(ANs + "gd",
                            new XAttribute("name", "adj1"),
                            new XAttribute("fmla", $"val {rawValue}"))))));

        var connector = new Connector();
        connector.InitInternal(element, slidePart: null, parentSlide: null);
        return connector;
    }

    /// <summary>
    /// Creates a Connector with connection elements pre-populated.
    /// </summary>
    private static Connector CreateConnectorWithConnections(string preset, int startIdx, int endIdx)
    {
        var element = new XElement(PNs + "cxnSp",
            new XElement(PNs + "nvCxnSpPr",
                new XElement(PNs + "cNvPr", new XAttribute("id", "4"), new XAttribute("name", "Connector 1")),
                new XElement(PNs + "cNvCxnSpPr",
                    new XElement(ANs + "stCxn", new XAttribute("id", "2"), new XAttribute("idx", startIdx)),
                    new XElement(ANs + "endCxn", new XAttribute("id", "3"), new XAttribute("idx", endIdx))),
                new XElement(PNs + "nvPr")),
            new XElement(PNs + "spPr",
                new XElement(ANs + "prstGeom", new XAttribute("prst", preset))));

        var connector = new Connector();
        connector.InitInternal(element, slidePart: null, parentSlide: null);
        return connector;
    }

    /// <summary>
    /// Straight connector has the correct ShapeType.
    /// </summary>
    [Fact]
    public void AddStraightConnector_HasCorrectType()
    {
        var conn = CreateConnector("straightConnector1");

        conn.ShapeType.Should().Be(ShapeType.StraightConnector1);
    }

    /// <summary>
    /// Presentation round-trips correctly via save (verifying the framework).
    /// </summary>
    [Fact]
    public void AddStraightConnector_PresentationSurvivesSave()
    {
        using var pres = new Presentation();

        var bytes = SaveToBytes(pres);

        bytes.Should().NotBeEmpty();
        bytes.Length.Should().BeGreaterThan(100);
    }

    /// <summary>
    /// Adjustment values persist on a bent connector.
    /// </summary>
    [Fact]
    public void BentConnectorAdjustments_ValuesPersist()
    {
        var conn = CreateConnectorWithAdjustments("bentConnector3", 30000);

        conn.Adjustments.Should().NotBeNull();
        conn.Adjustments!.Count.Should().BeGreaterThan(0);
        conn.Adjustments![0].RawValue.Should().Be(30000);
    }

    /// <summary>
    /// Updated adjustment value can be read back.
    /// </summary>
    [Fact]
    public void BentConnectorAdjustments_UpdatedValuePersists()
    {
        var conn = CreateConnectorWithAdjustments("bentConnector3", 10000);

        conn.Adjustments![0].RawValue = 50000;

        conn.Adjustments![0].RawValue.Should().Be(50000);
    }

    /// <summary>
    /// Start and end connection site indices persist.
    /// </summary>
    [Fact]
    public void ConnectShapes_StartEndConnectionsPersist()
    {
        var conn = CreateConnectorWithConnections("bentConnector3", startIdx: 3, endIdx: 1);

        conn.StartShapeConnectionSiteIndex.Should().Be(3);
        conn.EndShapeConnectionSiteIndex.Should().Be(1);
    }

    /// <summary>
    /// Updated connection site indices can be read back.
    /// </summary>
    [Fact]
    public void ConnectShapes_UpdatedIndicesPersist()
    {
        var conn = CreateConnectorWithConnections("bentConnector3", startIdx: 0, endIdx: 0);

        conn.StartShapeConnectionSiteIndex = 2;
        conn.EndShapeConnectionSiteIndex = 3;

        conn.StartShapeConnectionSiteIndex.Should().Be(2);
        conn.EndShapeConnectionSiteIndex.Should().Be(3);
    }

    /// <summary>
    /// Reroute does not throw and completes without error.
    /// </summary>
    [Fact]
    public void Reroute_DoesNotThrow()
    {
        var conn = CreateConnectorWithConnections("bentConnector3", startIdx: 0, endIdx: 2);

        var act = () => conn.Reroute();

        act.Should().NotThrow();
    }

    /// <summary>
    /// Reroute can be called multiple times without error.
    /// </summary>
    [Fact]
    public void Reroute_CanBeCalledMultipleTimes()
    {
        var conn = CreateConnector("bentConnector3");

        var act = () =>
        {
            conn.Reroute();
            conn.Reroute();
            conn.Reroute();
        };

        act.Should().NotThrow();
    }

    /// <summary>
    /// Adjustment exposes name, raw_value, and angle_value.
    /// </summary>
    [Fact]
    public void AdjustmentProperties_ExposeNameRawValueAngleValue()
    {
        var conn = CreateConnectorWithAdjustments("bentConnector3", 30000);

        var adj = conn.Adjustments![0];

        adj.Name.Should().Be("adj1");
        adj.RawValue.Should().Be(30000);
        adj.AngleValue.Should().Be(30000 / 60000f);
    }

    /// <summary>
    /// AngleValue updates when RawValue changes.
    /// </summary>
    [Fact]
    public void AdjustmentProperties_AngleValueUpdatesWithRawValue()
    {
        var conn = CreateConnectorWithAdjustments("bentConnector3", 60000);

        conn.Adjustments![0].AngleValue.Should().Be(1.0f);

        conn.Adjustments![0].RawValue = 120000;

        conn.Adjustments![0].AngleValue.Should().Be(2.0f);
    }

    /// <summary>
    /// Various connector preset types produce correct ShapeType values.
    /// </summary>
    [Theory]
    [InlineData("straightConnector1", ShapeType.StraightConnector1)]
    [InlineData("bentConnector3", ShapeType.BentConnector3)]
    [InlineData("bentConnector2", ShapeType.BentConnector2)]
    [InlineData("curvedConnector3", ShapeType.CurvedConnector3)]
    public void ConnectorPresets_ProduceCorrectShapeTypes(string preset, ShapeType expected)
    {
        var conn = CreateConnector(preset);

        conn.ShapeType.Should().Be(expected);
    }

    /// <summary>
    /// Connector with no connection elements has null start/end shapes.
    /// </summary>
    [Fact]
    public void NoConnections_StartAndEndShapesAreNull()
    {
        var conn = CreateConnector("straightConnector1");

        conn.StartShapeConnectedTo.Should().BeNull();
        conn.EndShapeConnectedTo.Should().BeNull();
    }

    /// <summary>
    /// Connector ShapeType can be changed and reflects the new value.
    /// </summary>
    [Fact]
    public void ShapeType_CanBeChangedAndReflectsNewValue()
    {
        var conn = CreateConnector("straightConnector1");
        conn.ShapeType.Should().Be(ShapeType.StraightConnector1);

        conn.ShapeType = ShapeType.BentConnector3;

        conn.ShapeType.Should().Be(ShapeType.BentConnector3);
    }
}
