using System.Xml.Linq;
using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests Connector shapes, adjustments, and connections.
/// </summary>
public sealed class ConnectorTests
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

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
    /// Add a straight connector with correct type.
    /// </summary>
    [Fact]
    public void AddStraightConnector_HasCorrectShapeType()
    {
        var conn = CreateConnector("straightConnector1");

        conn.ShapeType.Should().Be(ShapeType.StraightConnector1);
    }

    /// <summary>
    /// ShapeType enum value for straight connector is defined.
    /// </summary>
    [Fact]
    public void StraightConnector1_ShapeTypeIsDefined()
    {
        Enum.IsDefined(ShapeType.StraightConnector1).Should().BeTrue();
        ShapeType.StraightConnector1.Should().NotBe(ShapeType.NotDefined);
    }

    /// <summary>
    /// Adjustments collection is accessible on a bent connector.
    /// </summary>
    [Fact]
    public void BentConnector_AdjustmentsIsNotNull()
    {
        var conn = CreateConnector("bentConnector3");

        conn.Adjustments.Should().NotBeNull();
    }

    /// <summary>
    /// Bent connector ShapeType is correct.
    /// </summary>
    [Fact]
    public void BentConnector3_HasCorrectShapeType()
    {
        var conn = CreateConnector("bentConnector3");

        conn.ShapeType.Should().Be(ShapeType.BentConnector3);
    }

    /// <summary>
    /// Start/end connection site indices can be read from XML.
    /// </summary>
    [Fact]
    public void ConnectShapes_ConnectionSiteIndicesAreReadFromXml()
    {
        var conn = CreateConnectorWithConnections("bentConnector3", startIdx: 3, endIdx: 1);

        conn.StartShapeConnectionSiteIndex.Should().Be(3);
        conn.EndShapeConnectionSiteIndex.Should().Be(1);
    }

    /// <summary>
    /// Connection site indices default to zero when no connection elements exist.
    /// </summary>
    [Fact]
    public void ConnectionSiteIndex_DefaultsToZero()
    {
        var conn = CreateConnector("bentConnector3");

        conn.StartShapeConnectionSiteIndex.Should().Be(0);
        conn.EndShapeConnectionSiteIndex.Should().Be(0);
    }

    /// <summary>
    /// Connection site index can be updated in-place.
    /// </summary>
    [Fact]
    public void ConnectShapes_ConnectionSiteIndexCanBeUpdated()
    {
        var conn = CreateConnectorWithConnections("bentConnector3", startIdx: 0, endIdx: 0);

        conn.StartShapeConnectionSiteIndex = 3;
        conn.EndShapeConnectionSiteIndex = 1;

        conn.StartShapeConnectionSiteIndex.Should().Be(3);
        conn.EndShapeConnectionSiteIndex.Should().Be(1);
    }

    /// <summary>
    /// Reroute() does not throw when no shapes are connected.
    /// </summary>
    [Fact]
    public void Reroute_DoesNotThrow()
    {
        var conn = CreateConnector("bentConnector3");

        var act = () => conn.Reroute();

        act.Should().NotThrow();
    }

    /// <summary>
    /// Adjustments expose name, raw_value, angle_value.
    /// Verifies that BentConnector3 ShapeType enum is distinct.
    /// </summary>
    [Fact]
    public void AdjustmentProperties_BentConnector3IsDefined()
    {
        Enum.IsDefined(ShapeType.BentConnector3).Should().BeTrue();
        ShapeType.BentConnector3.Should().NotBe(ShapeType.NotDefined);
        ShapeType.BentConnector3.Should().NotBe(ShapeType.StraightConnector1);
    }

    /// <summary>
    /// Various connector ShapeType values are preserved.
    /// </summary>
    [Theory]
    [InlineData("straightConnector1", ShapeType.StraightConnector1)]
    [InlineData("bentConnector3", ShapeType.BentConnector3)]
    [InlineData("bentConnector2", ShapeType.BentConnector2)]
    [InlineData("curvedConnector3", ShapeType.CurvedConnector3)]
    public void ConnectorShapeType_MatchesPreset(string preset, ShapeType expected)
    {
        var conn = CreateConnector(preset);
        conn.ShapeType.Should().Be(expected);
    }

    /// <summary>
    /// Connector with null element returns NotDefined for ShapeType.
    /// </summary>
    [Fact]
    public void NullElement_ShapeTypeIsNotDefined()
    {
        var conn = new Connector();
        conn.InitInternal(null, slidePart: null, parentSlide: null);

        conn.ShapeType.Should().Be(ShapeType.NotDefined);
    }

    /// <summary>
    /// Connector with null element returns null for Adjustments.
    /// </summary>
    [Fact]
    public void NullElement_AdjustmentsIsNull()
    {
        var conn = new Connector();
        conn.InitInternal(null, slidePart: null, parentSlide: null);

        conn.Adjustments.Should().BeNull();
    }

    /// <summary>
    /// Adjustment RawValue can be read from XML guide definition.
    /// </summary>
    [Fact]
    public void BentConnector_AdjustmentRawValueIsReadFromXml()
    {
        var conn = CreateConnectorWithAdjustments("bentConnector3", 30000);

        conn.Adjustments!.Count.Should().BeGreaterThan(0);
        conn.Adjustments![0].RawValue.Should().Be(30000);
    }

    /// <summary>
    /// Adjustment RawValue can be updated in-place.
    /// </summary>
    [Fact]
    public void BentConnector_AdjustmentRawValueCanBeUpdated()
    {
        var conn = CreateConnectorWithAdjustments("bentConnector3", 10000);

        conn.Adjustments![0].RawValue = 30000;

        conn.Adjustments![0].RawValue.Should().Be(30000);
    }

    /// <summary>
    /// Adjustment exposes name, raw_value, and angle_value.
    /// </summary>
    [Fact]
    public void AdjustmentProperties_NameRawValueAngleValueAreAccessible()
    {
        var conn = CreateConnectorWithAdjustments("bentConnector3", 30000);

        var adj = conn.Adjustments![0];
        adj.Name.Should().NotBeNull();
        adj.RawValue.Should().Be(30000);
        adj.AngleValue.Should().Be(30000 / 60000f);
    }

    /// <summary>
    /// AngleValue is derived from RawValue with factor 60000.
    /// </summary>
    [Fact]
    public void AdjustmentProperties_AngleValueIsRawValueDividedByFactor()
    {
        var conn = CreateConnectorWithAdjustments("bentConnector3", 60000);

        conn.Adjustments![0].AngleValue.Should().Be(1.0f);
    }

    /// <summary>
    /// StartShapeConnectedTo is null when no connection exists.
    /// </summary>
    [Fact]
    public void StartShapeConnectedTo_NullWhenNoConnection()
    {
        var conn = CreateConnector("bentConnector3");

        conn.StartShapeConnectedTo.Should().BeNull();
    }

    /// <summary>
    /// EndShapeConnectedTo is null when no connection exists.
    /// </summary>
    [Fact]
    public void EndShapeConnectedTo_NullWhenNoConnection()
    {
        var conn = CreateConnector("bentConnector3");

        conn.EndShapeConnectedTo.Should().BeNull();
    }

    /// <summary>
    /// IsTextHolder is always false for connectors.
    /// </summary>
    [Fact]
    public void Connector_IsTextHolderIsFalse()
    {
        var conn = CreateConnector("straightConnector1");

        conn.IsTextHolder.Should().BeFalse();
    }

    /// <summary>
    /// ShapeType can be changed on an existing connector.
    /// </summary>
    [Fact]
    public void ShapeType_CanBeChanged()
    {
        var conn = CreateConnector("straightConnector1");
        conn.ShapeType.Should().Be(ShapeType.StraightConnector1);

        conn.ShapeType = ShapeType.BentConnector3;
        conn.ShapeType.Should().Be(ShapeType.BentConnector3);
    }
}
