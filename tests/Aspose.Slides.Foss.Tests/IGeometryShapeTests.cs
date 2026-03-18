using System.Xml.Linq;
using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests for the IGeometryShape interface contract: ShapeStyle, ShapeType, and Adjustments.
/// </summary>
public sealed class IGeometryShapeTests
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// <summary>
    /// Creates an AutoShape (IGeometryShape) backed by XML with the given preset geometry.
    /// </summary>
    private static IGeometryShape CreateAutoShape(string preset)
    {
        var element = new XElement(PNs + "sp",
            new XElement(PNs + "nvSpPr",
                new XElement(PNs + "cNvPr", new XAttribute("id", "2"), new XAttribute("name", "Shape 1")),
                new XElement(PNs + "cNvSpPr"),
                new XElement(PNs + "nvPr")),
            new XElement(PNs + "spPr",
                new XElement(ANs + "prstGeom", new XAttribute("prst", preset))));

        var shape = new AutoShape();
        shape.InitInternal(element, slidePart: null, parentSlide: null);
        return shape;
    }

    /// <summary>
    /// Creates a Connector (IGeometryShape) backed by XML with the given preset and adjustments.
    /// </summary>
    private static IGeometryShape CreateConnectorWithAdjustments(string preset, int rawValue)
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

    // ── ShapeType property ──

    /// <summary>
    /// AutoShape implements IGeometryShape and ShapeType is correct.
    /// </summary>
    [Fact]
    public void AutoShape_ImplementsIGeometryShape()
    {
        var shape = CreateAutoShape("rect");

        shape.Should().BeAssignableTo<IGeometryShape>();
        shape.ShapeType.Should().Be(ShapeType.Rectangle);
    }

    /// <summary>
    /// Connector implements IGeometryShape and ShapeType is correct.
    /// </summary>
    [Fact]
    public void Connector_ImplementsIGeometryShape()
    {
        var element = new XElement(PNs + "cxnSp",
            new XElement(PNs + "nvCxnSpPr",
                new XElement(PNs + "cNvPr", new XAttribute("id", "4"), new XAttribute("name", "Conn 1")),
                new XElement(PNs + "cNvCxnSpPr"),
                new XElement(PNs + "nvPr")),
            new XElement(PNs + "spPr",
                new XElement(ANs + "prstGeom", new XAttribute("prst", "straightConnector1"))));

        var conn = new Connector();
        conn.InitInternal(element, slidePart: null, parentSlide: null);
        IGeometryShape geo = conn;

        geo.ShapeType.Should().Be(ShapeType.StraightConnector1);
    }

    /// <summary>
    /// Various ShapeType values are preserved through IGeometryShape.
    /// </summary>
    [Theory]
    [InlineData("rect", ShapeType.Rectangle)]
    [InlineData("ellipse", ShapeType.Ellipse)]
    [InlineData("triangle", ShapeType.Triangle)]
    public void ShapeType_PreservedForVariousPresets(string preset, ShapeType expected)
    {
        var shape = CreateAutoShape(preset);

        shape.ShapeType.Should().Be(expected);
    }

    /// <summary>
    /// ShapeType can be read and compared through the IGeometryShape interface.
    /// </summary>
    [Fact]
    public void ShapeType_CanBeComparedAcrossShapes()
    {
        var rect = CreateAutoShape("rect");
        var ellipse = CreateAutoShape("ellipse");

        rect.ShapeType.Should().NotBe(ellipse.ShapeType);
    }

    /// <summary>
    /// ShapeType set via interface setter is readable.
    /// </summary>
    [Fact]
    public void ShapeType_SetterUpdatesValue()
    {
        var shape = CreateAutoShape("rect");
        shape.ShapeType.Should().Be(ShapeType.Rectangle);

        shape.ShapeType = ShapeType.Ellipse;

        shape.ShapeType.Should().Be(ShapeType.Ellipse);
    }

    /// <summary>
    /// Note: "on value changing all adjustment values will reset to their default values."
    /// ShapeType defaults to NotDefined when no XML element is present.
    /// </summary>
    [Fact]
    public void ShapeType_DefaultsToNotDefined_WhenNoElement()
    {
        var shape = new AutoShape();
        shape.InitInternal(null, slidePart: null, parentSlide: null);
        IGeometryShape geo = shape;

        geo.ShapeType.Should().Be(ShapeType.NotDefined);
    }

    // ── Adjustments property ──

    /// <summary>
    /// Adjustments expose name, raw_value, angle_value through IGeometryShape.
    /// </summary>
    [Fact]
    public void Adjustments_ExposeNameRawValueAngleValue()
    {
        var shape = CreateConnectorWithAdjustments("bentConnector3", 30000);

        shape.Adjustments.Should().NotBeNull();
        var adj = shape.Adjustments![0];
        adj.Name.Should().NotBeNull();
        adj.RawValue.Should().Be(30000);
        adj.AngleValue.Should().Be(30000 / 60000f);
    }

    /// <summary>
    /// Adjustment RawValue can be updated through IGeometryShape.Adjustments.
    /// </summary>
    [Fact]
    public void Adjustments_RawValueCanBeUpdated()
    {
        var shape = CreateConnectorWithAdjustments("bentConnector3", 10000);

        shape.Adjustments![0].RawValue = 30000;

        shape.Adjustments![0].RawValue.Should().Be(30000);
    }

    /// <summary>
    /// Adjustments count is accessible through the interface.
    /// </summary>
    [Fact]
    public void Adjustments_CountIsAccessible()
    {
        var shape = CreateConnectorWithAdjustments("bentConnector3", 50000);

        shape.Adjustments!.Count.Should().BeGreaterThan(0);
    }

    // ── ShapeStyle property ──

    /// <summary>
    /// ShapeStyle is accessible through IGeometryShape (currently returns null).
    /// </summary>
    [Fact]
    public void ShapeStyle_IsAccessibleOnAutoShape()
    {
        var shape = CreateAutoShape("rect");

        // ShapeStyle is read-only; verifying it does not throw
        var style = shape.ShapeStyle;
        // Default implementation returns null
        style.Should().BeNull();
    }

    // ── IShape inheritance ──

    /// <summary>
    /// IGeometryShape extends IShape, so IShape members are accessible.
    /// </summary>
    [Fact]
    public void IGeometryShape_ExtendsIShape()
    {
        var shape = CreateAutoShape("rect");

        shape.Should().BeAssignableTo<IShape>();
        shape.Name.Should().Be("Shape 1");
    }
}
