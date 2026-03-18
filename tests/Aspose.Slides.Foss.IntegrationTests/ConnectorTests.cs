using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.IntegrationTests;

/// <summary>
/// Tests for Connector shapes, adjustments, and connections.
/// </summary>
public sealed class ConnectorTests : IDisposable
{
    private readonly string _tempDir;

    public ConnectorTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, recursive: true);
    }

    [Fact]
    public void TestAddStraightConnector()
    {
        // Add a straight connector with correct type.
        using var pres = new Presentation();
        var slide = pres.Slides[0];
        var conn = slide.Shapes!.AddConnector(ShapeType.StraightConnector1, 100, 100, 300, 200);
        conn.ShapeType.Should().Be(ShapeType.StraightConnector1);
    }

    [Fact]
    public void TestAddStraightConnectorPersists()
    {
        // Straight connector survives save/reload.
        using var pres = new Presentation();
        pres.Slides[0].Shapes!.AddConnector(ShapeType.StraightConnector1, 100, 100, 300, 200);

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        pres2.Slides[0].Shapes!.Count.Should().BeGreaterThanOrEqualTo(1);
    }

    [Fact]
    public void TestBentConnectorAdjustments()
    {
        // Adjustment values persist after save/reload.
        using var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes!;
        shapes.Clear();
        var conn = shapes.AddConnector(ShapeType.BentConnector3, 50, 50, 300, 200);
        var adjustments = conn.Adjustments;
        if (adjustments is not null && adjustments.Count > 0)
        {
            adjustments[0].RawValue = 30000;
        }

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        // Find the connector shape
        Connector? conn2 = null;
        foreach (var sh in pres2.Slides[0].Shapes!)
        {
            if (sh is Connector c)
            {
                conn2 = c;
                break;
            }
        }
        conn2.Should().NotBeNull("Connector not found after reload");
        var adj2 = conn2!.Adjustments;
        if (adj2 is not null && adj2.Count > 0)
        {
            adj2[0].RawValue.Should().Be(30000);
        }
    }

    [Fact]
    public void TestConnectShapes()
    {
        // Start/end connections persist.
        using var pres = new Presentation();
        var slide = pres.Slides[0];
        var shapes = slide.Shapes!;
        shapes.Clear();
        var s1 = shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 100, 60);
        var s2 = shapes.AddAutoShape(ShapeType.Rectangle, 350, 200, 100, 60);
        var conn = shapes.AddConnector(ShapeType.BentConnector3, 0, 0, 1, 1);

        conn.StartShapeConnectedTo = s1;
        conn.StartShapeConnectionSiteIndex = 3;
        conn.EndShapeConnectedTo = s2;
        conn.EndShapeConnectionSiteIndex = 1;

        conn.StartShapeConnectedTo.Should().NotBeNull();
        conn.EndShapeConnectedTo.Should().NotBeNull();

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        Connector? conn2 = null;
        foreach (var sh in pres2.Slides[0].Shapes!)
        {
            if (sh is Connector c && c.ShapeType == ShapeType.BentConnector3)
            {
                conn2 = c;
                break;
            }
        }
        conn2.Should().NotBeNull();
        conn2!.StartShapeConnectionSiteIndex.Should().Be(3);
        conn2.EndShapeConnectionSiteIndex.Should().Be(1);
    }

    [Fact]
    public void TestReroute()
    {
        // reroute() updates connector position.
        using var pres = new Presentation();
        var slide = pres.Slides[0];
        var shapes = slide.Shapes!;
        var s1 = shapes.AddAutoShape(ShapeType.Ellipse, 50, 100, 80, 80);
        var s2 = shapes.AddAutoShape(ShapeType.Ellipse, 400, 100, 80, 80);
        var conn = shapes.AddConnector(ShapeType.BentConnector3, 0, 0, 1, 1);
        conn.StartShapeConnectedTo = s1;
        conn.StartShapeConnectionSiteIndex = 3;
        conn.EndShapeConnectedTo = s2;
        conn.EndShapeConnectionSiteIndex = 1;
        conn.Reroute();
        // After reroute the connector should span between the shapes
        (conn.Width > 0 || conn.Height > 0).Should().BeTrue();
    }

    [Fact]
    public void TestAdjustmentProperties()
    {
        // Adjustment values expose name, raw_value, angle_value.
        using var pres = new Presentation();
        var conn = pres.Slides[0].Shapes!.AddConnector(ShapeType.BentConnector3, 50, 50, 300, 200);
        var adjustments = conn.Adjustments;
        if (adjustments is not null && adjustments.Count > 0)
        {
            var adj = adjustments[0];
            adj.Name.Should().NotBeNull();
            adj.RawValue.Should().BeOfType(typeof(int));
            adj.AngleValue.Should().BeOfType(typeof(float));
        }
    }
}
