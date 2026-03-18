namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a connector shape that links two shapes.
/// </summary>
public interface IConnector : IGeometryShape
{
    /// <summary>
    /// Gets the connector lock settings.
    /// </summary>
    IConnectorLock? ConnectorLock { get; }

    /// <summary>
    /// Gets or sets the shape connected at the start of this connector.
    /// </summary>
    IShape? StartShapeConnectedTo { get; set; }

    /// <summary>
    /// Gets or sets the shape connected at the end of this connector.
    /// </summary>
    IShape? EndShapeConnectedTo { get; set; }

    /// <summary>
    /// Gets or sets the connection site index on the start shape.
    /// </summary>
    int StartShapeConnectionSiteIndex { get; set; }

    /// <summary>
    /// Gets or sets the connection site index on the end shape.
    /// </summary>
    int EndShapeConnectionSiteIndex { get; set; }

    /// <summary>
    /// Returns this shape as an <see cref="IGeometryShape"/>.
    /// </summary>
    IGeometryShape AsIGeometryShape { get; }

    /// <summary>
    /// Recalculates the connector path between connected shapes.
    /// </summary>
    void Reroute();
}
