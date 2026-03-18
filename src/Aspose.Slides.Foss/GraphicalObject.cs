namespace Aspose.Slides.Foss;

/// <summary>
/// Abstract base class for graphical objects on a slide.
/// </summary>
public abstract class GraphicalObject : Shape, IGraphicalObject
{
    /// <inheritdoc/>
    public abstract IGraphicalObjectLock? GraphicalObjectLock { get; }
}
