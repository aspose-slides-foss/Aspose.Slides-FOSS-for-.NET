namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a graphical object on a slide.
/// </summary>
public interface IGraphicalObject : IShape
{
    /// <summary>
    /// Gets the graphical object lock settings.
    /// </summary>
    IGraphicalObjectLock? GraphicalObjectLock { get; }
}
