namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a layout slide.
/// </summary>
public interface ILayoutSlide : IBaseSlide
{
    /// <summary>
    /// Returns or sets the master slide for a layout. Read/write.
    /// </summary>
    IMasterSlide? MasterSlide { get; set; }

    /// <summary>
    /// Returns the layout type of this layout slide. Read-only.
    /// </summary>
    SlideLayoutType LayoutType { get; }
}
