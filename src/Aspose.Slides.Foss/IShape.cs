namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a shape on a slide.
/// </summary>
public interface IShape : IHyperlinkContainer
{
    /// <summary>
    /// Gets a value indicating whether the shape acts as a text holder.
    /// </summary>
    bool IsTextHolder { get; }

    /// <summary>
    /// Gets the placeholder associated with the shape, or <c>null</c> if none.
    /// </summary>
    IPlaceholder? Placeholder { get; }

    /// <summary>
    /// Gets the custom data attached to the shape, or <c>null</c> if none.
    /// </summary>
    ICustomData? CustomData { get; }

    /// <summary>
    /// Gets the line formatting properties for the shape.
    /// </summary>
    ILineFormat LineFormat { get; }

    /// <summary>
    /// Gets the 3-D formatting properties for the shape.
    /// </summary>
    IThreeDFormat ThreeDFormat { get; }

    /// <summary>
    /// Gets the effect formatting properties for the shape.
    /// </summary>
    IEffectFormat EffectFormat { get; }

    /// <summary>
    /// Gets the fill formatting properties for the shape.
    /// </summary>
    IFillFormat FillFormat { get; }

    /// <summary>
    /// Gets the position of the shape in the z-order stack.
    /// Index 0 is the backmost; <c>Count - 1</c> is the frontmost.
    /// </summary>
    int ZOrderPosition { get; }

    /// <summary>
    /// Gets the number of connection sites available on the shape.
    /// </summary>
    int ConnectionSiteCount { get; }

    /// <summary>
    /// Gets the presentation-scoped internal identifier for the shape.
    /// Not guaranteed to be persistent across saves.
    /// </summary>
    int UniqueId { get; }

    /// <summary>
    /// Gets the slide-scoped unique identifier that remains constant
    /// for the shape's lifetime. Designed for PowerPoint/interop code
    /// to reliably reference the shape.
    /// </summary>
    int OfficeInteropShapeId { get; }

    /// <summary>
    /// Gets a value indicating whether the shape belongs to a group shape.
    /// </summary>
    bool IsGrouped { get; }

    /// <summary>
    /// Returns this instance viewed as an <see cref="ISlideComponent"/>.
    /// </summary>
    ISlideComponent AsISlideComponent { get; }

    /// <summary>
    /// Gets or sets the raw (unresolved) shape frame properties.
    /// </summary>
    IShapeFrame RawFrame { get; set; }

    /// <summary>
    /// Gets or sets the resolved shape frame properties.
    /// </summary>
    IShapeFrame Frame { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the shape is hidden from display.
    /// </summary>
    bool Hidden { get; set; }

    /// <summary>
    /// Gets or sets the rotation in degrees around the z-axis.
    /// Positive values indicate clockwise rotation; negative values indicate counterclockwise.
    /// </summary>
    float Rotation { get; set; }

    /// <summary>
    /// Gets or sets the x-coordinate of the upper-left corner, in points.
    /// </summary>
    float X { get; set; }

    /// <summary>
    /// Gets or sets the y-coordinate of the upper-left corner, in points.
    /// </summary>
    float Y { get; set; }

    /// <summary>
    /// Gets or sets the width of the shape, in points.
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// Gets or sets the height of the shape, in points.
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// Gets or sets the accessibility alternative text for the shape.
    /// </summary>
    string AlternativeText { get; set; }

    /// <summary>
    /// Gets or sets the title of the alternative text.
    /// </summary>
    string AlternativeTextTitle { get; set; }

    /// <summary>
    /// Gets or sets the user-visible name of the shape.
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the shape is marked as decorative.
    /// Decorative shapes are ignored by screen readers.
    /// </summary>
    bool IsDecorative { get; set; }
}
