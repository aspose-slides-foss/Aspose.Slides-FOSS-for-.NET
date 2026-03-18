namespace Aspose.Slides.Foss;

/// <summary>
/// Represents an AutoShape.
/// </summary>
public interface IAutoShape : IGeometryShape
{
    /// <summary>
    /// Gets the text frame contained in this shape.
    /// </summary>
    ITextFrame? TextFrame { get; }

    /// <summary>
    /// Gets a value indicating whether this shape is a text box.
    /// </summary>
    bool IsTextBox { get; }

    /// <summary>
    /// Returns this shape as an <see cref="IGeometryShape"/>.
    /// </summary>
    IGeometryShape AsIGeometryShape { get; }

    /// <summary>
    /// Adds or replaces the text frame with the specified text.
    /// </summary>
    /// <param name="text">The text content. Line breaks split into multiple paragraphs.</param>
    /// <returns>The new text frame, or <c>null</c> if the XML element is null.</returns>
    ITextFrame? AddTextFrame(string? text);
}
