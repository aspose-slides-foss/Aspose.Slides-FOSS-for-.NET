using System.Xml.Linq;

namespace Aspose.Slides.Foss.Internal;

/// <summary>
/// Creates shape wrapper objects from XML elements found in a slide's shape tree.
/// Delegates to <see cref="Pptx.ShapeFactory"/> for the actual creation logic.
/// </summary>
internal static class ShapeFactory
{
    /// <summary>
    /// Creates an <see cref="IShape"/> wrapper for the given XML element.
    /// Returns <c>null</c> if the element type is not recognized.
    /// </summary>
    internal static IShape? CreateShape(XElement element, SlidePart? slidePart, IBaseSlide? parentSlide)
        => Pptx.ShapeFactory.CreateShape(element, slidePart, parentSlide);
}
