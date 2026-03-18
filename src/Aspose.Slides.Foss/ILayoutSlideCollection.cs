using System.Collections;
using System.Collections.Generic;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a base class for collection of layout slides.
/// </summary>
public interface ILayoutSlideCollection : IEnumerable<ILayoutSlide>
{
    /// <summary>
    /// Returns the layout slide at the specified index. Read-only.
    /// </summary>
    /// <param name="index">The zero-based index of the layout slide to get.</param>
    ILayoutSlide this[int index] { get; }

    /// <summary>
    /// Returns this collection as an <see cref="ICollection"/> interface. Read-only.
    /// </summary>
    ICollection AsICollection { get; }

    /// <summary>
    /// Returns this collection as an <see cref="IEnumerable"/> interface. Read-only.
    /// </summary>
    IEnumerable AsIEnumerable { get; }

    /// <summary>
    /// Returns the first layout slide of the specified type.
    /// </summary>
    /// <param name="type">The type of layout slide to find.</param>
    /// <returns>The <see cref="ILayoutSlide"/> with the specified type, or <c>null</c> if not found.</returns>
    ILayoutSlide? GetByType(SlideLayoutType type);
}
