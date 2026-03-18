using Aspose.Slides.Foss.Drawing;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a collection of gradient stops.
/// </summary>
public interface IGradientStopCollection
{
    /// <summary>
    /// Gets all gradient stops as a list.
    /// </summary>
    IList<IGradientStop> AsICollection { get; }

    /// <summary>
    /// Gets an enumerable view of the gradient stops.
    /// </summary>
    IEnumerable<IGradientStop> AsIEnumerable { get; }

    /// <summary>
    /// Gets the number of gradient stops in the collection.
    /// </summary>
    int Count { get; }

    /// <summary>
    /// Gets the gradient stop at the specified index.
    /// </summary>
    /// <param name="index">The zero-based index of the gradient stop.</param>
    IGradientStop this[int index] { get; }

    /// <summary>
    /// Adds a gradient stop with the specified position and RGB color.
    /// </summary>
    /// <param name="position">The position of the stop (0.0 to 1.0).</param>
    /// <param name="color">The RGB color.</param>
    /// <returns>The newly created gradient stop.</returns>
    IGradientStop Add(float position, Color color);

    /// <summary>
    /// Adds a gradient stop with the specified position and preset color.
    /// </summary>
    /// <param name="position">The position of the stop (0.0 to 1.0).</param>
    /// <param name="presetColor">The preset color.</param>
    /// <returns>The newly created gradient stop.</returns>
    IGradientStop Add(float position, PresetColor presetColor);

    /// <summary>
    /// Adds a gradient stop with the specified position and scheme color.
    /// </summary>
    /// <param name="position">The position of the stop (0.0 to 1.0).</param>
    /// <param name="schemeColor">The scheme color.</param>
    /// <returns>The newly created gradient stop.</returns>
    IGradientStop Add(float position, SchemeColor schemeColor);

    /// <summary>
    /// Inserts a gradient stop at the specified index with an RGB color.
    /// </summary>
    /// <param name="index">The zero-based index at which to insert.</param>
    /// <param name="position">The position of the stop (0.0 to 1.0).</param>
    /// <param name="color">The RGB color.</param>
    void Insert(int index, float position, Color color);

    /// <summary>
    /// Inserts a gradient stop at the specified index with a preset color.
    /// </summary>
    /// <param name="index">The zero-based index at which to insert.</param>
    /// <param name="position">The position of the stop (0.0 to 1.0).</param>
    /// <param name="presetColor">The preset color.</param>
    void Insert(int index, float position, PresetColor presetColor);

    /// <summary>
    /// Inserts a gradient stop at the specified index with a scheme color.
    /// </summary>
    /// <param name="index">The zero-based index at which to insert.</param>
    /// <param name="position">The position of the stop (0.0 to 1.0).</param>
    /// <param name="schemeColor">The scheme color.</param>
    void Insert(int index, float position, SchemeColor schemeColor);

    /// <summary>
    /// Removes the gradient stop at the specified index.
    /// </summary>
    /// <param name="index">The zero-based index of the stop to remove.</param>
    void RemoveAt(int index);

    /// <summary>
    /// Removes all gradient stops from the collection.
    /// </summary>
    void Clear();
}
