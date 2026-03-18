namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a collection of shape adjustment values.
/// </summary>
public interface IAdjustValueCollection
{
    /// <summary>
    /// Gets the adjustment value at the specified index.
    /// </summary>
    IAdjustValue this[int index] { get; }

    /// <summary>
    /// Gets the number of adjustment values.
    /// </summary>
    int Count { get; }

    /// <summary>
    /// Gets all adjustment values as a list.
    /// </summary>
    IList<IAdjustValue> AsICollection { get; }

    /// <summary>
    /// Gets all adjustment values as an enumerable.
    /// </summary>
    IEnumerable<IAdjustValue> AsIEnumerable { get; }
}
