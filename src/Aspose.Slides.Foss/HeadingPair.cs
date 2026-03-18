namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a 'Heading pair' property of the document.
/// It indicates the group name of document parts and the number of parts in group.
/// </summary>
public sealed class HeadingPair : IHeadingPair
{
    /// <inheritdoc />
    public string Name { get; private set; } = string.Empty;

    /// <inheritdoc />
    public int Count { get; private set; }

    /// <summary>
    /// Initializes the heading pair with the specified group name and part count.
    /// </summary>
    /// <param name="name">The group name of document parts.</param>
    /// <param name="count">The number of parts in the group.</param>
    internal void InitInternal(string name, int count)
    {
        Name = name;
        Count = count;
    }

    internal HeadingPair(string name, int count)
    {
        Name = name;
        Count = count;
    }
}
