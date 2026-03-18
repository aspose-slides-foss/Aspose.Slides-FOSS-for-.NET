using System.Collections.Generic;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents an ordered, mutable collection of <see cref="IShape"/> objects
/// belonging to a slide or group shape.
/// </summary>
public interface IShapeCollection : IEnumerable<IShape>
{
    /// <summary>
    /// Gets the shape at the specified zero-based index.
    /// </summary>
    /// <param name="index">The zero-based index of the shape.</param>
    /// <returns>The shape at the specified index.</returns>
    IShape this[int index] { get; }

    /// <summary>
    /// Gets the parent group shape that owns this collection.
    /// </summary>
    IGroupShape? ParentGroup { get; }

    /// <summary>
    /// Exposes the collection as a generic list interface.
    /// </summary>
    IList<IShape> AsICollection { get; }

    /// <summary>
    /// Exposes the collection as an enumerable interface.
    /// </summary>
    IEnumerable<IShape> AsIEnumerable { get; }

    /// <summary>
    /// Returns all shapes in the collection as a new array.
    /// </summary>
    /// <returns>An array containing all shapes in the collection.</returns>
    IShape[] ToArray();

    /// <summary>
    /// Returns a subset of shapes starting at the specified index.
    /// </summary>
    /// <param name="startIndex">The zero-based start index.</param>
    /// <param name="count">The number of elements to return.</param>
    /// <returns>An array containing the specified range of shapes.</returns>
    IShape[] ToArray(int startIndex, int count);

    /// <summary>
    /// Moves a single shape to the specified position in the collection, changing its z-order.
    /// </summary>
    /// <param name="index">The target zero-based position.</param>
    /// <param name="shape">The shape to move.</param>
    void Reorder(int index, IShape shape);

    /// <summary>
    /// Moves multiple shapes so that they start at the specified position,
    /// preserving their relative order.
    /// </summary>
    /// <param name="index">The target zero-based position.</param>
    /// <param name="shapes">The shapes to move.</param>
    void Reorder(int index, IShape[] shapes);

    /// <summary>
    /// Creates a new auto shape and appends it to the collection.
    /// The shape is initialized from a default template.
    /// </summary>
    /// <param name="shapeType">The type of auto shape to create.</param>
    /// <param name="x">The x-coordinate of the upper-left corner.</param>
    /// <param name="y">The y-coordinate of the upper-left corner.</param>
    /// <param name="width">The width of the shape.</param>
    /// <param name="height">The height of the shape.</param>
    /// <returns>The created auto shape.</returns>
    IAutoShape AddAutoShape(ShapeType shapeType, float x, float y, float width, float height);

    /// <summary>
    /// Creates a new auto shape and appends it to the collection.
    /// </summary>
    /// <param name="shapeType">The type of auto shape to create.</param>
    /// <param name="x">The x-coordinate of the upper-left corner.</param>
    /// <param name="y">The y-coordinate of the upper-left corner.</param>
    /// <param name="width">The width of the shape.</param>
    /// <param name="height">The height of the shape.</param>
    /// <param name="createFromTemplate">
    /// When <c>true</c>, the shape is initialized from a default template with styling/text.
    /// When <c>false</c>, the shape is created bare/empty.
    /// </param>
    /// <returns>The created auto shape.</returns>
    IAutoShape AddAutoShape(ShapeType shapeType, float x, float y, float width, float height, bool createFromTemplate);

    /// <summary>
    /// Creates a new auto shape and inserts it at the specified position.
    /// The shape is initialized from a default template.
    /// </summary>
    /// <param name="index">The zero-based insertion index.</param>
    /// <param name="shapeType">The type of auto shape to create.</param>
    /// <param name="x">The x-coordinate of the upper-left corner.</param>
    /// <param name="y">The y-coordinate of the upper-left corner.</param>
    /// <param name="width">The width of the shape.</param>
    /// <param name="height">The height of the shape.</param>
    /// <returns>The created auto shape.</returns>
    IAutoShape InsertAutoShape(int index, ShapeType shapeType, float x, float y, float width, float height);

    /// <summary>
    /// Creates a new auto shape and inserts it at the specified position.
    /// </summary>
    /// <param name="index">The zero-based insertion index.</param>
    /// <param name="shapeType">The type of auto shape to create.</param>
    /// <param name="x">The x-coordinate of the upper-left corner.</param>
    /// <param name="y">The y-coordinate of the upper-left corner.</param>
    /// <param name="width">The width of the shape.</param>
    /// <param name="height">The height of the shape.</param>
    /// <param name="createFromTemplate">
    /// When <c>true</c>, the shape is initialized from a default template.
    /// When <c>false</c>, the shape is created bare/empty.
    /// </param>
    /// <returns>The created auto shape.</returns>
    IAutoShape InsertAutoShape(int index, ShapeType shapeType, float x, float y, float width, float height, bool createFromTemplate);

    /// <summary>
    /// Creates a connector shape and appends it to the collection.
    /// The shape is initialized from a default template.
    /// </summary>
    /// <param name="shapeType">The type of connector to create.</param>
    /// <param name="x">The x-coordinate of the upper-left corner.</param>
    /// <param name="y">The y-coordinate of the upper-left corner.</param>
    /// <param name="width">The width of the connector.</param>
    /// <param name="height">The height of the connector.</param>
    /// <returns>The created connector.</returns>
    IConnector AddConnector(ShapeType shapeType, float x, float y, float width, float height);

    /// <summary>
    /// Creates a connector shape and appends it to the collection.
    /// </summary>
    /// <param name="shapeType">The type of connector to create.</param>
    /// <param name="x">The x-coordinate of the upper-left corner.</param>
    /// <param name="y">The y-coordinate of the upper-left corner.</param>
    /// <param name="width">The width of the connector.</param>
    /// <param name="height">The height of the connector.</param>
    /// <param name="createFromTemplate">
    /// When <c>true</c>, the connector is initialized from a default template.
    /// When <c>false</c>, the connector is created bare/empty.
    /// </param>
    /// <returns>The created connector.</returns>
    IConnector AddConnector(ShapeType shapeType, float x, float y, float width, float height, bool createFromTemplate);

    /// <summary>
    /// Creates a connector shape and inserts it at the specified position.
    /// The shape is initialized from a default template.
    /// </summary>
    /// <param name="index">The zero-based insertion index.</param>
    /// <param name="shapeType">The type of connector to create.</param>
    /// <param name="x">The x-coordinate of the upper-left corner.</param>
    /// <param name="y">The y-coordinate of the upper-left corner.</param>
    /// <param name="width">The width of the connector.</param>
    /// <param name="height">The height of the connector.</param>
    /// <returns>The created connector.</returns>
    IConnector InsertConnector(int index, ShapeType shapeType, float x, float y, float width, float height);

    /// <summary>
    /// Creates a connector shape and inserts it at the specified position.
    /// </summary>
    /// <param name="index">The zero-based insertion index.</param>
    /// <param name="shapeType">The type of connector to create.</param>
    /// <param name="x">The x-coordinate of the upper-left corner.</param>
    /// <param name="y">The y-coordinate of the upper-left corner.</param>
    /// <param name="width">The width of the connector.</param>
    /// <param name="height">The height of the connector.</param>
    /// <param name="createFromTemplate">
    /// When <c>true</c>, the connector is initialized from a default template.
    /// When <c>false</c>, the connector is created bare/empty.
    /// </param>
    /// <returns>The created connector.</returns>
    IConnector InsertConnector(int index, ShapeType shapeType, float x, float y, float width, float height, bool createFromTemplate);

    /// <summary>
    /// Creates a picture frame and appends it to the collection.
    /// </summary>
    /// <param name="shapeType">The type of picture frame to create.</param>
    /// <param name="x">The x-coordinate of the upper-left corner.</param>
    /// <param name="y">The y-coordinate of the upper-left corner.</param>
    /// <param name="width">The width of the picture frame.</param>
    /// <param name="height">The height of the picture frame.</param>
    /// <param name="image">The image to display in the frame.</param>
    /// <returns>The created picture frame.</returns>
    IPictureFrame AddPictureFrame(ShapeType shapeType, float x, float y, float width, float height, IPPImage image);

    /// <summary>
    /// Creates a picture frame and inserts it at the specified position.
    /// </summary>
    /// <param name="index">The zero-based insertion index.</param>
    /// <param name="shapeType">The type of picture frame to create.</param>
    /// <param name="x">The x-coordinate of the upper-left corner.</param>
    /// <param name="y">The y-coordinate of the upper-left corner.</param>
    /// <param name="width">The width of the picture frame.</param>
    /// <param name="height">The height of the picture frame.</param>
    /// <param name="image">The image to display in the frame.</param>
    /// <returns>The created picture frame.</returns>
    IPictureFrame InsertPictureFrame(int index, ShapeType shapeType, float x, float y, float width, float height, IPPImage image);

    /// <summary>
    /// Creates a table and appends it to the collection.
    /// </summary>
    /// <param name="x">The x-coordinate of the upper-left corner.</param>
    /// <param name="y">The y-coordinate of the upper-left corner.</param>
    /// <param name="columnWidths">The widths of each column.</param>
    /// <param name="rowHeights">The heights of each row.</param>
    /// <returns>The created table.</returns>
    ITable AddTable(float x, float y, double[] columnWidths, double[] rowHeights);

    /// <summary>
    /// Creates a table and inserts it at the specified position.
    /// </summary>
    /// <param name="index">The zero-based insertion index.</param>
    /// <param name="x">The x-coordinate of the upper-left corner.</param>
    /// <param name="y">The y-coordinate of the upper-left corner.</param>
    /// <param name="columnWidths">The widths of each column.</param>
    /// <param name="rowHeights">The heights of each row.</param>
    /// <returns>The created table.</returns>
    ITable InsertTable(int index, float x, float y, double[] columnWidths, double[] rowHeights);

    /// <summary>
    /// Returns the zero-based index of the specified shape, or -1 if not found.
    /// </summary>
    /// <param name="shape">The shape to locate.</param>
    /// <returns>The zero-based index, or -1 if the shape is not in the collection.</returns>
    int IndexOf(IShape shape);

    /// <summary>
    /// Removes the shape at the specified zero-based index.
    /// </summary>
    /// <param name="index">The zero-based index of the shape to remove.</param>
    void RemoveAt(int index);

    /// <summary>
    /// Removes the first occurrence of the specified shape from the collection.
    /// </summary>
    /// <param name="shape">The shape to remove.</param>
    void Remove(IShape shape);

    /// <summary>
    /// Removes all shapes from the collection.
    /// </summary>
    void Clear();

    /// <summary>
    /// Gets the number of shapes in the collection.
    /// </summary>
    int Count { get; }
}
