using System.Xml.Linq;
using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests the Row public API: Height, MinimalHeight, RowFormat,
/// AsICellCollection, AsIBulkTextFormattable, SetTextFormat.
/// </summary>
public sealed class RowTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const float EmuPerPoint = 12700f;

    /// <summary>
    /// Creates a Row backed by XML with the given height in points.
    /// </summary>
    private static Row CreateRow(float heightPoints, List<Cell>? cells = null)
    {
        var emuValue = (long)Math.Round(heightPoints * EmuPerPoint);
        var tr = new XElement(ANs + "tr", new XAttribute("h", emuValue));
        var row = new Row();
        row.InitInternal(tr, 0, cells ?? [], slidePart: null, parentSlide: null);
        return row;
    }

    /// <summary>
    /// Creates a Row with no backing XML element.
    /// </summary>
    private static Row CreateRowWithoutXml()
    {
        return new Row();
    }

    // ---------------------------------------------------------------
    // ---------------------------------------------------------------

    /// <summary>
    /// Row heights match constructor arguments.
    /// </summary>
    [Theory]
    [InlineData(30)]
    [InlineData(50)]
    [InlineData(70)]
    public void Height_MatchesConstructorArgument(float expected)
    {
        var row = CreateRow(expected);

        row.Height.Should().Be(expected);
    }

    /// <summary>
    /// Multiple rows with different heights each report correct value.
    /// </summary>
    [Fact]
    public void Height_MultipleRowsPreserveDistinctHeights()
    {
        var row1 = CreateRow(30);
        var row2 = CreateRow(50);
        var row3 = CreateRow(70);

        row1.Height.Should().Be(30);
        row2.Height.Should().Be(50);
        row3.Height.Should().Be(70);
    }

    /// <summary>
    /// Height defaults to zero when no XML element is present.
    /// </summary>
    [Fact]
    public void Height_DefaultsToZeroWithoutXml()
    {
        var row = CreateRowWithoutXml();

        row.Height.Should().Be(0f);
    }

    // ---------------------------------------------------------------
    // ---------------------------------------------------------------

    /// <summary>
    /// MinimalHeight getter returns the same value as Height.
    /// </summary>
    [Theory]
    [InlineData(30)]
    [InlineData(50)]
    [InlineData(70)]
    public void MinimalHeight_MatchesConstructorArgument(float expected)
    {
        var row = CreateRow(expected);

        row.MinimalHeight.Should().Be(expected);
    }

    /// <summary>
    /// MinimalHeight can be updated and reflects the new value.
    /// </summary>
    [Fact]
    public void MinimalHeight_CanBeUpdated()
    {
        var row = CreateRow(30);

        row.MinimalHeight = 100;

        row.MinimalHeight.Should().Be(100);
    }

    /// <summary>
    /// MinimalHeight round-trips through XML correctly (EMU conversion).
    /// </summary>
    [Fact]
    public void MinimalHeight_RoundTripsThroughXml()
    {
        var row = CreateRow(50);
        row.MinimalHeight = 75;

        row.MinimalHeight.Should().Be(75);
        row.Height.Should().Be(75);
    }

    /// <summary>
    /// MinimalHeight defaults to zero without XML.
    /// </summary>
    [Fact]
    public void MinimalHeight_DefaultsToZeroWithoutXml()
    {
        var row = CreateRowWithoutXml();

        row.MinimalHeight.Should().Be(0f);
    }

    /// <summary>
    /// MinimalHeight setter is safe when no XML element is present.
    /// </summary>
    [Fact]
    public void MinimalHeight_SetterDoesNotThrowWithoutXml()
    {
        var row = CreateRowWithoutXml();

        var act = () => row.MinimalHeight = 100;

        act.Should().NotThrow();
    }

    // ---------------------------------------------------------------
    // ---------------------------------------------------------------

    /// <summary>
    /// RowFormat is accessible and not null.
    /// </summary>
    [Fact]
    public void RowFormat_IsNotNull()
    {
        var row = CreateRow(30);

        row.RowFormat.Should().NotBeNull();
    }

    /// <summary>
    /// RowFormat returns IRowFormat instance.
    /// </summary>
    [Fact]
    public void RowFormat_ImplementsIRowFormat()
    {
        var row = CreateRow(30);

        row.RowFormat.Should().BeAssignableTo<IRowFormat>();
    }

    // ---------------------------------------------------------------
    // ---------------------------------------------------------------

    /// <summary>
    /// AsICellCollection returns this row as ICellCollection.
    /// </summary>
    [Fact]
    public void AsICellCollection_IsNotNull()
    {
        var row = CreateRow(30);

        row.AsICellCollection.Should().NotBeNull();
    }

    /// <summary>
    /// AsICellCollection returns the same instance.
    /// </summary>
    [Fact]
    public void AsICellCollection_ReturnsSelf()
    {
        var row = CreateRow(30);

        row.AsICellCollection.Should().BeSameAs(row);
    }

    /// <summary>
    /// AsICellCollection is assignable to ICellCollection.
    /// </summary>
    [Fact]
    public void AsICellCollection_ImplementsICellCollection()
    {
        var row = CreateRow(30);

        row.AsICellCollection.Should().BeAssignableTo<ICellCollection>();
    }

    // ---------------------------------------------------------------
    // ---------------------------------------------------------------

    /// <summary>
    /// AsIBulkTextFormattable returns this row as IBulkTextFormattable.
    /// </summary>
    [Fact]
    public void AsIBulkTextFormattable_IsNotNull()
    {
        var row = CreateRow(30);

        row.AsIBulkTextFormattable.Should().NotBeNull();
    }

    /// <summary>
    /// AsIBulkTextFormattable returns the same instance.
    /// </summary>
    [Fact]
    public void AsIBulkTextFormattable_ReturnsSelf()
    {
        var row = CreateRow(30);

        row.AsIBulkTextFormattable.Should().BeSameAs(row);
    }

    /// <summary>
    /// AsIBulkTextFormattable is assignable to IBulkTextFormattable.
    /// </summary>
    [Fact]
    public void AsIBulkTextFormattable_ImplementsIBulkTextFormattable()
    {
        var row = CreateRow(30);

        row.AsIBulkTextFormattable.Should().BeAssignableTo<IBulkTextFormattable>();
    }

    // ---------------------------------------------------------------
    // ---------------------------------------------------------------

    /// <summary>
    /// SetTextFormat(IBasePortionFormat) throws on null argument.
    /// </summary>
    [Fact]
    public void SetTextFormat_BasePortionFormat_ThrowsOnNull()
    {
        var row = CreateRow(30);

        var act = () => row.SetTextFormat((IBasePortionFormat)null!);

        act.Should().Throw<ArgumentNullException>();
    }

    /// <summary>
    /// SetTextFormat(IParagraphFormat) throws on null argument.
    /// </summary>
    [Fact]
    public void SetTextFormat_ParagraphFormat_ThrowsOnNull()
    {
        var row = CreateRow(30);

        var act = () => row.SetTextFormat((IParagraphFormat)null!);

        act.Should().Throw<ArgumentNullException>();
    }

    /// <summary>
    /// SetTextFormat(ITextFrameFormat) throws on null argument.
    /// </summary>
    [Fact]
    public void SetTextFormat_TextFrameFormat_ThrowsOnNull()
    {
        var row = CreateRow(30);

        var act = () => row.SetTextFormat((ITextFrameFormat)null!);

        act.Should().Throw<ArgumentNullException>();
    }

    // ---------------------------------------------------------------
    // Row as CellCollection (inherited behavior)
    // ---------------------------------------------------------------

    /// <summary>
    /// Row with no cells has Count of zero.
    /// </summary>
    [Fact]
    public void Count_IsZeroWithNoCells()
    {
        var row = CreateRow(30);

        row.Count.Should().Be(0);
    }

    /// <summary>
    /// Row can be instantiated without errors.
    /// </summary>
    [Fact]
    public void Row_CanBeInstantiated()
    {
        var row = new Row();

        row.Should().NotBeNull();
    }

    /// <summary>
    /// Row implements IRow interface.
    /// </summary>
    [Fact]
    public void Row_ImplementsIRow()
    {
        var row = CreateRow(30);

        row.Should().BeAssignableTo<IRow>();
    }

    /// <summary>
    /// Row implements IBulkTextFormattable interface.
    /// </summary>
    [Fact]
    public void Row_ImplementsIBulkTextFormattable()
    {
        var row = CreateRow(30);

        row.Should().BeAssignableTo<IBulkTextFormattable>();
    }
}
