using System.Xml.Linq;
using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests the Column public API: Width, ColumnFormat, AsICellCollection,
/// AsIBulkTextFormattable, SetTextFormat.
/// </summary>
public sealed class ColumnTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const float EmuPerPoint = 12700f;

    /// <summary>
    /// Creates a Column backed by XML with the given width in points.
    /// </summary>
    private static Column CreateColumn(float widthPoints, List<Cell>? cells = null)
    {
        var emuValue = (long)Math.Round(widthPoints * EmuPerPoint);
        var gridCol = new XElement(ANs + "gridCol", new XAttribute("w", emuValue));
        var column = new Column();
        column.InitInternal(gridCol, 0, cells ?? new List<Cell>(), slidePart: null, parentSlide: null);
        return column;
    }

    /// <summary>
    /// Creates a Column with no backing XML element (null).
    /// </summary>
    private static Column CreateColumnWithoutXml()
    {
        var column = new Column();
        return column;
    }

    // ---------------------------------------------------------------
    // Width property
    // ---------------------------------------------------------------

    /// <summary>
    /// Column widths match constructor arguments.
    /// </summary>
    [Theory]
    [InlineData(100)]
    [InlineData(200)]
    [InlineData(300)]
    public void Width_MatchesConstructorArgument(float expected)
    {
        var col = CreateColumn(expected);

        col.Width.Should().Be(expected);
    }

    /// <summary>
    /// Multiple columns with different widths each report correct value.
    /// </summary>
    [Fact]
    public void Width_MultipleColumnsPreserveDistinctWidths()
    {
        var col1 = CreateColumn(100);
        var col2 = CreateColumn(200);
        var col3 = CreateColumn(300);

        col1.Width.Should().Be(100);
        col2.Width.Should().Be(200);
        col3.Width.Should().Be(300);
    }

    /// <summary>
    /// Width can be updated in-place.
    /// </summary>
    [Fact]
    public void Width_CanBeUpdated()
    {
        var col = CreateColumn(100);

        col.Width = 250;

        col.Width.Should().Be(250);
    }

    /// <summary>
    /// Width defaults to zero when no XML element is present.
    /// </summary>
    [Fact]
    public void Width_DefaultsToZeroWithoutXml()
    {
        var col = CreateColumnWithoutXml();

        col.Width.Should().Be(0f);
    }

    /// <summary>
    /// Width setter is safe when no XML element is present.
    /// </summary>
    [Fact]
    public void Width_SetterDoesNotThrowWithoutXml()
    {
        var col = CreateColumnWithoutXml();

        var act = () => col.Width = 100;

        act.Should().NotThrow();
    }

    /// <summary>
    /// Width round-trips through XML correctly (EMU conversion).
    /// </summary>
    [Fact]
    public void Width_RoundTripsThroughXml()
    {
        var col = CreateColumn(150);
        col.Width = 275;

        col.Width.Should().Be(275);
    }

    // ---------------------------------------------------------------
    // ColumnFormat property
    // ---------------------------------------------------------------

    /// <summary>
    /// ColumnFormat is accessible and not null.
    /// </summary>
    [Fact]
    public void ColumnFormat_IsNotNull()
    {
        var col = CreateColumn(100);

        col.ColumnFormat.Should().NotBeNull();
    }

    /// <summary>
    /// ColumnFormat returns IColumnFormat instance.
    /// </summary>
    [Fact]
    public void ColumnFormat_ImplementsIColumnFormat()
    {
        var col = CreateColumn(100);

        col.ColumnFormat.Should().BeAssignableTo<IColumnFormat>();
    }

    // ---------------------------------------------------------------
    // AsICellCollection property
    // ---------------------------------------------------------------

    /// <summary>
    /// AsICellCollection returns this column as ICellCollection.
    /// </summary>
    [Fact]
    public void AsICellCollection_IsNotNull()
    {
        var col = CreateColumn(100);

        col.AsICellCollection.Should().NotBeNull();
    }

    /// <summary>
    /// AsICellCollection returns the same instance.
    /// </summary>
    [Fact]
    public void AsICellCollection_ReturnsSelf()
    {
        var col = CreateColumn(100);

        col.AsICellCollection.Should().BeSameAs(col);
    }

    /// <summary>
    /// AsICellCollection is assignable to ICellCollection.
    /// </summary>
    [Fact]
    public void AsICellCollection_ImplementsICellCollection()
    {
        var col = CreateColumn(100);

        col.AsICellCollection.Should().BeAssignableTo<ICellCollection>();
    }

    // ---------------------------------------------------------------
    // AsIBulkTextFormattable property
    // ---------------------------------------------------------------

    /// <summary>
    /// AsIBulkTextFormattable returns this column as IBulkTextFormattable.
    /// </summary>
    [Fact]
    public void AsIBulkTextFormattable_IsNotNull()
    {
        var col = CreateColumn(100);

        col.AsIBulkTextFormattable.Should().NotBeNull();
    }

    /// <summary>
    /// AsIBulkTextFormattable returns the same instance.
    /// </summary>
    [Fact]
    public void AsIBulkTextFormattable_ReturnsSelf()
    {
        var col = CreateColumn(100);

        col.AsIBulkTextFormattable.Should().BeSameAs(col);
    }

    /// <summary>
    /// AsIBulkTextFormattable is assignable to IBulkTextFormattable.
    /// </summary>
    [Fact]
    public void AsIBulkTextFormattable_ImplementsIBulkTextFormattable()
    {
        var col = CreateColumn(100);

        col.AsIBulkTextFormattable.Should().BeAssignableTo<IBulkTextFormattable>();
    }

    // ---------------------------------------------------------------
    // SetTextFormat methods
    // ---------------------------------------------------------------

    /// <summary>
    /// SetTextFormat(IBasePortionFormat) throws on null argument.
    /// </summary>
    [Fact]
    public void SetTextFormat_BasePortionFormat_ThrowsOnNull()
    {
        var col = CreateColumn(100);

        var act = () => col.SetTextFormat((IBasePortionFormat)null!);

        act.Should().Throw<ArgumentNullException>();
    }

    /// <summary>
    /// SetTextFormat(IParagraphFormat) throws on null argument.
    /// </summary>
    [Fact]
    public void SetTextFormat_ParagraphFormat_ThrowsOnNull()
    {
        var col = CreateColumn(100);

        var act = () => col.SetTextFormat((IParagraphFormat)null!);

        act.Should().Throw<ArgumentNullException>();
    }

    /// <summary>
    /// SetTextFormat(ITextFrameFormat) throws on null argument.
    /// </summary>
    [Fact]
    public void SetTextFormat_TextFrameFormat_ThrowsOnNull()
    {
        var col = CreateColumn(100);

        var act = () => col.SetTextFormat((ITextFrameFormat)null!);

        act.Should().Throw<ArgumentNullException>();
    }

    // ---------------------------------------------------------------
    // Column as CellCollection (inherited behavior)
    // ---------------------------------------------------------------

    /// <summary>
    /// Column with no cells has Count of zero.
    /// </summary>
    [Fact]
    public void Count_IsZeroWithNoCells()
    {
        var col = CreateColumn(100);

        col.Count.Should().Be(0);
    }

    /// <summary>
    /// Column can be instantiated without errors.
    /// </summary>
    [Fact]
    public void Column_CanBeInstantiated()
    {
        var col = new Column();

        col.Should().NotBeNull();
    }

    /// <summary>
    /// Column implements IColumn interface.
    /// </summary>
    [Fact]
    public void Column_ImplementsIColumn()
    {
        var col = CreateColumn(100);

        col.Should().BeAssignableTo<IColumn>();
    }

    /// <summary>
    /// Column implements IBulkTextFormattable interface.
    /// </summary>
    [Fact]
    public void Column_ImplementsIBulkTextFormattable()
    {
        var col = CreateColumn(100);

        col.Should().BeAssignableTo<IBulkTextFormattable>();
    }
}
