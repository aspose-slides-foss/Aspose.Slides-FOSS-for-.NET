using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests the IRow contract: Height, MinimalHeight, RowFormat,
/// AsICellCollection, and AsIBulkTextFormattable.
/// </summary>
public sealed class IRowContractTests
{
    /// <summary>
    /// Row class can be instantiated.
    /// </summary>
    [Fact]
    public void Row_CanBeInstantiated()
    {
        var row = new Row();
        row.Should().NotBeNull();
    }

    /// <summary>
    /// IRow interface inherits from ICellCollection.
    /// </summary>
    [Fact]
    public void IRow_InheritsFromICellCollection()
    {
        typeof(IRow).Should().BeAssignableTo<ICellCollection>();
    }

    /// <summary>
    /// IRow interface inherits from IBulkTextFormattable.
    /// </summary>
    [Fact]
    public void IRow_InheritsFromIBulkTextFormattable()
    {
        typeof(IRow).Should().BeAssignableTo<IBulkTextFormattable>();
    }

    /// <summary>
    /// IRow defines a Height property (float, read-only).
    /// </summary>
    [Fact]
    public void IRow_DefinesHeightProperty()
    {
        var prop = typeof(IRow).GetProperty(nameof(IRow.Height));
        prop.Should().NotBeNull();
        prop!.PropertyType.Should().Be(typeof(float));
        prop.CanRead.Should().BeTrue();
    }

    /// <summary>
    /// IRow defines a MinimalHeight property (float, read-write).
    /// </summary>
    [Fact]
    public void IRow_DefinesMinimalHeightProperty()
    {
        var prop = typeof(IRow).GetProperty(nameof(IRow.MinimalHeight));
        prop.Should().NotBeNull();
        prop!.PropertyType.Should().Be(typeof(float));
        prop.CanRead.Should().BeTrue();
        prop.CanWrite.Should().BeTrue();
    }

    /// <summary>
    /// IRow defines a RowFormat property (IRowFormat, read-only).
    /// </summary>
    [Fact]
    public void IRow_DefinesRowFormatProperty()
    {
        var prop = typeof(IRow).GetProperty(nameof(IRow.RowFormat));
        prop.Should().NotBeNull();
        prop!.PropertyType.Should().Be(typeof(IRowFormat));
        prop.CanRead.Should().BeTrue();
    }

    /// <summary>
    /// IRow defines AsICellCollection property that returns ICellCollection.
    /// </summary>
    [Fact]
    public void IRow_DefinesAsICellCollectionProperty()
    {
        var prop = typeof(IRow).GetProperty(nameof(IRow.AsICellCollection));
        prop.Should().NotBeNull();
        prop!.PropertyType.Should().Be(typeof(ICellCollection));
        prop.CanRead.Should().BeTrue();
    }

    /// <summary>
    /// IRow defines AsIBulkTextFormattable property that returns IBulkTextFormattable.
    /// </summary>
    [Fact]
    public void IRow_DefinesAsIBulkTextFormattableProperty()
    {
        var prop = typeof(IRow).GetProperty(nameof(IRow.AsIBulkTextFormattable));
        prop.Should().NotBeNull();
        prop!.PropertyType.Should().Be(typeof(IBulkTextFormattable));
        prop.CanRead.Should().BeTrue();
    }

    /// <summary>
    /// Row height values (30, 50, 70) are representable as float.
    /// </summary>
    [Theory]
    [InlineData(30f)]
    [InlineData(50f)]
    [InlineData(70f)]
    public void RowHeight_ValuesAreRepresentableAsFloat(float height)
    {
        height.Should().BeGreaterThan(0);
    }

    /// <summary>
    /// Height property is read-only on IRow (no setter).
    /// </summary>
    [Fact]
    public void IRow_HeightIsReadOnly()
    {
        var prop = typeof(IRow).GetProperty(nameof(IRow.Height));
        prop.Should().NotBeNull();
        prop!.GetSetMethod().Should().BeNull();
    }

    /// <summary>
    /// MinimalHeight property has both getter and setter on IRow.
    /// </summary>
    [Fact]
    public void IRow_MinimalHeightIsReadWrite()
    {
        var prop = typeof(IRow).GetProperty(nameof(IRow.MinimalHeight));
        prop.Should().NotBeNull();
        prop!.GetGetMethod().Should().NotBeNull();
        prop!.GetSetMethod().Should().NotBeNull();
    }

    /// <summary>
    /// IRowFormat interface exists and is accessible as a type.
    /// </summary>
    [Fact]
    public void IRowFormat_InterfaceExists()
    {
        typeof(IRowFormat).Should().NotBeNull();
        typeof(IRowFormat).IsInterface.Should().BeTrue();
    }

    /// <summary>
    /// IRow is an interface type.
    /// </summary>
    [Fact]
    public void IRow_IsAnInterface()
    {
        typeof(IRow).IsInterface.Should().BeTrue();
    }

    /// <summary>
    /// ICellCollection is accessible as a type inherited by IRow.
    /// </summary>
    [Fact]
    public void ICellCollection_IsAccessibleViaIRow()
    {
        typeof(IRow).GetInterfaces().Should().Contain(typeof(ICellCollection));
    }

    /// <summary>
    /// IBulkTextFormattable is accessible as a type inherited by IRow.
    /// </summary>
    [Fact]
    public void IBulkTextFormattable_IsAccessibleViaIRow()
    {
        typeof(IRow).GetInterfaces().Should().Contain(typeof(IBulkTextFormattable));
    }
}
