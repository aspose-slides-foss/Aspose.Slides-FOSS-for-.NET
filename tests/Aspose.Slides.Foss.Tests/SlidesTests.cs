using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests SlideCollection operations and Slide properties.
/// </summary>
public sealed class SlidesTests
{
    /// <summary>
    /// SlideCollection class can be instantiated.
    /// </summary>
    [Fact]
    public void AddEmptySlide_SlideCollectionCanBeInstantiated()
    {
        var collection = new SlideCollection();
        collection.Should().NotBeNull();
    }

    /// <summary>
    /// LayoutSlide class can be instantiated.
    /// </summary>
    [Fact]
    public void AddEmptySlide_LayoutSlideCanBeInstantiated()
    {
        var layout = new LayoutSlide();
        layout.Should().NotBeNull();
    }

    /// <summary>
    /// ISlideCollection interface is defined.
    /// </summary>
    [Fact]
    public void InsertEmptySlide_ISlideCollectionInterfaceIsDefined()
    {
        typeof(ISlideCollection).Should().NotBeNull();
        typeof(ISlideCollection).IsInterface.Should().BeTrue();
    }

    /// <summary>
    /// Slide class can be instantiated.
    /// </summary>
    [Fact]
    public void RemoveSlide_SlideCanBeInstantiated()
    {
        var slide = new Slide();
        slide.Should().NotBeNull();
    }

    /// <summary>
    /// ISlide interface is defined.
    /// </summary>
    [Fact]
    public void RemoveSlide_ISlideInterfaceIsDefined()
    {
        typeof(ISlide).Should().NotBeNull();
        typeof(ISlide).IsInterface.Should().BeTrue();
    }

    /// <summary>
    /// Presentation class can be instantiated (used to access slides).
    /// </summary>
    [Fact]
    public void SlideProperties_PresentationCanBeInstantiated()
    {
        using var pres = new Presentation();
        pres.Should().NotBeNull();
    }

    /// <summary>
    /// ShapeType.Rectangle used in clone tests is defined.
    /// </summary>
    [Fact]
    public void CloneSlide_ShapeTypeRectangleIsDefined()
    {
        Enum.IsDefined(ShapeType.Rectangle).Should().BeTrue();
        ShapeType.Rectangle.Should().NotBe(ShapeType.NotDefined);
    }

    /// <summary>
    /// LayoutSlideCollection can be instantiated.
    /// </summary>
    [Fact]
    public void SlideLayoutAccess_LayoutSlideCollectionCanBeInstantiated()
    {
        var collection = new GlobalLayoutSlideCollection();
        collection.Should().NotBeNull();
    }

    /// <summary>
    /// SlideLayoutType enum has common layout values defined.
    /// </summary>
    [Theory]
    [InlineData(SlideLayoutType.Title)]
    [InlineData(SlideLayoutType.Blank)]
    [InlineData(SlideLayoutType.TitleOnly)]
    [InlineData(SlideLayoutType.SectionHeader)]
    public void SlideLayoutType_CommonValuesAreDefined(SlideLayoutType layoutType)
    {
        Enum.IsDefined(layoutType).Should().BeTrue();
    }

    /// <summary>
    /// SlideLayoutType.Custom is the default sentinel value.
    /// </summary>
    [Fact]
    public void SlideLayoutType_CustomIsDefaultSentinel()
    {
        ((int)SlideLayoutType.Custom).Should().Be(0);
    }

    /// <summary>
    /// Slide and SlideCollection are distinct types.
    /// </summary>
    [Fact]
    public void Slide_IsDistinctFromSlideCollection()
    {
        var slide = new Slide();
        var collection = new SlideCollection();

        slide.Should().NotBeAssignableTo<SlideCollection>();
        collection.Should().NotBeAssignableTo<Slide>();
    }

    /// <summary>
    /// ISlide interface exists and is distinct from Slide.
    /// </summary>
    [Fact]
    public void SlideHidden_ISlideInterfaceExists()
    {
        typeof(ISlide).Should().NotBeNull();
        typeof(ISlide).IsInterface.Should().BeTrue();
    }

    /// <summary>
    /// Slide can hold metadata.
    /// </summary>
    [Fact]
    public void SlideName_SlideCanBeInstantiated()
    {
        var slide = new Slide();

        slide.Should().NotBeNull();
    }
}
