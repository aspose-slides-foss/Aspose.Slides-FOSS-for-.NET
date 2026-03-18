using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests for IMasterSlideCollection operations.
/// </summary>
public sealed class IMasterSlideCollectionTests
{
    /// <summary>
    /// A new presentation has at least one master slide.
    /// </summary>
    [Fact]
    public void NewPresentation_HasMasterSlides()
    {
        using var pres = new Presentation();

        pres.Masters.Count.Should().BeGreaterThanOrEqualTo(1);
    }

    /// <summary>
    /// Master slide is accessible by index.
    /// </summary>
    [Fact]
    public void Indexer_ReturnsMasterSlide()
    {
        using var pres = new Presentation();

        var master = pres.Masters[0];

        master.Should().NotBeNull();
    }

    /// <summary>
    /// AsICollection returns the master slides as a list.
    /// </summary>
    [Fact]
    public void AsICollection_ReturnsMasterSlidesAsList()
    {
        using var pres = new Presentation();

        pres.Masters.AsICollection.Should().NotBeNull();
        pres.Masters.AsICollection.Count.Should().BeGreaterThanOrEqualTo(1);
    }

    /// <summary>
    /// AsIEnumerable returns the master slides as an enumerable.
    /// </summary>
    [Fact]
    public void AsIEnumerable_ReturnsMasterSlidesAsEnumerable()
    {
        using var pres = new Presentation();

        pres.Masters.AsIEnumerable.Should().NotBeNull();
        pres.Masters.AsIEnumerable.Count().Should().BeGreaterThanOrEqualTo(1);
    }

    /// <summary>
    /// AddClone duplicates a master slide and increases count.
    /// </summary>
    [Fact]
    public void AddClone_DuplicatesMasterSlide()
    {
        using var pres = new Presentation();

        var initialCount = pres.Masters.Count;
        var sourceMaster = pres.Masters[0];

        var cloned = pres.Masters.AddClone(sourceMaster);

        cloned.Should().NotBeNull();
        pres.Masters.Count.Should().Be(initialCount + 1);
    }

    /// <summary>
    /// AddClone preserves layout slides from the source master.
    /// Analogous to test_clone_slide checking shapes are preserved on clone.
    /// </summary>
    [Fact]
    public void AddClone_PreservesLayoutSlides()
    {
        using var pres = new Presentation();

        var sourceMaster = pres.Masters[0];
        var sourceLayoutCount = sourceMaster.LayoutSlides.Count;

        var cloned = pres.Masters.AddClone(sourceMaster);

        cloned.LayoutSlides.Count.Should().Be(sourceLayoutCount);
    }

    /// <summary>
    /// Cloned master slide is a distinct object from the source.
    /// </summary>
    [Fact]
    public void AddClone_ReturnsDifferentInstance()
    {
        using var pres = new Presentation();

        var sourceMaster = pres.Masters[0];
        var cloned = pres.Masters.AddClone(sourceMaster);

        cloned.Should().NotBeSameAs(sourceMaster);
    }
}
