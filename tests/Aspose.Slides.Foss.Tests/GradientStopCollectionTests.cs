using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests the GradientStopCollection public API: Add, Insert, RemoveAt, Clear,
/// Count, indexer, AsICollection, AsIEnumerable, and IEnumerable support.
/// </summary>
public sealed class GradientStopCollectionTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// <summary>
    /// Creates a GradientStopCollection backed by a fresh &lt;a:gsLst&gt; element.
    /// </summary>
    private static GradientStopCollection CreateCollection()
    {
        var gsLst = new XElement(ANs + "gsLst");
        var collection = new GradientStopCollection();
        collection.InitInternal(gsLst, slidePart: null, parentSlide: null);
        return collection;
    }

    // --- Count ---

    [Fact]
    public void Count_EmptyCollection_IsZero()
    {
        var sut = CreateCollection();
        sut.Count.Should().Be(0);
    }

    [Fact]
    public void Count_AfterAddingTwoStops_IsTwo()
    {
        var sut = CreateCollection();
        sut.Add(0.0f, Color.Blue);
        sut.Add(1.0f, Color.Red);
        sut.Count.Should().Be(2);
    }

    // --- Add with RGB Color ---

    [Fact]
    public void Add_RgbColor_ReturnsGradientStop()
    {
        var sut = CreateCollection();
        var stop = sut.Add(0.5f, Color.Green);
        stop.Should().NotBeNull();
        stop.Should().BeAssignableTo<IGradientStop>();
    }

    [Fact]
    public void Add_RgbColor_IncreasesCount()
    {
        var sut = CreateCollection();
        sut.Add(0.0f, Color.Blue);
        sut.Count.Should().Be(1);
        sut.Add(1.0f, Color.Red);
        sut.Count.Should().Be(2);
    }

    // --- Add with PresetColor ---

    [Fact]
    public void Add_PresetColor_IncreasesCount()
    {
        var sut = CreateCollection();
        sut.Add(0.0f, PresetColor.Blue);
        sut.Count.Should().Be(1);
    }

    [Fact]
    public void Add_PresetColor_ReturnsGradientStop()
    {
        var sut = CreateCollection();
        var stop = sut.Add(0.25f, PresetColor.Red);
        stop.Should().NotBeNull();
        stop.Should().BeAssignableTo<IGradientStop>();
    }

    // --- Add with SchemeColor ---

    [Fact]
    public void Add_SchemeColor_IncreasesCount()
    {
        var sut = CreateCollection();
        sut.Add(0.0f, SchemeColor.Accent1);
        sut.Count.Should().Be(1);
    }

    [Fact]
    public void Add_SchemeColor_ReturnsGradientStop()
    {
        var sut = CreateCollection();
        var stop = sut.Add(0.75f, SchemeColor.Accent2);
        stop.Should().NotBeNull();
        stop.Should().BeAssignableTo<IGradientStop>();
    }

    // --- Indexer ---

    [Fact]
    public void Indexer_ReturnsStopAtGivenIndex()
    {
        var sut = CreateCollection();
        sut.Add(0.0f, Color.Blue);
        sut.Add(1.0f, Color.Red);

        var stop0 = sut[0];
        var stop1 = sut[1];
        stop0.Should().NotBeNull();
        stop1.Should().NotBeNull();
    }

    [Fact]
    public void Indexer_NegativeIndex_ThrowsIndexOutOfRange()
    {
        var sut = CreateCollection();
        sut.Add(0.0f, Color.Blue);

        var act = () => sut[-1];
        act.Should().Throw<IndexOutOfRangeException>();
    }

    [Fact]
    public void Indexer_IndexBeyondCount_ThrowsIndexOutOfRange()
    {
        var sut = CreateCollection();
        sut.Add(0.0f, Color.Blue);

        var act = () => sut[1];
        act.Should().Throw<IndexOutOfRangeException>();
    }

    [Fact]
    public void Indexer_EmptyCollection_ThrowsIndexOutOfRange()
    {
        var sut = CreateCollection();

        var act = () => sut[0];
        act.Should().Throw<IndexOutOfRangeException>();
    }

    // --- Insert ---

    [Fact]
    public void Insert_RgbColor_IncreasesCount()
    {
        var sut = CreateCollection();
        sut.Add(0.0f, Color.Blue);
        sut.Add(1.0f, Color.Red);

        sut.Insert(1, 0.5f, Color.Green);
        sut.Count.Should().Be(3);
    }

    [Fact]
    public void Insert_AtBeginning_ShiftsExistingStops()
    {
        var sut = CreateCollection();
        sut.Add(0.5f, Color.Blue);
        sut.Add(1.0f, Color.Red);

        sut.Insert(0, 0.0f, Color.Green);
        sut.Count.Should().Be(3);
    }

    [Fact]
    public void Insert_PresetColor_IncreasesCount()
    {
        var sut = CreateCollection();
        sut.Add(0.0f, Color.Blue);
        sut.Insert(0, 0.0f, PresetColor.Red);
        sut.Count.Should().Be(2);
    }

    [Fact]
    public void Insert_SchemeColor_IncreasesCount()
    {
        var sut = CreateCollection();
        sut.Add(0.0f, Color.Blue);
        sut.Insert(0, 0.0f, SchemeColor.Accent1);
        sut.Count.Should().Be(2);
    }

    [Fact]
    public void Insert_BeyondCount_AppendsToEnd()
    {
        var sut = CreateCollection();
        sut.Add(0.0f, Color.Blue);
        sut.Insert(10, 1.0f, Color.Red);
        sut.Count.Should().Be(2);
    }

    // --- RemoveAt ---

    [Fact]
    public void RemoveAt_ValidIndex_DecreasesCount()
    {
        var sut = CreateCollection();
        sut.Add(0.0f, Color.Blue);
        sut.Add(0.5f, Color.Green);
        sut.Add(1.0f, Color.Red);

        sut.RemoveAt(1);
        sut.Count.Should().Be(2);
    }

    [Fact]
    public void RemoveAt_FirstElement_DecreasesCount()
    {
        var sut = CreateCollection();
        sut.Add(0.0f, Color.Blue);
        sut.Add(1.0f, Color.Red);

        sut.RemoveAt(0);
        sut.Count.Should().Be(1);
    }

    [Fact]
    public void RemoveAt_LastElement_DecreasesCount()
    {
        var sut = CreateCollection();
        sut.Add(0.0f, Color.Blue);
        sut.Add(1.0f, Color.Red);

        sut.RemoveAt(1);
        sut.Count.Should().Be(1);
    }

    [Fact]
    public void RemoveAt_InvalidIndex_DoesNotChangeCount()
    {
        var sut = CreateCollection();
        sut.Add(0.0f, Color.Blue);

        sut.RemoveAt(5);
        sut.Count.Should().Be(1);
    }

    // --- Clear ---

    [Fact]
    public void Clear_RemovesAllStops()
    {
        var sut = CreateCollection();
        sut.Add(0.0f, Color.Blue);
        sut.Add(0.5f, Color.Green);
        sut.Add(1.0f, Color.Red);

        sut.Clear();
        sut.Count.Should().Be(0);
    }

    [Fact]
    public void Clear_EmptyCollection_RemainsEmpty()
    {
        var sut = CreateCollection();
        sut.Clear();
        sut.Count.Should().Be(0);
    }

    // --- AsICollection ---

    [Fact]
    public void AsICollection_ReturnsListWithCorrectCount()
    {
        var sut = CreateCollection();
        sut.Add(0.0f, Color.Blue);
        sut.Add(1.0f, Color.Red);

        var list = sut.AsICollection;
        list.Should().HaveCount(2);
    }

    [Fact]
    public void AsICollection_EmptyCollection_ReturnsEmptyList()
    {
        var sut = CreateCollection();
        sut.AsICollection.Should().BeEmpty();
    }

    // --- AsIEnumerable ---

    [Fact]
    public void AsIEnumerable_ReturnsCorrectCount()
    {
        var sut = CreateCollection();
        sut.Add(0.0f, Color.Blue);
        sut.Add(1.0f, Color.Red);

        sut.AsIEnumerable.Should().HaveCount(2);
    }

    // --- IEnumerable ---

    [Fact]
    public void GetEnumerator_IteratesAllStops()
    {
        var sut = CreateCollection();
        sut.Add(0.0f, Color.Blue);
        sut.Add(0.5f, Color.Green);
        sut.Add(1.0f, Color.Red);

        var count = 0;
        foreach (var stop in sut)
        {
            stop.Should().NotBeNull();
            count++;
        }
        count.Should().Be(3);
    }

    [Fact]
    public void GetEnumerator_EmptyCollection_YieldsNothing()
    {
        var sut = CreateCollection();
        var stops = sut.ToList();
        stops.Should().BeEmpty();
    }

    // --- Interface assignment ---

    [Fact]
    public void GradientStopCollection_ImplementsIGradientStopCollection()
    {
        var sut = CreateCollection();
        sut.Should().BeAssignableTo<IGradientStopCollection>();
    }

    [Fact]
    public void GradientStopCollection_ImplementsIEnumerable()
    {
        var sut = CreateCollection();
        sut.Should().BeAssignableTo<IEnumerable<IGradientStop>>();
    }

    // --- Add then access round-trip (from test_gradient_fill) ---

    [Fact]
    public void Add_BlueAndRed_StopsAreAccessibleByIndex()
    {
        var sut = CreateCollection();
        sut.Add(0.0f, Color.Blue);
        sut.Add(1.0f, Color.Red);

        sut.Count.Should().BeGreaterThanOrEqualTo(2);
        sut[0].Should().NotBeNull();
        sut[1].Should().NotBeNull();
    }

    // --- Multiple Add overloads in sequence ---

    [Fact]
    public void Add_MixedColorTypes_AllIncreaseCount()
    {
        var sut = CreateCollection();
        sut.Add(0.0f, Color.Blue);
        sut.Add(0.33f, PresetColor.Green);
        sut.Add(0.66f, SchemeColor.Accent1);
        sut.Add(1.0f, Color.Red);

        sut.Count.Should().Be(4);
    }

    // --- Clear then re-add ---

    [Fact]
    public void Clear_ThenReAdd_CountReflectsNewStops()
    {
        var sut = CreateCollection();
        sut.Add(0.0f, Color.Blue);
        sut.Add(1.0f, Color.Red);
        sut.Clear();
        sut.Count.Should().Be(0);

        sut.Add(0.5f, Color.Green);
        sut.Count.Should().Be(1);
    }
}
