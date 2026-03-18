using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests the IPortionCollection contract: Add, Insert, Remove, RemoveAt, Clear, Count,
/// IndexOf, Contains, indexer, AsIEnumerable.
/// </summary>
public sealed class IPortionCollectionTests
{
    private sealed class TestPortion : IPortion
    {
        public string Label { get; }

        public TestPortion(string label = "") => Label = label;

        public override IBasePortionFormat? PortionFormat => null;
        public override string Text { get => Label; set { } }
        public override ISlideComponent AsISlideComponent => this;
        public override IBaseSlide? Slide => null;
        public override IPresentationComponent AsIPresentationComponent => this;
        public override IPresentation? Presentation => null;
    }

    // -------------------------------------------------------------------
    // "Adding a Portion appends text."
    // Verifies Add increases Count and the item is accessible.
    // -------------------------------------------------------------------

    [Fact]
    public void Add_IncreasesCountAndItemIsAccessible()
    {
        var collection = new PortionCollection();
        var portion = new TestPortion("Hello ");

        collection.Add(portion);

        collection.Count.Should().Be(1);
        collection[0].Should().BeSameAs(portion);
    }

    [Fact]
    public void Add_MultipleItems_AllAccessibleByIndex()
    {
        var collection = new PortionCollection();
        var p1 = new TestPortion("First");
        var p2 = new TestPortion("Second");
        var p3 = new TestPortion("Third");
        collection.Add(p1);
        collection.Add(p2);
        collection.Add(p3);

        collection.Count.Should().Be(3);
        collection[0].Should().BeSameAs(p1);
        collection[1].Should().BeSameAs(p2);
        collection[2].Should().BeSameAs(p3);
    }

    // -------------------------------------------------------------------
    // "Gradient stops and angle persist."
    // Verifies that added items are retained in the collection (>= 2 stops).
    // -------------------------------------------------------------------

    [Fact]
    public void Add_RetainsAllAddedItems()
    {
        var collection = new PortionCollection();
        collection.Add(new TestPortion("Stop1"));
        collection.Add(new TestPortion("Stop2"));

        collection.Count.Should().BeGreaterThanOrEqualTo(2);
    }

    // -------------------------------------------------------------------
    // "Removing a comment persists."
    // Adds 3 items, removes at index 1, verifies count is 2.
    // -------------------------------------------------------------------

    [Fact]
    public void RemoveAt_RemovesItemAtSpecifiedIndex()
    {
        var collection = new PortionCollection();
        var p1 = new TestPortion("C1");
        var p2 = new TestPortion("C2");
        var p3 = new TestPortion("C3");
        collection.Add(p1);
        collection.Add(p2);
        collection.Add(p3);
        collection.Count.Should().Be(3);

        collection.RemoveAt(1);

        collection.Count.Should().Be(2);
        collection[0].Should().BeSameAs(p1);
        collection[1].Should().BeSameAs(p3);
    }

    // -------------------------------------------------------------------
    // "remove_at removes by index."
    // Adds 2 items, removes at index 0, verifies count is 1.
    // -------------------------------------------------------------------

    [Fact]
    public void RemoveAt_FirstElement_LeavesRemainingItems()
    {
        var collection = new PortionCollection();
        collection.Add(new TestPortion("Rect"));
        var ellipse = new TestPortion("Ellipse");
        collection.Add(ellipse);

        collection.RemoveAt(0);

        collection.Count.Should().Be(1);
        collection[0].Should().BeSameAs(ellipse);
    }

    // -------------------------------------------------------------------
    // "remove_at removes by index."
    // Adds 2 items, removes at index 1, count is 1.
    // -------------------------------------------------------------------

    [Fact]
    public void RemoveAt_LastElement_DecreasesCount()
    {
        var collection = new PortionCollection();
        collection.Add(new TestPortion("Slide1"));
        collection.Add(new TestPortion("Slide2"));

        collection.RemoveAt(1);

        collection.Count.Should().Be(1);
    }

    // -------------------------------------------------------------------
    // "index_of returns the correct position."
    // -------------------------------------------------------------------

    [Fact]
    public void IndexOf_ReturnsCorrectPosition()
    {
        var collection = new PortionCollection();
        var p0 = new TestPortion("First");
        var p1 = new TestPortion("Second");
        collection.Add(p0);
        collection.Add(p1);

        collection.IndexOf(p0).Should().Be(0);
        collection.IndexOf(p1).Should().Be(1);
    }

    [Fact]
    public void IndexOf_NonExistentItem_ReturnsNegativeOne()
    {
        var collection = new PortionCollection();
        collection.Add(new TestPortion("Existing"));

        collection.IndexOf(new TestPortion("Other")).Should().Be(-1);
    }

    // -------------------------------------------------------------------
    // Clear — behavioral analog of test_clear_comments
    // "clear() removes all comments from an author."
    // -------------------------------------------------------------------

    [Fact]
    public void Clear_RemovesAllItems()
    {
        var collection = new PortionCollection();
        collection.Add(new TestPortion("C1"));
        collection.Add(new TestPortion("C2"));

        collection.Clear();

        collection.Count.Should().Be(0);
    }

    // -------------------------------------------------------------------
    // Insert — behavioral analog of test_insert_comment
    // "insert_comment places at the correct index."
    // -------------------------------------------------------------------

    [Fact]
    public void Insert_PlacesItemAtCorrectIndex()
    {
        var collection = new PortionCollection();
        var first = new TestPortion("First");
        var third = new TestPortion("Third");
        var second = new TestPortion("Second");
        collection.Add(first);
        collection.Add(third);

        collection.Insert(1, second);

        collection.Count.Should().Be(3);
        collection[0].Should().BeSameAs(first);
        collection[1].Should().BeSameAs(second);
        collection[2].Should().BeSameAs(third);
    }

    // -------------------------------------------------------------------
    // Contains
    // -------------------------------------------------------------------

    [Fact]
    public void Contains_ReturnsTrueForExistingItem()
    {
        var collection = new PortionCollection();
        var portion = new TestPortion("Existing");
        collection.Add(portion);

        collection.Contains(portion).Should().BeTrue();
    }

    [Fact]
    public void Contains_ReturnsFalseForNonExistentItem()
    {
        var collection = new PortionCollection();
        collection.Add(new TestPortion("Existing"));

        collection.Contains(new TestPortion("Other")).Should().BeFalse();
    }

    // -------------------------------------------------------------------
    // Remove by reference — behavioral analog of test_remove_shape
    // "Removing a shape by reference decreases count."
    // -------------------------------------------------------------------

    [Fact]
    public void Remove_ByReference_DecreasesCount()
    {
        var collection = new PortionCollection();
        var toRemove = new TestPortion("ToRemove");
        var toKeep = new TestPortion("ToKeep");
        collection.Add(toRemove);
        collection.Add(toKeep);

        var removed = collection.Remove(toRemove);

        removed.Should().BeTrue();
        collection.Count.Should().Be(1);
        collection[0].Should().BeSameAs(toKeep);
    }

    [Fact]
    public void Remove_NonExistentItem_ReturnsFalse()
    {
        var collection = new PortionCollection();
        collection.Add(new TestPortion("Existing"));

        var removed = collection.Remove(new TestPortion("Other"));

        removed.Should().BeFalse();
        collection.Count.Should().Be(1);
    }

    // -------------------------------------------------------------------
    // Contract properties
    // -------------------------------------------------------------------

    [Fact]
    public void Count_IsZeroForNewCollection()
    {
        var collection = new PortionCollection();

        collection.Count.Should().Be(0);
    }

    [Fact]
    public void AsIEnumerable_ReturnsAllItems()
    {
        var collection = new PortionCollection();
        var a = new TestPortion("A");
        var b = new TestPortion("B");
        collection.Add(a);
        collection.Add(b);

        var items = collection.AsIEnumerable.ToList();

        items.Should().HaveCount(2);
        items[0].Should().BeSameAs(a);
        items[1].Should().BeSameAs(b);
    }

    [Fact]
    public void PortionCollection_ImplementsIPortionCollection()
    {
        var collection = new PortionCollection();

        collection.Should().BeAssignableTo<IPortionCollection>();
    }

    [Fact]
    public void PortionCollection_ImplementsIEnumerable()
    {
        var collection = new PortionCollection();
        collection.Add(new TestPortion("X"));

        var items = collection.ToList();

        items.Should().HaveCount(1);
    }

    [Fact]
    public void PortionCollection_InheritsFromISlideComponent()
    {
        var collection = new PortionCollection();

        collection.Should().BeAssignableTo<ISlideComponent>();
    }

    // -------------------------------------------------------------------
    // IsReadOnly — always returns false
    // -------------------------------------------------------------------

    [Fact]
    public void IsReadOnly_ReturnsFalse()
    {
        var collection = new PortionCollection();

        collection.IsReadOnly.Should().BeFalse();
    }
}
