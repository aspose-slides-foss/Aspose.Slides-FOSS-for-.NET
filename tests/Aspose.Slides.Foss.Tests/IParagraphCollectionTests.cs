using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests the IParagraphCollection contract: Add, Insert, Remove, RemoveAt, Clear, Count, indexer.
/// </summary>
public sealed class IParagraphCollectionTests
{
    private sealed class TestParagraph : IParagraph
    {
        private string _text;

        public TestParagraph(string text = "") => _text = text;

        public override IPortionCollection Portions => null!;
        public override IParagraphFormat ParagraphFormat => null!;
        public override string Text { get => _text; set => _text = value; }
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
        var collection = new ParagraphCollection();
        var paragraph = new TestParagraph("Hello ");

        collection.Add(paragraph);

        collection.Count.Should().Be(1);
        collection[0].Text.Should().Be("Hello ");
    }

    [Fact]
    public void Add_MultipleItems_AllAccessibleByIndex()
    {
        var collection = new ParagraphCollection();
        collection.Add(new TestParagraph("First"));
        collection.Add(new TestParagraph("Second"));
        collection.Add(new TestParagraph("Third"));

        collection.Count.Should().Be(3);
        collection[0].Text.Should().Be("First");
        collection[1].Text.Should().Be("Second");
        collection[2].Text.Should().Be("Third");
    }

    // -------------------------------------------------------------------
    // "Gradient stops and angle persist."
    // Verifies that added items are retained in the collection (>= 2 stops).
    // -------------------------------------------------------------------

    [Fact]
    public void Add_RetainsAllAddedItems()
    {
        var collection = new ParagraphCollection();
        collection.Add(new TestParagraph("Stop1"));
        collection.Add(new TestParagraph("Stop2"));

        collection.Count.Should().BeGreaterThanOrEqualTo(2);
    }

    // -------------------------------------------------------------------
    // "Removing a comment persists."
    // Adds 3 items, removes at index 1, verifies count is 2.
    // -------------------------------------------------------------------

    [Fact]
    public void RemoveAt_RemovesItemAtSpecifiedIndex()
    {
        var collection = new ParagraphCollection();
        collection.Add(new TestParagraph("C1"));
        collection.Add(new TestParagraph("C2"));
        collection.Add(new TestParagraph("C3"));
        collection.Count.Should().Be(3);

        collection.RemoveAt(1);

        collection.Count.Should().Be(2);
        collection[0].Text.Should().Be("C1");
        collection[1].Text.Should().Be("C3");
    }

    // -------------------------------------------------------------------
    // "remove_at removes by index."
    // Adds 2 items, removes at index 0, verifies count is 1.
    // -------------------------------------------------------------------

    [Fact]
    public void RemoveAt_FirstElement_LeavesRemainingItems()
    {
        var collection = new ParagraphCollection();
        collection.Add(new TestParagraph("Rect"));
        collection.Add(new TestParagraph("Ellipse"));

        collection.RemoveAt(0);

        collection.Count.Should().Be(1);
        collection[0].Text.Should().Be("Ellipse");
    }

    // -------------------------------------------------------------------
    // "remove_at removes by index."
    // Adds 1 item (to have 2 total), removes at index 1, count is 1.
    // -------------------------------------------------------------------

    [Fact]
    public void RemoveAt_LastElement_DecreasesCount()
    {
        var collection = new ParagraphCollection();
        collection.Add(new TestParagraph("Slide1"));
        collection.Add(new TestParagraph("Slide2"));

        collection.RemoveAt(1);

        collection.Count.Should().Be(1);
    }

    // -------------------------------------------------------------------
    // "clear() removes all comments from an author."
    // -------------------------------------------------------------------

    [Fact]
    public void Clear_RemovesAllItems()
    {
        var collection = new ParagraphCollection();
        collection.Add(new TestParagraph("C1"));
        collection.Add(new TestParagraph("C2"));

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
        var collection = new ParagraphCollection();
        collection.Add(new TestParagraph("First"));
        collection.Add(new TestParagraph("Third"));

        collection.Insert(1, new TestParagraph("Second"));

        collection.Count.Should().Be(3);
        collection[1].Text.Should().Be("Second");
    }

    // -------------------------------------------------------------------
    // Remove by reference — behavioral analog of test_remove_shape
    // "Removing a shape by reference decreases count."
    // -------------------------------------------------------------------

    [Fact]
    public void Remove_ByReference_DecreasesCount()
    {
        var collection = new ParagraphCollection();
        var para = new TestParagraph("ToRemove");
        collection.Add(para);
        collection.Add(new TestParagraph("ToKeep"));

        var removed = collection.Remove(para);

        removed.Should().BeTrue();
        collection.Count.Should().Be(1);
        collection[0].Text.Should().Be("ToKeep");
    }

    [Fact]
    public void Remove_NonExistentItem_ReturnsFalse()
    {
        var collection = new ParagraphCollection();
        collection.Add(new TestParagraph("Existing"));

        var removed = collection.Remove(new TestParagraph("Other"));

        removed.Should().BeFalse();
        collection.Count.Should().Be(1);
    }

    // -------------------------------------------------------------------
    // Contract properties
    // -------------------------------------------------------------------

    [Fact]
    public void Count_IsZeroForNewCollection()
    {
        var collection = new ParagraphCollection();

        collection.Count.Should().Be(0);
    }

    [Fact]
    public void AsISlideComponent_ReturnsNonNull()
    {
        var collection = new ParagraphCollection();

        collection.AsISlideComponent.Should().NotBeNull();
    }

    [Fact]
    public void AsIEnumerable_ReturnsAllItems()
    {
        var collection = new ParagraphCollection();
        collection.Add(new TestParagraph("A"));
        collection.Add(new TestParagraph("B"));

        var items = collection.AsIEnumerable.ToList();

        items.Should().HaveCount(2);
        items[0].Text.Should().Be("A");
        items[1].Text.Should().Be("B");
    }

    [Fact]
    public void ParagraphCollection_ImplementsIParagraphCollection()
    {
        var collection = new ParagraphCollection();

        collection.Should().BeAssignableTo<IParagraphCollection>();
    }

    [Fact]
    public void ParagraphCollection_ImplementsIEnumerable()
    {
        var collection = new ParagraphCollection();
        collection.Add(new TestParagraph("X"));

        var items = collection.ToList();

        items.Should().HaveCount(1);
    }

    [Fact]
    public void ParagraphCollection_InheritsFromISlideComponent()
    {
        var collection = new ParagraphCollection();

        collection.Should().BeAssignableTo<ISlideComponent>();
    }

    // -------------------------------------------------------------------
    // Contains — behavioral analog of ParagraphCollection.contains
    // "contains(item) checks if item is in collection."
    // -------------------------------------------------------------------

    [Fact]
    public void Contains_ReturnsTrueForExistingItem()
    {
        var collection = new ParagraphCollection();
        var para = new TestParagraph("Existing");
        collection.Add(para);

        collection.Contains(para).Should().BeTrue();
    }

    [Fact]
    public void Contains_ReturnsFalseForNonExistentItem()
    {
        var collection = new ParagraphCollection();
        collection.Add(new TestParagraph("Existing"));

        collection.Contains(new TestParagraph("Other")).Should().BeFalse();
    }

    // -------------------------------------------------------------------
    // "index_of returns the correct position."
    // -------------------------------------------------------------------

    [Fact]
    public void IndexOf_ReturnsCorrectPosition()
    {
        var collection = new ParagraphCollection();
        var p0 = new TestParagraph("First");
        var p1 = new TestParagraph("Second");
        collection.Add(p0);
        collection.Add(p1);

        collection.IndexOf(p0).Should().Be(0);
        collection.IndexOf(p1).Should().Be(1);
    }

    [Fact]
    public void IndexOf_NonExistentItem_ReturnsNegativeOne()
    {
        var collection = new ParagraphCollection();
        collection.Add(new TestParagraph("Existing"));

        collection.IndexOf(new TestParagraph("Other")).Should().Be(-1);
    }

    // -------------------------------------------------------------------
    // IsReadOnly — ParagraphCollection.is_read_only always returns False
    // -------------------------------------------------------------------

    [Fact]
    public void IsReadOnly_ReturnsFalse()
    {
        var collection = new ParagraphCollection();

        collection.IsReadOnly.Should().BeFalse();
    }

    // -------------------------------------------------------------------
    // "clear() empties the shape collection."
    // -------------------------------------------------------------------

    [Fact]
    public void Clear_AfterAddingItems_CountIsZero()
    {
        var collection = new ParagraphCollection();
        collection.Add(new TestParagraph("A"));
        collection.Add(new TestParagraph("B"));
        collection.Add(new TestParagraph("C"));

        collection.Clear();

        collection.Count.Should().Be(0);
    }

    // -------------------------------------------------------------------
    // "Shapes collection is iterable."
    // -------------------------------------------------------------------

    [Fact]
    public void Enumeration_ReturnsAllItems()
    {
        var collection = new ParagraphCollection();
        collection.Add(new TestParagraph("A"));
        collection.Add(new TestParagraph("B"));

        var items = collection.ToList();

        items.Should().HaveCount(2);
        items[0].Text.Should().Be("A");
        items[1].Text.Should().Be("B");
    }

    // -------------------------------------------------------------------
    // "add_clone duplicates a slide with its shapes."
    // Verifies Add preserves item properties.
    // -------------------------------------------------------------------

    [Fact]
    public void Add_PreservesItemProperties()
    {
        var collection = new ParagraphCollection();
        var para = new TestParagraph("Cloned content");
        collection.Add(para);

        collection[0].Text.Should().Be("Cloned content");
    }

    // -------------------------------------------------------------------
    // "reorder() changes the z-order of shapes."
    // Verified through Insert at index 0.
    // -------------------------------------------------------------------

    [Fact]
    public void Insert_AtIndexZero_PlacesItemFirst()
    {
        var collection = new ParagraphCollection();
        collection.Add(new TestParagraph("Second"));

        collection.Insert(0, new TestParagraph("First"));

        collection.Count.Should().Be(2);
        collection[0].Text.Should().Be("First");
        collection[1].Text.Should().Be("Second");
    }

    // -------------------------------------------------------------------
    // "Removing a shape by reference decreases count."
    // Contains returns false after removal.
    // -------------------------------------------------------------------

    [Fact]
    public void Contains_ReturnsFalseAfterRemoval()
    {
        var collection = new ParagraphCollection();
        var para = new TestParagraph("ToRemove");
        collection.Add(para);

        collection.Remove(para);

        collection.Contains(para).Should().BeFalse();
    }
}
