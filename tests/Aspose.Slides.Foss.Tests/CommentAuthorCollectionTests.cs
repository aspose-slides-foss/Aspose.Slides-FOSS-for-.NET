using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Internal;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

public sealed class CommentAuthorCollectionTests
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";

    private static XElement MakeAuthorElem(int id, string name, string initials)
    {
        return new XElement(PNs + "cmAuthor",
            new XAttribute("id", id),
            new XAttribute("name", name),
            new XAttribute("initials", initials));
    }

    private static CommentAuthorCollection CreateCollection(XElement? authorsRoot = null)
    {
        authorsRoot ??= new XElement(PNs + "cmAuthorLst");
        var authorsPart = new CommentAuthorsPart(authorsRoot);
        var package = new OpcPackage();

        var collection = new CommentAuthorCollection();
        collection.InitInternal(authorsPart, package);
        return collection;
    }

    [Fact]
    public void AddAuthor_ReturnsAuthorWithCorrectNameAndInitials()
    {
        var collection = CreateCollection();

        var author = collection.AddAuthor("Alice", "A");

        author.Name.Should().Be("Alice");
        author.Initials.Should().Be("A");
    }

    [Fact]
    public void AddAuthor_IncreasesCount()
    {
        var collection = CreateCollection();

        collection.AddAuthor("Alice", "A");

        collection.Count.Should().Be(1);
    }

    [Fact]
    public void AddAuthor_MultipleAuthorsCoexist()
    {
        var collection = CreateCollection();

        collection.AddAuthor("Alice", "A");
        collection.AddAuthor("Bob", "B");

        collection.Count.Should().Be(2);
    }

    [Fact]
    public void Indexer_ReturnsCorrectAuthor()
    {
        var collection = CreateCollection();
        collection.AddAuthor("Alice", "A");
        collection.AddAuthor("Bob", "B");

        collection[0].Name.Should().Be("Alice");
        collection[1].Name.Should().Be("Bob");
    }

    [Fact]
    public void Remove_RemovesAuthorFromCollection()
    {
        var collection = CreateCollection();
        collection.AddAuthor("Alice", "A");
        collection.AddAuthor("Bob", "B");

        collection.Remove(collection[0]);

        collection.Count.Should().Be(1);
        collection[0].Name.Should().Be("Bob");
    }

    [Fact]
    public void RemoveAt_RemovesByIndex()
    {
        var collection = CreateCollection();
        collection.AddAuthor("Alice", "A");
        collection.AddAuthor("Bob", "B");

        collection.RemoveAt(0);

        collection.Count.Should().Be(1);
        collection[0].Name.Should().Be("Bob");
    }

    [Fact]
    public void Clear_RemovesAllAuthors()
    {
        var collection = CreateCollection();
        collection.AddAuthor("Alice", "A");
        collection.AddAuthor("Bob", "B");

        collection.Clear();

        collection.Count.Should().Be(0);
    }

    [Fact]
    public void ToArray_ReturnsAllAuthors()
    {
        var collection = CreateCollection();
        collection.AddAuthor("Alice", "A");
        collection.AddAuthor("Bob", "B");

        var array = collection.ToArray();

        array.Should().HaveCount(2);
        array[0].Name.Should().Be("Alice");
        array[1].Name.Should().Be("Bob");
    }

    [Fact]
    public void FindByName_ReturnsMatchingAuthors()
    {
        var collection = CreateCollection();
        collection.AddAuthor("Alice", "A");
        collection.AddAuthor("Bob", "B");
        collection.AddAuthor("Alice", "AL");

        var results = collection.FindByName("Alice");

        results.Should().HaveCount(2);
        results.Should().AllSatisfy(a => a.Name.Should().Be("Alice"));
    }

    [Fact]
    public void FindByName_ReturnsEmptyForNoMatch()
    {
        var collection = CreateCollection();
        collection.AddAuthor("Alice", "A");

        var results = collection.FindByName("Charlie");

        results.Should().BeEmpty();
    }

    [Fact]
    public void FindByNameAndInitials_ReturnsExactMatch()
    {
        var collection = CreateCollection();
        collection.AddAuthor("Alice", "A");
        collection.AddAuthor("Alice", "AL");
        collection.AddAuthor("Bob", "B");

        var results = collection.FindByNameAndInitials("Alice", "AL");

        results.Should().HaveCount(1);
        results[0].Initials.Should().Be("AL");
    }

    [Fact]
    public void FindByNameAndInitials_ReturnsEmptyForNoMatch()
    {
        var collection = CreateCollection();
        collection.AddAuthor("Alice", "A");

        var results = collection.FindByNameAndInitials("Alice", "X");

        results.Should().BeEmpty();
    }

    [Fact]
    public void AsICollection_ReturnsAllAuthorsAsList()
    {
        var collection = CreateCollection();
        collection.AddAuthor("Alice", "A");
        collection.AddAuthor("Bob", "B");

        IList<ICommentAuthor> list = collection.AsICollection;

        list.Should().HaveCount(2);
    }

    [Fact]
    public void AsIEnumerable_ReturnsAllAuthorsAsEnumerable()
    {
        var collection = CreateCollection();
        collection.AddAuthor("Alice", "A");
        collection.AddAuthor("Bob", "B");

        collection.AsIEnumerable.Count().Should().Be(2);
    }

    [Fact]
    public void RemoveAt_OutOfRange_DoesNothing()
    {
        var collection = CreateCollection();
        collection.AddAuthor("Alice", "A");

        collection.RemoveAt(5);

        collection.Count.Should().Be(1);
    }

    [Fact]
    public void Indexer_OutOfRange_ThrowsIndexOutOfRangeException()
    {
        var collection = CreateCollection();

        var act = () => collection[0];

        act.Should().Throw<IndexOutOfRangeException>();
    }
}
