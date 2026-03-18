using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Internal;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

public sealed class CommentAuthorTests
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";

    private static XElement MakeAuthorElem(int id, string name, string initials)
    {
        return new XElement(PNs + "cmAuthor",
            new XAttribute("id", id),
            new XAttribute("name", name),
            new XAttribute("initials", initials));
    }

    private static CommentAuthor CreateAuthor(
        int id, string name, string initials,
        XElement? authorsRoot = null,
        OpcPackage? package = null,
        Presentation? presentation = null)
    {
        authorsRoot ??= new XElement(PNs + "cmAuthorLst", MakeAuthorElem(id, name, initials));
        var authorElem = authorsRoot.Elements(PNs + "cmAuthor")
            .First(e => int.Parse(e.Attribute("id")!.Value) == id);
        var authorsPart = new CommentAuthorsPart(authorsRoot);
        package ??= new OpcPackage();

        var author = new CommentAuthor();
        author.InitInternal(new AuthorData(authorElem), authorsPart, package, presentation);
        return author;
    }

    [Fact]
    public void Name_ReturnsAuthorName()
    {
        var author = CreateAuthor(0, "Alice", "A");

        author.Name.Should().Be("Alice");
    }

    [Fact]
    public void Initials_ReturnsAuthorInitials()
    {
        var author = CreateAuthor(0, "Alice", "A");

        author.Initials.Should().Be("A");
    }

    [Fact]
    public void Name_CanBeModified()
    {
        var author = CreateAuthor(0, "Alice", "A");

        author.Name = "Alicia";

        author.Name.Should().Be("Alicia");
    }

    [Fact]
    public void Initials_CanBeModified()
    {
        var author = CreateAuthor(0, "Alice", "A");

        author.Initials = "AL";

        author.Initials.Should().Be("AL");
    }

    [Fact]
    public void Comments_ReturnsNonNullCollection()
    {
        var author = CreateAuthor(0, "Alice", "A");

        author.Comments.Should().NotBeNull();
    }

    [Fact]
    public void Comments_ReturnsSameInstanceOnRepeatedAccess()
    {
        var author = CreateAuthor(0, "Alice", "A");

        var first = author.Comments;
        var second = author.Comments;

        first.Should().BeSameAs(second);
    }

    [Fact]
    public void MultipleAuthors_CoexistInSameXml()
    {
        var authorsRoot = new XElement(PNs + "cmAuthorLst",
            MakeAuthorElem(0, "Alice", "A"),
            MakeAuthorElem(1, "Bob", "B"));

        var alice = CreateAuthor(0, "Alice", "A", authorsRoot);
        var bob = CreateAuthor(1, "Bob", "B", authorsRoot);

        alice.Name.Should().Be("Alice");
        bob.Name.Should().Be("Bob");
        authorsRoot.Elements(PNs + "cmAuthor").Should().HaveCount(2);
    }

    [Fact]
    public void Remove_RemovesAuthorFromXml()
    {
        var authorsRoot = new XElement(PNs + "cmAuthorLst",
            MakeAuthorElem(0, "Alice", "A"),
            MakeAuthorElem(1, "Bob", "B"));

        var alice = CreateAuthor(0, "Alice", "A", authorsRoot);

        authorsRoot.Elements(PNs + "cmAuthor").Should().HaveCount(2);

        alice.Remove();

        authorsRoot.Elements(PNs + "cmAuthor").Should().HaveCount(1);
        authorsRoot.Elements(PNs + "cmAuthor").First().Attribute("name")!.Value.Should().Be("Bob");
    }

    [Fact]
    public void Remove_LeavesOtherAuthorsIntact()
    {
        var authorsRoot = new XElement(PNs + "cmAuthorLst",
            MakeAuthorElem(0, "Alice", "A"),
            MakeAuthorElem(1, "Bob", "B"));

        var alice = CreateAuthor(0, "Alice", "A", authorsRoot);
        var bob = CreateAuthor(1, "Bob", "B", authorsRoot);

        alice.Remove();

        bob.Name.Should().Be("Bob");
        bob.Initials.Should().Be("B");
    }

    [Fact]
    public void Clear_CanBeCalledOnCommentsCollection()
    {
        var author = CreateAuthor(0, "Alice", "A");

        var act = () => author.Comments.Clear();

        act.Should().NotThrow();
    }
}
