using System.Xml.Linq;

namespace Aspose.Slides.Foss.Internal.Pptx.CommentAuthorsPart;

/// <summary>
/// Raw data for a comment author parsed from XML.
/// Wraps the underlying <c>&lt;p:cmAuthor&gt;</c> element and provides typed access to its attributes.
/// </summary>
public sealed class AuthorData
{
    internal readonly XElement Elem;

    /// <summary>
    /// Initializes a new <see cref="AuthorData"/> wrapping the given XML element.
    /// </summary>
    public AuthorData(XElement elem)
    {
        Elem = elem;
    }

    /// <summary>
    /// Gets the author ID.
    /// </summary>
    public int Id =>
        int.Parse(Elem.Attribute("id")?.Value ?? "0");

    /// <summary>
    /// Gets or sets the author name.
    /// </summary>
    public string Name
    {
        get => Elem.Attribute("name")?.Value ?? "";
        set => Elem.SetAttributeValue("name", value);
    }

    /// <summary>
    /// Gets or sets the author initials.
    /// </summary>
    public string Initials
    {
        get => Elem.Attribute("initials")?.Value ?? "";
        set => Elem.SetAttributeValue("initials", value);
    }

    /// <summary>
    /// Gets or sets the last comment index for this author.
    /// </summary>
    public int LastIdx
    {
        get => int.Parse(Elem.Attribute("lastIdx")?.Value ?? "0");
        set => Elem.SetAttributeValue("lastIdx", value);
    }

    /// <summary>
    /// Gets the color index for this author.
    /// </summary>
    public int ClrIdx =>
        int.Parse(Elem.Attribute("clrIdx")?.Value ?? "0");
}
