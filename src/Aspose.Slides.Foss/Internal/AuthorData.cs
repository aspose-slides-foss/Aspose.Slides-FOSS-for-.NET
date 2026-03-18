using System.Xml.Linq;

namespace Aspose.Slides.Foss.Internal;

/// <summary>
/// Raw data for a comment author parsed from XML.
/// </summary>
internal sealed class AuthorData
{
    private readonly XElement _elem;

    internal AuthorData(XElement elem)
    {
        _elem = elem;
    }

    internal int Id =>
        int.Parse(_elem.Attribute("id")?.Value ?? "0");

    internal string Name
    {
        get => _elem.Attribute("name")?.Value ?? "";
        set => _elem.SetAttributeValue("name", value);
    }

    internal string Initials
    {
        get => _elem.Attribute("initials")?.Value ?? "";
        set => _elem.SetAttributeValue("initials", value);
    }
}
