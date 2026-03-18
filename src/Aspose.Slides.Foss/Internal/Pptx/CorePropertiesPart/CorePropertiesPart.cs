using System.Globalization;
using System.Xml.Linq;

namespace Aspose.Slides.Foss.Internal.Pptx.CorePropertiesPart;

/// <summary>
/// Parses and serializes docProps/core.xml (Dublin Core metadata).
/// Handles: title, subject, creator (author), keywords, description (comments),
/// category, contentStatus, contentType, lastModifiedBy, revision, created, modified, lastPrinted.
/// </summary>
public sealed class CorePropertiesPart
{
    private static readonly XNamespace CpNs = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private static readonly XNamespace DcNs = "http://purl.org/dc/elements/1.1/";
    private static readonly XNamespace DcTermsNs = "http://purl.org/dc/terms/";
    private static readonly XNamespace XsiNs = "http://www.w3.org/2001/XMLSchema-instance";
    private static readonly XNamespace DcmiTypeNs = "http://purl.org/dc/dcmitype/";

    /// <summary>
    /// The part path within the OPC package.
    /// </summary>
    public const string PartName = "docProps/core.xml";

    private readonly Internal.OpcPackage _package;
    private XElement? _root;
    private bool _dirty;

    // String properties

    /// <summary>Gets or sets the document title.</summary>
    public string? Title { get; set; }

    /// <summary>Gets or sets the document subject.</summary>
    public string? Subject { get; set; }

    /// <summary>Gets or sets the document creator (author).</summary>
    public string? Creator { get; set; }

    /// <summary>Gets or sets the document keywords.</summary>
    public string? Keywords { get; set; }

    /// <summary>Gets or sets the document description (comments).</summary>
    public string? Description { get; set; }

    /// <summary>Gets or sets the document category.</summary>
    public string? Category { get; set; }

    /// <summary>Gets or sets the content status.</summary>
    public string? ContentStatus { get; set; }

    /// <summary>Gets or sets the content type.</summary>
    public string? ContentType { get; set; }

    /// <summary>Gets or sets the last modified by user.</summary>
    public string? LastModifiedBy { get; set; }

    /// <summary>Gets or sets the revision number.</summary>
    public string? Revision { get; set; }

    // Date properties

    /// <summary>Gets or sets the creation date.</summary>
    public DateTime? Created { get; set; }

    /// <summary>Gets or sets the last modified date.</summary>
    public DateTime? Modified { get; set; }

    /// <summary>Gets or sets the last printed date.</summary>
    public DateTime? LastPrinted { get; set; }

    /// <summary>
    /// Initializes the core properties part by parsing docProps/core.xml from the package.
    /// </summary>
    /// <param name="package">The OPC package containing the presentation.</param>
    internal CorePropertiesPart(Internal.OpcPackage package)
    {
        _package = package;
        Parse();
    }

    /// <summary>
    /// Parses the docProps/core.xml part from the package.
    /// </summary>
    public void Parse()
    {
        var data = _package.GetPart(PartName);
        if (data is null)
            return;

        using var stream = new MemoryStream(data);
        var doc = XDocument.Load(stream);
        _root = doc.Root;

        Title = GetText(DcNs + "title");
        Subject = GetText(DcNs + "subject");
        Creator = GetText(DcNs + "creator");
        Keywords = GetText(CpNs + "keywords");
        Description = GetText(DcNs + "description");
        Category = GetText(CpNs + "category");
        ContentStatus = GetText(CpNs + "contentStatus");
        ContentType = GetText(CpNs + "contentType");
        LastModifiedBy = GetText(CpNs + "lastModifiedBy");
        Revision = GetText(CpNs + "revision");

        Created = ParseW3cdtf(GetText(DcTermsNs + "created"));
        Modified = ParseW3cdtf(GetText(DcTermsNs + "modified"));
        LastPrinted = ParseW3cdtf(GetText(CpNs + "lastPrinted"));
    }

    /// <summary>
    /// Gets the text content of a child element by its fully qualified name.
    /// </summary>
    /// <param name="tag">The fully qualified element name.</param>
    /// <returns>The text content, or <c>null</c> if not found or empty.</returns>
    public string? GetText(XName tag)
    {
        if (_root is null)
            return null;
        var el = _root.Element(tag);
        if (el is not null && !string.IsNullOrEmpty(el.Value))
            return el.Value;
        return null;
    }

    /// <summary>
    /// Marks the part as dirty so it will be serialized on save.
    /// </summary>
    public void MarkDirty()
    {
        _dirty = true;
    }

    /// <summary>
    /// Serializes the core properties back to the OPC package.
    /// </summary>
    public void Save()
    {
        if (!_dirty && _root is not null)
            return;

        var root = new XElement(CpNs + "coreProperties",
            new XAttribute(XNamespace.Xmlns + "cp", CpNs),
            new XAttribute(XNamespace.Xmlns + "dc", DcNs),
            new XAttribute(XNamespace.Xmlns + "dcterms", DcTermsNs),
            new XAttribute(XNamespace.Xmlns + "dcmitype", DcmiTypeNs),
            new XAttribute(XNamespace.Xmlns + "xsi", XsiNs));

        SetDc(root, "title", Title);
        SetDc(root, "subject", Subject);
        SetDc(root, "creator", Creator);
        SetCp(root, "keywords", Keywords);
        SetDc(root, "description", Description);
        SetCp(root, "category", Category);
        SetCp(root, "contentStatus", ContentStatus);
        SetCp(root, "contentType", ContentType);
        SetCp(root, "lastModifiedBy", LastModifiedBy);
        SetCp(root, "revision", Revision);

        SetDctermsDate(root, "created", Created);
        SetDctermsDate(root, "modified", Modified);
        if (LastPrinted is not null)
            SetCp(root, "lastPrinted", FormatW3cdtf(LastPrinted));

        var doc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
        using var ms = new MemoryStream();
        doc.Save(ms);
        _package.SetPart(PartName, ms.ToArray());
        _dirty = false;
    }

    /// <summary>
    /// Sets a Dublin Core element on the root if the value is not null.
    /// </summary>
    /// <param name="root">The root element to add to.</param>
    /// <param name="localName">The local element name.</param>
    /// <param name="value">The text value, or <c>null</c> to skip.</param>
    public void SetDc(XElement root, string localName, string? value)
    {
        if (value is not null)
        {
            var el = new XElement(DcNs + localName) { Value = value };
            root.Add(el);
        }
    }

    /// <summary>
    /// Sets a core property element on the root if the value is not null.
    /// </summary>
    /// <param name="root">The root element to add to.</param>
    /// <param name="localName">The local element name.</param>
    /// <param name="value">The text value, or <c>null</c> to skip.</param>
    public void SetCp(XElement root, string localName, string? value)
    {
        if (value is not null)
        {
            var el = new XElement(CpNs + localName) { Value = value };
            root.Add(el);
        }
    }

    /// <summary>
    /// Sets a dcterms date element with xsi:type attribute on the root if the value is not null.
    /// </summary>
    /// <param name="root">The root element to add to.</param>
    /// <param name="localName">The local element name.</param>
    /// <param name="dt">The datetime value, or <c>null</c> to skip.</param>
    public void SetDctermsDate(XElement root, string localName, DateTime? dt)
    {
        if (dt is not null)
        {
            var el = new XElement(DcTermsNs + localName);
            el.SetAttributeValue(XsiNs + "type", "dcterms:W3CDTF");
            el.Value = FormatW3cdtf(dt)!;
            root.Add(el);
        }
    }

    /// <summary>
    /// Resets all properties to null.
    /// </summary>
    public void Clear()
    {
        Title = null;
        Subject = null;
        Creator = null;
        Keywords = null;
        Description = null;
        Category = null;
        ContentStatus = null;
        ContentType = null;
        LastModifiedBy = null;
        Revision = null;
        Created = null;
        Modified = null;
        LastPrinted = null;
        _dirty = true;
    }

    /// <summary>
    /// Parses a W3CDTF datetime string to a <see cref="DateTime"/> value.
    /// </summary>
    /// <param name="text">The W3CDTF string to parse.</param>
    /// <returns>The parsed <see cref="DateTime"/>, or <c>null</c> if the input is null, empty, or invalid.</returns>
    public static DateTime? ParseW3cdtf(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return null;

        text = text.Trim();
        try
        {
            if (text.EndsWith('Z'))
                text = text[..^1] + "+00:00";

            return DateTime.Parse(text, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);
        }
        catch (FormatException)
        {
            return null;
        }
    }

    /// <summary>
    /// Formats a <see cref="DateTime"/> as a W3CDTF string.
    /// </summary>
    /// <param name="dt">The datetime to format.</param>
    /// <returns>The W3CDTF string, or <c>null</c> if the input is null.</returns>
    public static string? FormatW3cdtf(DateTime? dt)
    {
        if (dt is null)
            return null;

        var value = dt.Value;
        if (value.Kind == DateTimeKind.Unspecified)
            value = DateTime.SpecifyKind(value, DateTimeKind.Utc);

        return value.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);
    }
}
