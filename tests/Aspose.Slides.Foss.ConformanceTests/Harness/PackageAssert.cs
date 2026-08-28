using System.Text;
using System.Xml.Linq;
using System.Xml.XPath;

namespace Aspose.Slides.Foss.ConformanceTests.Harness;

/// <summary>
/// Assertions made against the bytes of a produced package.
/// <para>
/// These are the rules a consumer applies. They are deliberately independent of the library: they
/// know only ECMA-376 and the Open Packaging Conventions.
/// </para>
/// </summary>
internal static class PackageAssert
{
    /// <summary>
    /// The content type ECMA-376 requires for each kind of part, keyed by the path pattern the part
    /// name has to match. A part whose name matches a key must resolve exactly this content type.
    /// </summary>
    private static readonly (string Pattern, string ContentType)[] RequiredContentTypes =
    [
        ("ppt/presentation.xml", "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"),
        ("ppt/slides/slide*.xml", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"),
        ("ppt/slideLayouts/slideLayout*.xml", "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"),
        ("ppt/slideMasters/slideMaster*.xml", "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"),
        ("ppt/notesSlides/notesSlide*.xml", "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"),
        ("ppt/notesMasters/notesMaster*.xml", "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"),
        ("ppt/handoutMasters/handoutMaster*.xml", "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml"),
        ("ppt/comments/*.xml", "application/vnd.openxmlformats-officedocument.presentationml.comments+xml"),
        ("ppt/commentAuthors.xml", "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"),
        ("ppt/theme/theme*.xml", "application/vnd.openxmlformats-officedocument.theme+xml"),
        ("ppt/media/*.png", "image/png"),
        ("ppt/media/*.jpeg", "image/jpeg"),
        ("ppt/media/*.jpg", "image/jpeg"),
        ("ppt/media/*.gif", "image/gif"),
        ("docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
        ("docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml"),
    ];

    /// <summary>
    /// Asserts that every <c>r:id</c>, <c>r:embed</c> and <c>r:link</c> in every XML part of the
    /// package is declared as a <c>Relationship Id</c> in that part's own <c>.rels</c>.
    /// <para>
    /// A dangling reference is the single most common way a package that looks right in a diff is
    /// refused by PowerPoint or read back with the content missing.
    /// </para>
    /// </summary>
    internal static void AllRelationshipReferencesResolve(PptxPackage package)
    {
        var failures = new List<string>();

        foreach (var partName in package.XmlPartNames)
        {
            var declared = package.Relationships(partName).Select(r => r.Id).ToHashSet(StringComparer.Ordinal);

            foreach (var element in package.Xml(partName).Descendants())
            {
                foreach (var attribute in element.Attributes())
                {
                    if (attribute.Name.Namespace != Ns.R)
                        continue;
                    if (attribute.Name.LocalName is not ("id" or "embed" or "link"))
                        continue;

                    if (!declared.Contains(attribute.Value))
                    {
                        failures.Add(
                            $"{partName}: <{element.Name.LocalName} r:{attribute.Name.LocalName}=\"{attribute.Value}\"> " +
                            $"is not declared in {PptxPackage.RelationshipsPartNameFor(partName)} " +
                            $"(declared: {(declared.Count == 0 ? "none" : string.Join(", ", declared))})");
                    }
                }
            }
        }

        Assert.True(failures.Count == 0,
            $"{failures.Count} relationship reference(s) do not resolve:{Environment.NewLine}" +
            string.Join(Environment.NewLine, failures));
    }

    /// <summary>
    /// Asserts that every part in the package resolves some content type, through an <c>Override</c>
    /// or a <c>Default</c>.
    /// </summary>
    internal static void EveryPartResolvesAContentType(PptxPackage package)
    {
        var failures = package.ContentPartNames
            .Where(name => package.ContentTypeOf(name) is null)
            .Select(name => $"{name}: no Override and no Default for its extension")
            .ToList();

        Assert.True(failures.Count == 0,
            $"{failures.Count} part(s) have no content type:{Environment.NewLine}" +
            string.Join(Environment.NewLine, failures));
    }

    /// <summary>
    /// Asserts that every part whose kind ECMA-376 defines resolves the content type ECMA-376 gives
    /// it. Per ISO/IEC 29500-2 §10.1.2 the content type is the part's identity: a slide part that
    /// falls through to <c>Default Extension="xml"</c> is an <c>application/xml</c> part, not a slide,
    /// and a strict reader will say so.
    /// </summary>
    internal static void EveryPartHasItsRequiredContentType(PptxPackage package)
    {
        var failures = new List<string>();

        foreach (var name in package.ContentPartNames)
        {
            var required = RequiredContentTypes
                .Where(r => Matches(name, r.Pattern))
                .Select(r => r.ContentType)
                .FirstOrDefault();
            if (required is null)
                continue;

            var actual = package.ContentTypeOf(name);
            if (!string.Equals(actual, required, StringComparison.Ordinal))
                failures.Add($"{name}: expected '{required}', resolved '{actual ?? "nothing"}'");
        }

        Assert.True(failures.Count == 0,
            $"{failures.Count} part(s) do not resolve their required content type:{Environment.NewLine}" +
            string.Join(Environment.NewLine, failures));
    }

    /// <summary>
    /// Asserts that no <c>Override</c> in <c>[Content_Types].xml</c> names a part the package does not
    /// contain — the signature of a part that was deleted without its bookkeeping.
    /// </summary>
    internal static void NoOverrideNamesAMissingPart(PptxPackage package)
    {
        var failures = package.OverriddenPartNames
            .Where(name => !package.Contains(name))
            .Select(name => $"Override names '/{name}', which is not in the package")
            .ToList();

        Assert.True(failures.Count == 0,
            $"{failures.Count} dangling content-type Override(s):{Environment.NewLine}" +
            string.Join(Environment.NewLine, failures));
    }

    /// <summary>
    /// Asserts that every internal relationship in the package targets a part that exists.
    /// </summary>
    internal static void EveryInternalRelationshipTargetExists(PptxPackage package)
    {
        var failures = new List<string>();

        foreach (var partName in package.PartNames.Where(PptxPackage.IsRelationshipsPart))
        {
            var owner = OwnerOfRelationshipsPart(partName);
            foreach (var relationship in package.Relationships(owner))
            {
                if (!string.Equals(relationship.TargetMode, "Internal", StringComparison.OrdinalIgnoreCase))
                    continue;

                var target = PptxPackage.ResolveTarget(owner, relationship.Target);
                if (!package.Contains(target))
                    failures.Add($"{partName}: {relationship.Id} targets '{relationship.Target}' ('{target}'), which is not in the package");
            }
        }

        Assert.True(failures.Count == 0,
            $"{failures.Count} relationship(s) target a missing part:{Environment.NewLine}" +
            string.Join(Environment.NewLine, failures));
    }

    /// <summary>
    /// Asserts that an XPath-addressed element exists in a part and carries the given attributes.
    /// Attribute names are written with their prefix — <c>"val"</c>, <c>"r:embed"</c>.
    /// </summary>
    internal static XElement ElementWithAttributes(
        PptxPackage package,
        string partName,
        string xpath,
        params (string Name, string Value)[] attributes)
    {
        var element = SingleElement(package, partName, xpath);

        foreach (var (name, expected) in attributes)
        {
            var actual = AttributeValue(element, name);
            Assert.True(expected == actual,
                $"{partName}: {xpath} — attribute '{name}' is {(actual is null ? "absent" : $"'{actual}'")}, expected '{expected}'." +
                $"{Environment.NewLine}Element as written: {element}");
        }

        return element;
    }

    /// <summary>Asserts that an XPath expression matches exactly one element, and returns it.</summary>
    internal static XElement SingleElement(PptxPackage package, string partName, string xpath)
    {
        var matches = Select(package, partName, xpath);
        Assert.True(matches.Count == 1,
            $"{partName}: expected exactly one element matching '{xpath}', found {matches.Count}." +
            $"{Environment.NewLine}Part as written:{Environment.NewLine}{package.Text(partName)}");
        return matches[0];
    }

    /// <summary>Returns every element an XPath expression matches in a part.</summary>
    internal static IReadOnlyList<XElement> Select(PptxPackage package, string partName, string xpath) =>
        package.Xml(partName).XPathSelectElements(xpath, Ns.Resolver()).ToList();

    /// <summary>
    /// Asserts that the children of <paramref name="parent"/> appear in the relative order the schema
    /// requires. Children not named in <paramref name="schemaOrder"/> are ignored; children that are
    /// named must appear in that order, because OOXML sequences are ordered and a reader silently
    /// discards what arrives out of place.
    /// </summary>
    internal static void ChildrenInSchemaOrder(XElement parent, params XName[] schemaOrder)
    {
        var rank = schemaOrder.Select((name, index) => (name, index))
            .ToDictionary(pair => pair.name, pair => pair.index);

        var observed = parent.Elements()
            .Where(child => rank.ContainsKey(child.Name))
            .Select(child => (child.Name, Rank: rank[child.Name]))
            .ToList();

        for (var i = 1; i < observed.Count; i++)
        {
            if (observed[i].Rank < observed[i - 1].Rank)
            {
                var written = string.Join(", ", parent.Elements().Select(e => e.Name.LocalName));
                var required = string.Join(", ", schemaOrder.Select(n => n.LocalName));
                Assert.Fail(
                    $"<{parent.Name.LocalName}> children are out of schema order." +
                    $"{Environment.NewLine}written:  {written}" +
                    $"{Environment.NewLine}required: {required}");
            }
        }
    }

    /// <summary>Reads a prefixed attribute value off an element, or <c>null</c> when it is absent.</summary>
    internal static string? AttributeValue(XElement element, string prefixedName)
    {
        var colon = prefixedName.IndexOf(':');
        if (colon < 0)
            return (string?)element.Attribute(prefixedName);

        var prefix = prefixedName[..colon];
        var local = prefixedName[(colon + 1)..];
        XNamespace ns = prefix switch
        {
            "r" => Ns.R,
            "a" => Ns.A,
            "p" => Ns.P,
            "p14" => Ns.P14,
            _ => XNamespace.None,
        };
        return (string?)element.Attribute(ns + local);
    }

    /// <summary>
    /// Asserts that no attribute anywhere in a part has a name outside the OOXML vocabulary in the
    /// sense that matters here: a name beginning with an underscore is an internal marker that leaked
    /// into a published file.
    /// </summary>
    internal static void NoInternalMarkerAttributes(PptxPackage package, string partName)
    {
        var leaked = package.Xml(partName).Descendants()
            .SelectMany(e => e.Attributes().Select(a => (Element: e, Attribute: a)))
            .Where(pair => pair.Attribute.Name.LocalName.StartsWith('_'))
            .Select(pair => $"<{pair.Element.Name.LocalName} {pair.Attribute.Name.LocalName}=\"{pair.Attribute.Value}\">")
            .ToList();

        Assert.True(leaked.Count == 0,
            $"{partName} carries {leaked.Count} internal marker attribute(s):{Environment.NewLine}" +
            string.Join(Environment.NewLine, leaked));
    }

    /// <summary>Describes a part for a failure message, indented.</summary>
    internal static string Describe(PptxPackage package, string partName)
    {
        var builder = new StringBuilder();
        builder.AppendLine($"--- {partName}");
        builder.Append(package.Text(partName));
        return builder.ToString();
    }

    private static string OwnerOfRelationshipsPart(string relationshipsPartName)
    {
        // "ppt/slides/_rels/slide1.xml.rels" -> "ppt/slides/slide1.xml"; "_rels/.rels" -> "".
        var withoutSuffix = relationshipsPartName[..^".rels".Length];
        var slash = withoutSuffix.LastIndexOf('/');
        var fileName = slash < 0 ? withoutSuffix : withoutSuffix[(slash + 1)..];
        var relsDirectory = slash < 0 ? string.Empty : withoutSuffix[..slash];
        var directory = relsDirectory.EndsWith("_rels", StringComparison.Ordinal)
            ? relsDirectory[..Math.Max(0, relsDirectory.Length - "_rels".Length)].TrimEnd('/')
            : relsDirectory;
        return directory.Length == 0 ? fileName : $"{directory}/{fileName}";
    }

    private static bool Matches(string partName, string pattern)
    {
        var star = pattern.IndexOf('*');
        if (star < 0)
            return string.Equals(partName, pattern, StringComparison.OrdinalIgnoreCase);

        var prefix = pattern[..star];
        var suffix = pattern[(star + 1)..];
        return partName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
            && partName.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)
            && partName.Length >= prefix.Length + suffix.Length;
    }
}
