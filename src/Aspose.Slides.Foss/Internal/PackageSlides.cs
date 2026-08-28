using System.Xml.Linq;

namespace Aspose.Slides.Foss.Internal;

/// <summary>
/// Package-level operations on the set of slides a package contains.
/// </summary>
/// <remarks>
/// A slide exists in a package when four things agree: the part itself, the relationship from
/// <c>ppt/presentation.xml</c>, the <c>&lt;p:sldId&gt;</c> entry that names that relationship, and
/// the content-type <c>Override</c> for the part. Dropping any one of them on its own leaves a
/// package that a reader can still open and that is nonetheless wrong — an orphan part nothing can
/// reach, or a relationship pointing at a part that is gone. Removal therefore has to be one
/// operation, which is what this class is.
/// </remarks>
internal static class PackageSlides
{
    internal const string SlideRelType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";

    private const string NotesSlideRelType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";

    private const string CommentsRelType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace RNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private const string PresentationPartName = "ppt/presentation.xml";

    /// <summary>
    /// Lists the slide parts the presentation registers, in <c>&lt;p:sldIdLst&gt;</c> order.
    /// </summary>
    internal static List<string> ListSlideParts(PresentationPart presentationPart)
    {
        var parts = new List<string>();
        foreach (var (_, relId) in presentationPart.GetSlideEntries())
        {
            var rel = presentationPart.Rels.GetById(relId);
            if (rel is not null && !rel.IsExternal)
                parts.Add(ResolveFromPresentation(rel.Target));
        }
        return parts;
    }

    /// <summary>
    /// Lists the slide parts a written package registers, reading <c>ppt/presentation.xml</c> and
    /// its relationships out of the package rather than from a live presentation part.
    /// </summary>
    /// <remarks>
    /// This is the view a reader gets. It is also the only correct one once the parts have been
    /// written, because the relationships part is written by more than one owner during a save.
    /// </remarks>
    internal static List<string> ListSlideParts(OpcPackage package)
    {
        var parts = new List<string>();

        var presentationData = package.GetPart(PresentationPartName);
        var relsData = package.GetPart(OpcPaths.RelsPartNameFor(PresentationPartName));
        if (presentationData is null || relsData is null)
            return parts;

        var rels = new RelsManager();
        rels.Load(relsData);

        using var ms = new MemoryStream(presentationData);
        var sldIdLst = XDocument.Load(ms).Root?.Element(PNs + "sldIdLst");
        if (sldIdLst is null)
            return parts;

        foreach (var sldId in sldIdLst.Elements(PNs + "sldId"))
        {
            var relId = sldId.Attribute(RNs + "id")?.Value;
            if (relId is null)
                continue;

            var rel = rels.GetById(relId);
            if (rel is not null && !rel.IsExternal)
                parts.Add(ResolveFromPresentation(rel.Target));
        }

        return parts;
    }

    /// <summary>
    /// Removes a slide from a package: its <c>&lt;p:sldId&gt;</c> entry, the presentation
    /// relationship that entry names, the part, the part's relationships, its content-type
    /// <c>Override</c>, and any notes slide or comments part that only this slide reached.
    /// </summary>
    /// <param name="package">The package to modify.</param>
    /// <param name="presentationPart">The presentation part of that package.</param>
    /// <param name="slidePartName">The absolute part name of the slide to remove.</param>
    /// <remarks>
    /// The caller flushes <paramref name="presentationPart"/>; batching several removals before one
    /// flush is why this does not do it.
    /// </remarks>
    internal static void RemoveSlide(OpcPackage package, PresentationPart presentationPart,
        string slidePartName)
    {
        var dependents = DependentParts(package, slidePartName);

        // The p:sldId entry and the relationship it names go together: an entry whose r:id no
        // longer resolves is a dangling reference, and a relationship no entry names is an orphan.
        foreach (var rel in presentationPart.Rels.FindByType(SlideRelType).ToList())
        {
            if (rel.IsExternal ||
                !string.Equals(ResolveFromPresentation(rel.Target), slidePartName,
                    StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            RemoveSlideIdEntry(presentationPart, rel.Id);
            presentationPart.Rels.Remove(rel.Id);
        }

        DeletePart(package, slidePartName);

        // A notes slide or comments part belongs to exactly one slide, but check rather than assume:
        // deleting a part another slide still points at would trade one broken package for another.
        var stillReferenced = PartsReferencedBySlides(package, presentationPart);
        foreach (var dependent in dependents)
        {
            if (!stillReferenced.Contains(dependent))
                DeletePart(package, dependent);
        }
    }

    /// <summary>
    /// Deletes a part, the relationships it owns and its content-type declaration.
    /// </summary>
    internal static void DeletePart(OpcPackage package, string partName)
    {
        package.RemovePart(partName);
        package.RemovePart(OpcPaths.RelsPartNameFor(partName));
        OpcRegistration.RemoveContentTypeOverride(package, partName);
    }

    private static void RemoveSlideIdEntry(PresentationPart presentationPart, string relId)
    {
        var sldIdLst = presentationPart.Document.Root?.Element(PNs + "sldIdLst");
        if (sldIdLst is null)
            return;

        foreach (var sldId in sldIdLst.Elements(PNs + "sldId").ToList())
        {
            if (string.Equals(sldId.Attribute(RNs + "id")?.Value, relId, StringComparison.Ordinal))
                sldId.Remove();
        }
    }

    /// <summary>
    /// The notes slide and comments parts a slide reaches through its own relationships.
    /// </summary>
    /// <remarks>
    /// Media is deliberately not collected. Removing the last slide that used an image leaves
    /// <c>ppt/media/imageN.png</c> in the package: the file is larger than it needs to be, but it is
    /// valid, nothing dangles, and PowerPoint keeps orphaned media too. Collecting it would mean
    /// proving no other part — a layout, a master, a notes slide, an unknown part this library
    /// preserves verbatim — still references it, and getting that wrong deletes a picture the user
    /// can still see.
    /// </remarks>
    private static List<string> DependentParts(OpcPackage package, string slidePartName)
    {
        var result = new List<string>();
        var baseDir = OpcPaths.DirectoryOf(slidePartName);

        foreach (var rel in LoadRelationships(package, slidePartName))
        {
            if (rel.IsExternal)
                continue;
            if (rel.Type != NotesSlideRelType && rel.Type != CommentsRelType)
                continue;

            result.Add(OpcPaths.Resolve(baseDir, rel.Target));
        }

        return result;
    }

    private static HashSet<string> PartsReferencedBySlides(OpcPackage package,
        PresentationPart presentationPart)
    {
        var referenced = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var slidePartName in ListSlideParts(presentationPart))
        {
            var baseDir = OpcPaths.DirectoryOf(slidePartName);
            foreach (var rel in LoadRelationships(package, slidePartName))
            {
                if (!rel.IsExternal)
                    referenced.Add(OpcPaths.Resolve(baseDir, rel.Target));
            }
        }

        return referenced;
    }

    private static IEnumerable<Relationship> LoadRelationships(OpcPackage package, string partName)
    {
        var data = package.GetPart(OpcPaths.RelsPartNameFor(partName));
        if (data is null)
            return [];

        var rels = new RelsManager();
        rels.Load(data);
        return rels.All;
    }

    private static string ResolveFromPresentation(string target)
    {
        return OpcPaths.Resolve(OpcPaths.DirectoryOf(PresentationPartName), target);
    }
}
