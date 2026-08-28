using System.Text;
using System.Xml.Linq;

namespace Aspose.Slides.Foss.Internal;

/// <summary>
/// The notes master a package needs before it can carry notes slides.
/// </summary>
/// <remarks>
/// ECMA-376 Part 1 §13.3.5 gives a notes slide exactly one implicit relationship to a notes master,
/// and §13.3.4 registers that master in <c>&lt;p:notesMasterIdLst&gt;</c> on the presentation.
/// PowerPoint reads the notes text back without either, which is what makes the omission easy to
/// ship: nothing shows it until a consumer that checks the relationship graph refuses the package,
/// or a notes placeholder has nothing to inherit its formatting from.
/// </remarks>
internal static class NotesMasterPart
{
    internal const string PartName = "ppt/notesMasters/notesMaster1.xml";

    internal const string RelType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster";

    private const string ThemeRelType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

    private const string ThemeDirectory = "ppt/theme";

    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace RNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>
    /// Returns the notes master part name, creating the part and registering it if the package does
    /// not have one yet.
    /// </summary>
    /// <param name="package">The package to inspect and, if needed, extend.</param>
    /// <param name="presentationPart">
    /// The presentation part of that package, or <c>null</c> when it is not available — in which
    /// case the part is still created but cannot be registered, and the caller gets the part name
    /// so the notes slide can at least point at something that exists.
    /// </param>
    internal static string EnsureInPackage(OpcPackage package, PresentationPart? presentationPart)
    {
        var existing = FindExisting(package, presentationPart);
        if (existing is not null)
            return existing;

        package.SetPart(PartName, Encoding.UTF8.GetBytes(NotesMasterXml));
        OpcRegistration.AddContentTypeOverride(package, PartName, PartContentTypes.NotesMaster);

        RelateOwnTheme(package);

        Register(presentationPart);
        return PartName;
    }

    /// <summary>
    /// Gives the notes master a theme part of its own, copied from the theme the package already
    /// uses so the notes inherit the same fonts and colours.
    /// </summary>
    /// <remarks>
    /// A theme part belongs to exactly one master. Relating the notes master to the slide master's
    /// theme produces a package that validates clean against ECMA-376 and that every library reads
    /// back correctly, and that PowerPoint refuses to open — the whole file, not the notes. Decks
    /// PowerPoint writes carry <c>ppt/theme/theme1.xml</c> for the slide master and
    /// <c>ppt/theme/theme2.xml</c> for the notes master, which is what this reproduces.
    /// </remarks>
    private static void RelateOwnTheme(OpcPackage package)
    {
        var source = FindThemeToCopy(package);
        if (source is null)
            return;

        var themePartName = NextFreeThemePartName(package);
        package.SetPart(themePartName, source);
        OpcRegistration.AddContentTypeOverride(package, themePartName, PartContentTypes.Theme);

        var rels = new RelsManager { Package = package, OwnerPartName = PartName };
        rels.Add(ThemeRelType, OpcPaths.Relative(OpcPaths.DirectoryOf(PartName), themePartName));
        rels.Save();
    }

    /// <summary>
    /// Returns the bytes of a theme part already in the package, preferring the first one, or
    /// <c>null</c> when the package has no theme at all.
    /// </summary>
    private static byte[]? FindThemeToCopy(OpcPackage package)
    {
        var firstTheme = package.GetSortedPartNames()
            .FirstOrDefault(name =>
                name.StartsWith(ThemeDirectory + "/", StringComparison.OrdinalIgnoreCase) &&
                name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));

        return firstTheme is null ? null : package.GetPart(firstTheme);
    }

    /// <summary>
    /// Returns the lowest <c>ppt/theme/themeN.xml</c> name the package does not already use.
    /// </summary>
    private static string NextFreeThemePartName(OpcPackage package)
    {
        for (var index = 1; ; index++)
        {
            var candidate = $"{ThemeDirectory}/theme{index}.xml";
            if (package.GetPart(candidate) is null)
                return candidate;
        }
    }

    /// <summary>
    /// Finds the notes master the presentation already registers, if there is one.
    /// </summary>
    private static string? FindExisting(OpcPackage package, PresentationPart? presentationPart)
    {
        if (presentationPart is not null)
        {
            foreach (var rel in presentationPart.Rels.FindByType(RelType))
            {
                if (!rel.IsExternal)
                    return OpcPaths.Resolve("ppt", rel.Target);
            }
        }

        return package.GetPart(PartName) is not null ? PartName : null;
    }

    /// <summary>
    /// Adds the presentation relationship and the <c>&lt;p:notesMasterIdLst&gt;</c> entry naming it.
    /// </summary>
    private static void Register(PresentationPart? presentationPart)
    {
        if (presentationPart is null)
            return;

        var relId = presentationPart.Rels.Add(RelType,
            OpcPaths.Relative("ppt", PartName));

        var root = presentationPart.Document.Root;
        if (root is null)
            return;

        var idLst = root.Element(PNs + "notesMasterIdLst");
        if (idLst is null)
        {
            idLst = new XElement(PNs + "notesMasterIdLst");

            // CT_Presentation orders these: sldMasterIdLst, notesMasterIdLst, handoutMasterIdLst,
            // sldIdLst. Anything out of order is a schema error, not a formatting preference.
            var masterIdLst = root.Element(PNs + "sldMasterIdLst");
            if (masterIdLst is not null)
                masterIdLst.AddAfterSelf(idLst);
            else
                root.AddFirst(idLst);
        }

        idLst.RemoveNodes();
        idLst.Add(new XElement(PNs + "notesMasterId", new XAttribute(RNs + "id", relId)));

        presentationPart.Flush();
    }

    private const string NotesMasterXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <p:notesMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:cSld>
            <p:spTree>
              <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
              <p:grpSpPr/>
            </p:spTree>
          </p:cSld>
          <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"
                    accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6"
                    hlink="hlink" folHlink="folHlink"/>
          <p:notesStyle/>
        </p:notesMaster>
        """;
}
