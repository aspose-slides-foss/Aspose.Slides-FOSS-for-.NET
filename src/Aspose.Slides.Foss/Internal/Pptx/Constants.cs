using System.Collections.Frozen;

namespace Aspose.Slides.Foss.Internal.Pptx;

/// <summary>
/// XML namespace constants and common element/attribute names for the PPTX format.
/// </summary>
internal static class Constants
{
    /// <summary>
    /// All XML namespace URIs used in PowerPoint Open XML documents, keyed by prefix.
    /// </summary>
    public static FrozenDictionary<string, string> Namespaces { get; } = new Dictionary<string, string>
    {
        // PresentationML namespace (main presentation elements)
        ["p"] = "http://schemas.openxmlformats.org/presentationml/2006/main",

        // DrawingML namespace (shapes, text, effects)
        ["a"] = "http://schemas.openxmlformats.org/drawingml/2006/main",

        // Relationships namespace (in XML content, e.g., r:id attributes)
        ["r"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",

        // Package relationships namespace (in .rels files)
        ["pr"] = "http://schemas.openxmlformats.org/package/2006/relationships",

        // Content types namespace
        ["ct"] = "http://schemas.openxmlformats.org/package/2006/content-types",

        // Core properties namespace (Dublin Core)
        ["cp"] = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
        ["dc"] = "http://purl.org/dc/elements/1.1/",
        ["dcterms"] = "http://purl.org/dc/terms/",

        // Extended properties namespace
        ["ep"] = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",

        // VML namespace (legacy vector markup)
        ["v"] = "urn:schemas-microsoft-com:vml",

        // Office namespace
        ["o"] = "urn:schemas-microsoft-com:office:office",

        // Chart namespace
        ["c"] = "http://schemas.openxmlformats.org/drawingml/2006/chart",

        // Diagram namespace
        ["dgm"] = "http://schemas.openxmlformats.org/drawingml/2006/diagram",

        // Picture namespace
        ["pic"] = "http://schemas.openxmlformats.org/drawingml/2006/picture",

        // Math namespace
        ["m"] = "http://schemas.openxmlformats.org/officeDocument/2006/math",

        // Microsoft Office extensions (2010+)
        ["p14"] = "http://schemas.microsoft.com/office/powerpoint/2010/main",
        ["p15"] = "http://schemas.microsoft.com/office/powerpoint/2012/main",
        ["a14"] = "http://schemas.microsoft.com/office/drawing/2010/main",

        // Markup Compatibility namespace
        ["mc"] = "http://schemas.openxmlformats.org/markup-compatibility/2006",
    }.ToFrozenDictionary();

    /// <summary>
    /// 1 point = 12,700 EMUs (914,400 EMU/inch / 72 points/inch).
    /// </summary>
    public const int EmuPerPoint = 12_700;

    /// <summary>
    /// OOXML stores rotation in 60,000ths of a degree.
    /// </summary>
    public const int RotationUnit = 60_000;
}

/// <summary>
/// Namespace helper providing formatted namespace strings for element lookups.
/// Use <c>Ns.P</c> for <c>{http://...}</c> style prefixes.
/// </summary>
internal static class Ns
{
    /// <summary>PresentationML namespace prefix.</summary>
    public static readonly string P = $"{{{Constants.Namespaces["p"]}}}";

    /// <summary>DrawingML namespace prefix.</summary>
    public static readonly string A = $"{{{Constants.Namespaces["a"]}}}";

    /// <summary>Document relationships namespace prefix.</summary>
    public static readonly string R = $"{{{Constants.Namespaces["r"]}}}";

    /// <summary>Package relationships namespace prefix.</summary>
    public static readonly string Pr = $"{{{Constants.Namespaces["pr"]}}}";

    /// <summary>Content types namespace prefix.</summary>
    public static readonly string Ct = $"{{{Constants.Namespaces["ct"]}}}";

    /// <summary>Charts namespace prefix.</summary>
    public static readonly string C = $"{{{Constants.Namespaces["c"]}}}";

    /// <summary>Pictures namespace prefix.</summary>
    public static readonly string Pic = $"{{{Constants.Namespaces["pic"]}}}";

    /// <summary>
    /// Gets a formatted namespace string by prefix.
    /// </summary>
    /// <param name="prefix">Namespace prefix (e.g., "p", "a", "r").</param>
    /// <returns>Formatted namespace string like <c>{http://...}</c>.</returns>
    /// <exception cref="KeyNotFoundException">If the prefix is not found.</exception>
    public static string Get(string prefix) => $"{{{Constants.Namespaces[prefix]}}}";
}

/// <summary>
/// Common PPTX element names with full namespace qualification.
/// </summary>
internal static class Elements
{
    // Presentation elements
    /// <summary>p:presentation element.</summary>
    public static readonly string Presentation = $"{Ns.P}presentation";
    /// <summary>p:sldIdLst element.</summary>
    public static readonly string SldIdLst = $"{Ns.P}sldIdLst";
    /// <summary>p:sldId element.</summary>
    public static readonly string SldId = $"{Ns.P}sldId";
    /// <summary>p:sldMasterIdLst element.</summary>
    public static readonly string SldMasterIdLst = $"{Ns.P}sldMasterIdLst";
    /// <summary>p:sldMasterId element.</summary>
    public static readonly string SldMasterId = $"{Ns.P}sldMasterId";
    /// <summary>p:sldSz element.</summary>
    public static readonly string SldSz = $"{Ns.P}sldSz";
    /// <summary>p:notesSz element.</summary>
    public static readonly string NotesSz = $"{Ns.P}notesSz";

    // Slide elements
    /// <summary>p:sld element.</summary>
    public static readonly string Sld = $"{Ns.P}sld";
    /// <summary>p:cSld element.</summary>
    public static readonly string CSld = $"{Ns.P}cSld";
    /// <summary>p:spTree element.</summary>
    public static readonly string SpTree = $"{Ns.P}spTree";

    // Shape elements
    /// <summary>p:sp element.</summary>
    public static readonly string Sp = $"{Ns.P}sp";
    /// <summary>p:nvSpPr element.</summary>
    public static readonly string NvSpPr = $"{Ns.P}nvSpPr";
    /// <summary>p:cNvPr element.</summary>
    public static readonly string CNvPr = $"{Ns.P}cNvPr";
    /// <summary>p:spPr element.</summary>
    public static readonly string SpPr = $"{Ns.P}spPr";
    /// <summary>p:txBody element.</summary>
    public static readonly string TxBody = $"{Ns.P}txBody";

    // DrawingML elements
    /// <summary>a:p (paragraph) element.</summary>
    public static readonly string AP = $"{Ns.A}p";
    /// <summary>a:r (run) element.</summary>
    public static readonly string AR = $"{Ns.A}r";
    /// <summary>a:t (text) element.</summary>
    public static readonly string AT = $"{Ns.A}t";
    /// <summary>a:xfrm (transform) element.</summary>
    public static readonly string AXfrm = $"{Ns.A}xfrm";
    /// <summary>a:off (offset) element.</summary>
    public static readonly string AOff = $"{Ns.A}off";
    /// <summary>a:ext (extents) element.</summary>
    public static readonly string AExt = $"{Ns.A}ext";

    // Fill elements
    /// <summary>a:noFill element.</summary>
    public static readonly string ANoFill = $"{Ns.A}noFill";
    /// <summary>a:solidFill element.</summary>
    public static readonly string ASolidFill = $"{Ns.A}solidFill";
    /// <summary>a:gradFill element.</summary>
    public static readonly string AGradFill = $"{Ns.A}gradFill";
    /// <summary>a:pattFill element.</summary>
    public static readonly string APattFill = $"{Ns.A}pattFill";
    /// <summary>a:blipFill element.</summary>
    public static readonly string ABlipFill = $"{Ns.A}blipFill";
    /// <summary>a:grpFill element.</summary>
    public static readonly string AGrpFill = $"{Ns.A}grpFill";

    // Color elements
    /// <summary>a:srgbClr element.</summary>
    public static readonly string ASrgbClr = $"{Ns.A}srgbClr";
    /// <summary>a:schemeClr element.</summary>
    public static readonly string ASchemeClr = $"{Ns.A}schemeClr";
    /// <summary>a:prstClr element.</summary>
    public static readonly string APrstClr = $"{Ns.A}prstClr";
    /// <summary>a:sysClr element.</summary>
    public static readonly string ASysClr = $"{Ns.A}sysClr";
    /// <summary>a:hlsClr element.</summary>
    public static readonly string AHlsClr = $"{Ns.A}hlsClr";
    /// <summary>a:scrgbClr element.</summary>
    public static readonly string AScrgbClr = $"{Ns.A}scrgbClr";

    // Gradient elements
    /// <summary>a:gsLst element.</summary>
    public static readonly string AGsLst = $"{Ns.A}gsLst";
    /// <summary>a:gs element.</summary>
    public static readonly string AGs = $"{Ns.A}gs";
    /// <summary>a:lin element.</summary>
    public static readonly string ALin = $"{Ns.A}lin";
    /// <summary>a:path element.</summary>
    public static readonly string APath = $"{Ns.A}path";
    /// <summary>a:tileRect element.</summary>
    public static readonly string ATileRect = $"{Ns.A}tileRect";

    // Pattern elements
    /// <summary>a:fgClr element.</summary>
    public static readonly string AFgClr = $"{Ns.A}fgClr";
    /// <summary>a:bgClr element.</summary>
    public static readonly string ABgClr = $"{Ns.A}bgClr";

    // Text elements
    /// <summary>a:bodyPr element.</summary>
    public static readonly string ABodyPr = $"{Ns.A}bodyPr";
    /// <summary>a:lstStyle element.</summary>
    public static readonly string ALstStyle = $"{Ns.A}lstStyle";
    /// <summary>a:rPr element.</summary>
    public static readonly string ARPr = $"{Ns.A}rPr";
    /// <summary>a:pPr element.</summary>
    public static readonly string APPr = $"{Ns.A}pPr";
    /// <summary>a:endParaRPr element.</summary>
    public static readonly string AEndParaRPr = $"{Ns.A}endParaRPr";
    /// <summary>a:highlight element.</summary>
    public static readonly string AHighlight = $"{Ns.A}highlight";
    /// <summary>a:uLnTx element.</summary>
    public static readonly string AULnTx = $"{Ns.A}uLnTx";
    /// <summary>a:uLn element.</summary>
    public static readonly string AULn = $"{Ns.A}uLn";
    /// <summary>a:uFillTx element.</summary>
    public static readonly string AUFillTx = $"{Ns.A}uFillTx";
    /// <summary>a:uFill element.</summary>
    public static readonly string AUFill = $"{Ns.A}uFill";
    /// <summary>a:defRPr element.</summary>
    public static readonly string ADefRPr = $"{Ns.A}defRPr";
    /// <summary>a:tabLst element.</summary>
    public static readonly string ATabLst = $"{Ns.A}tabLst";
    /// <summary>a:lnSpc element.</summary>
    public static readonly string ALnSpc = $"{Ns.A}lnSpc";
    /// <summary>a:spcBef element.</summary>
    public static readonly string ASpcBef = $"{Ns.A}spcBef";
    /// <summary>a:spcAft element.</summary>
    public static readonly string ASpcAft = $"{Ns.A}spcAft";
    /// <summary>a:spcPct element.</summary>
    public static readonly string ASpcPct = $"{Ns.A}spcPct";
    /// <summary>a:spcPts element.</summary>
    public static readonly string ASpcPts = $"{Ns.A}spcPts";
    /// <summary>a:buNone element.</summary>
    public static readonly string ABuNone = $"{Ns.A}buNone";
    /// <summary>a:buChar element.</summary>
    public static readonly string ABuChar = $"{Ns.A}buChar";
    /// <summary>a:buAutoNum element.</summary>
    public static readonly string ABuAutoNum = $"{Ns.A}buAutoNum";
    /// <summary>a:buFont element.</summary>
    public static readonly string ABuFont = $"{Ns.A}buFont";
    /// <summary>a:buSzPct element.</summary>
    public static readonly string ABuSzPct = $"{Ns.A}buSzPct";
    /// <summary>a:buSzPts element.</summary>
    public static readonly string ABuSzPts = $"{Ns.A}buSzPts";
    /// <summary>a:buClr element.</summary>
    public static readonly string ABuClr = $"{Ns.A}buClr";
    /// <summary>a:buClrTx element.</summary>
    public static readonly string ABuClrTx = $"{Ns.A}buClrTx";
    /// <summary>a:buFontTx element.</summary>
    public static readonly string ABuFontTx = $"{Ns.A}buFontTx";
    /// <summary>a:buSzTx element.</summary>
    public static readonly string ABuSzTx = $"{Ns.A}buSzTx";
    /// <summary>a:buBlip element.</summary>
    public static readonly string ABuBlip = $"{Ns.A}buBlip";
    /// <summary>a:latin element.</summary>
    public static readonly string ALatin = $"{Ns.A}latin";
    /// <summary>a:ea element.</summary>
    public static readonly string AEa = $"{Ns.A}ea";
    /// <summary>a:cs element.</summary>
    public static readonly string ACs = $"{Ns.A}cs";
    /// <summary>a:sym element.</summary>
    public static readonly string ASym = $"{Ns.A}sym";

    // Autofit elements
    /// <summary>a:noAutofit element.</summary>
    public static readonly string ANoAutofit = $"{Ns.A}noAutofit";
    /// <summary>a:spAutoFit element.</summary>
    public static readonly string ASpAutoFit = $"{Ns.A}spAutoFit";
    /// <summary>a:normAutofit element.</summary>
    public static readonly string ANormAutofit = $"{Ns.A}normAutofit";

    // Text warp element
    /// <summary>a:prstTxWarp element.</summary>
    public static readonly string APrstTxWarp = $"{Ns.A}prstTxWarp";

    // Line elements
    /// <summary>a:ln element.</summary>
    public static readonly string ALn = $"{Ns.A}ln";
    /// <summary>a:prstDash element.</summary>
    public static readonly string APrstDash = $"{Ns.A}prstDash";
    /// <summary>a:custDash element.</summary>
    public static readonly string ACustDash = $"{Ns.A}custDash";
    /// <summary>a:round element.</summary>
    public static readonly string ARound = $"{Ns.A}round";
    /// <summary>a:bevel element.</summary>
    public static readonly string ABevel = $"{Ns.A}bevel";
    /// <summary>a:miter element.</summary>
    public static readonly string AMiter = $"{Ns.A}miter";
    /// <summary>a:headEnd element.</summary>
    public static readonly string AHeadEnd = $"{Ns.A}headEnd";
    /// <summary>a:tailEnd element.</summary>
    public static readonly string ATailEnd = $"{Ns.A}tailEnd";

    // 3D elements
    /// <summary>a:scene3d element.</summary>
    public static readonly string AScene3D = $"{Ns.A}scene3d";
    /// <summary>a:sp3d element.</summary>
    public static readonly string ASp3D = $"{Ns.A}sp3d";
    /// <summary>a:camera element.</summary>
    public static readonly string ACamera = $"{Ns.A}camera";
    /// <summary>a:lightRig element.</summary>
    public static readonly string ALightRig = $"{Ns.A}lightRig";
    /// <summary>a:bevelT element.</summary>
    public static readonly string ABevelT = $"{Ns.A}bevelT";
    /// <summary>a:bevelB element.</summary>
    public static readonly string ABevelB = $"{Ns.A}bevelB";
    /// <summary>a:contourClr element.</summary>
    public static readonly string AContourClr = $"{Ns.A}contourClr";
    /// <summary>a:extrusionClr element.</summary>
    public static readonly string AExtrusionClr = $"{Ns.A}extrusionClr";
    /// <summary>a:rot element.</summary>
    public static readonly string ARot = $"{Ns.A}rot";
    /// <summary>a:effectLst element.</summary>
    public static readonly string AEffectLst = $"{Ns.A}effectLst";
    /// <summary>a:effectDag element.</summary>
    public static readonly string AEffectDag = $"{Ns.A}effectDag";
    /// <summary>a:extLst element.</summary>
    public static readonly string AExtLst = $"{Ns.A}extLst";

    // Effect elements
    /// <summary>a:blur element.</summary>
    public static readonly string ABlur = $"{Ns.A}blur";
    /// <summary>a:fillOverlay element.</summary>
    public static readonly string AFillOverlay = $"{Ns.A}fillOverlay";
    /// <summary>a:glow element.</summary>
    public static readonly string AGlow = $"{Ns.A}glow";
    /// <summary>a:innerShdw element.</summary>
    public static readonly string AInnerShdw = $"{Ns.A}innerShdw";
    /// <summary>a:outerShdw element.</summary>
    public static readonly string AOuterShdw = $"{Ns.A}outerShdw";
    /// <summary>a:prstShdw element.</summary>
    public static readonly string APrstShdw = $"{Ns.A}prstShdw";
    /// <summary>a:reflection element.</summary>
    public static readonly string AReflection = $"{Ns.A}reflection";
    /// <summary>a:softEdge element.</summary>
    public static readonly string ASoftEdge = $"{Ns.A}softEdge";

    // Table elements
    /// <summary>a:graphic element.</summary>
    public static readonly string AGraphic = $"{Ns.A}graphic";
    /// <summary>a:graphicData element.</summary>
    public static readonly string AGraphicData = $"{Ns.A}graphicData";
    /// <summary>a:tbl element.</summary>
    public static readonly string ATbl = $"{Ns.A}tbl";
    /// <summary>a:tblPr element.</summary>
    public static readonly string ATblPr = $"{Ns.A}tblPr";
    /// <summary>a:tblGrid element.</summary>
    public static readonly string ATblGrid = $"{Ns.A}tblGrid";
    /// <summary>a:gridCol element.</summary>
    public static readonly string AGridCol = $"{Ns.A}gridCol";
    /// <summary>a:tr element.</summary>
    public static readonly string ATr = $"{Ns.A}tr";
    /// <summary>a:tc element.</summary>
    public static readonly string ATc = $"{Ns.A}tc";
    /// <summary>a:tcPr element.</summary>
    public static readonly string ATcPr = $"{Ns.A}tcPr";
    /// <summary>a:tblStyle element.</summary>
    public static readonly string ATblStyle = $"{Ns.A}tblStyle";
    /// <summary>a:tableStyleId element.</summary>
    public static readonly string ATableStyleId = $"{Ns.A}tableStyleId";

    // Table border elements
    /// <summary>a:lnL element.</summary>
    public static readonly string ALnL = $"{Ns.A}lnL";
    /// <summary>a:lnR element.</summary>
    public static readonly string ALnR = $"{Ns.A}lnR";
    /// <summary>a:lnT element.</summary>
    public static readonly string ALnT = $"{Ns.A}lnT";
    /// <summary>a:lnB element.</summary>
    public static readonly string ALnB = $"{Ns.A}lnB";
    /// <summary>a:lnTlToBr element.</summary>
    public static readonly string ALnTlToBr = $"{Ns.A}lnTlToBr";
    /// <summary>a:lnBlToTr element.</summary>
    public static readonly string ALnBlToTr = $"{Ns.A}lnBlToTr";

    // GraphicFrame elements
    /// <summary>p:graphicFrame element.</summary>
    public static readonly string PGraphicFrame = $"{Ns.P}graphicFrame";
    /// <summary>p:nvGraphicFramePr element.</summary>
    public static readonly string PNvGraphicFramePr = $"{Ns.P}nvGraphicFramePr";
    /// <summary>p:cNvGraphicFramePr element.</summary>
    public static readonly string PCNvGraphicFramePr = $"{Ns.P}cNvGraphicFramePr";
    /// <summary>a:graphicFrameLocking element.</summary>
    public static readonly string AGraphicFrameLocking = $"{Ns.A}graphicFrameLocking";
    /// <summary>p:nvPr element.</summary>
    public static readonly string PNvPr = $"{Ns.P}nvPr";
    /// <summary>p:xfrm element.</summary>
    public static readonly string PXfrm = $"{Ns.P}xfrm";

    /// <summary>Table graphic data URI.</summary>
    public const string TableUri = "http://schemas.openxmlformats.org/drawingml/2006/table";

    // DrawingML txBody (for table cells)
    /// <summary>a:txBody element.</summary>
    public static readonly string ATxBody = $"{Ns.A}txBody";
}

/// <summary>
/// Common PPTX attribute names.
/// </summary>
internal static class Attributes
{
    // With namespace (e.g., r:id)
    /// <summary>r:id attribute.</summary>
    public static readonly string RId = $"{Ns.R}id";
    /// <summary>r:embed attribute.</summary>
    public static readonly string REmbed = $"{Ns.R}embed";

    // Without namespace
    /// <summary>id attribute.</summary>
    public const string Id = "id";
    /// <summary>name attribute.</summary>
    public const string Name = "name";
    /// <summary>cx attribute.</summary>
    public const string Cx = "cx";
    /// <summary>cy attribute.</summary>
    public const string Cy = "cy";
    /// <summary>x attribute.</summary>
    public const string X = "x";
    /// <summary>y attribute.</summary>
    public const string Y = "y";
}
