using System.Xml.Linq;
using Aspose.Slides.Foss.ConformanceTests.Harness;
using Aspose.Slides.Foss.Export;

namespace Aspose.Slides.Foss.ConformanceTests;

/// <summary>
/// The one-line effect enablers have to produce an element ECMA-376 §20.1.8 considers complete. An
/// effect element that is missing a required attribute or its mandatory colour child is not a subtle
/// formatting difference: PowerPoint refuses the whole file.
/// </summary>
public sealed class EffectConformanceTests : IDisposable
{
    /// <summary>The DrawingML colour choice group — an effect that needs a colour needs one of these.</summary>
    private static readonly string[] ColourChoiceElements =
        ["srgbClr", "schemeClr", "sysClr", "scrgbClr", "hslClr", "prstClr"];

    /// <summary>The DrawingML fill group.</summary>
    private static readonly string[] FillElements =
        ["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"];

    private readonly TestWorkspace _workspace = new();

    public void Dispose() => _workspace.Dispose();

    [Theory]
    [InlineData("Blur")]
    [InlineData("FillOverlay")]
    [InlineData("Glow")]
    [InlineData("InnerShadow")]
    [InlineData("OuterShadow")]
    [InlineData("PresetShadow")]
    [InlineData("Reflection")]
    [InlineData("SoftEdge")]
    public void EnablingAnEffectProducesASchemaValidPackage(string effect)
    {
        var path = WriteShapeWithEffect(effect, $"effect-{effect}.pptx");

        SchemaValidation.HasNoSchemaErrors(path);
    }

    [Theory]
    [InlineData("OuterShadow", "outerShdw")]
    [InlineData("InnerShadow", "innerShdw")]
    [InlineData("Glow", "glow")]
    [InlineData("PresetShadow", "prstShdw")]
    public void AnEnabledShadowOrGlowCarriesTheColourItsTypeRequires(string effect, string tag)
    {
        var path = WriteShapeWithEffect(effect, $"effect-colour-{effect}.pptx");

        using var package = PptxPackage.Open(path);
        var element = PackageAssert.SingleElement(package, "ppt/slides/slide1.xml", $"//a:effectLst/a:{tag}");

        var colour = element.Elements().FirstOrDefault(child => ColourChoiceElements.Contains(child.Name.LocalName));

        Assert.True(colour is not null,
            $"<a:{tag}> was written as '{element}' — it carries no colour child. " +
            $"ECMA-376 §20.1.8 requires exactly one of: {string.Join(", ", ColourChoiceElements)}.");
    }

    [Fact]
    public void AnEnabledPresetShadowCarriesItsPresetAttribute()
    {
        var path = WriteShapeWithEffect("PresetShadow", "effect-prstShdw-prst.pptx");

        using var package = PptxPackage.Open(path);
        var element = PackageAssert.SingleElement(package, "ppt/slides/slide1.xml", "//a:effectLst/a:prstShdw");

        Assert.True(element.Attribute("prst") is not null,
            $"<a:prstShdw> was written as '{element}'. 'prst' is a required attribute of CT_PresetShadowEffect.");
    }

    [Fact]
    public void AnEnabledSoftEdgeCarriesItsRadius()
    {
        var path = WriteShapeWithEffect("SoftEdge", "effect-softEdge-rad.pptx");

        using var package = PptxPackage.Open(path);
        var element = PackageAssert.SingleElement(package, "ppt/slides/slide1.xml", "//a:effectLst/a:softEdge");

        Assert.True(element.Attribute("rad") is not null,
            $"<a:softEdge> was written as '{element}'. 'rad' is a required attribute of CT_SoftEdgesEffect.");
    }

    [Fact]
    public void AnEnabledFillOverlayCarriesItsBlendModeAndAFill()
    {
        var path = WriteShapeWithEffect("FillOverlay", "effect-fillOverlay.pptx");

        using var package = PptxPackage.Open(path);
        var element = PackageAssert.SingleElement(package, "ppt/slides/slide1.xml", "//a:effectLst/a:fillOverlay");

        Assert.True(element.Attribute("blend") is not null,
            $"<a:fillOverlay> was written as '{element}'. 'blend' is a required attribute of CT_FillOverlayEffect.");

        var fill = element.Elements().FirstOrDefault(child => FillElements.Contains(child.Name.LocalName));
        Assert.True(fill is not null,
            $"<a:fillOverlay> was written as '{element}' — it carries no fill child. " +
            $"CT_FillOverlayEffect requires exactly one of: {string.Join(", ", FillElements)}.");
    }

    private string WriteShapeWithEffect(string effect, string fileName)
    {
        var path = _workspace.PathFor(fileName);

        using var presentation = new Presentation();
        var shape = presentation.Slides[0].Shapes!.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);
        Enable(shape.EffectFormat!, effect);
        presentation.Save(path, SaveFormat.Pptx);

        return path;
    }

    private static void Enable(IEffectFormat format, string effect)
    {
        switch (effect)
        {
            case "Blur": format.EnableBlurEffect(); break;
            case "FillOverlay": format.EnableFillOverlayEffect(); break;
            case "Glow": format.EnableGlowEffect(); break;
            case "InnerShadow": format.EnableInnerShadowEffect(); break;
            case "OuterShadow": format.EnableOuterShadowEffect(); break;
            case "PresetShadow": format.EnablePresetShadowEffect(); break;
            case "Reflection": format.EnableReflectionEffect(); break;
            case "SoftEdge": format.EnableSoftEdgeEffect(); break;
            default: throw new ArgumentOutOfRangeException(nameof(effect), effect, "Unknown effect.");
        }
    }
}
