using Aspose.Slides.Foss.ConformanceTests.Harness;
using Aspose.Slides.Foss.Export;

namespace Aspose.Slides.Foss.ConformanceTests;

/// <summary>
/// Package-level behaviour that is already correct, pinned so that it stays correct.
/// <para>
/// Unlike the rest of this project, every test here passes on the code as it stands. They exist
/// because the repairs the other files describe touch the same writers, and a fix that quietly
/// regresses one of these would otherwise be invisible.
/// </para>
/// </summary>
public sealed class PackageRegressionTests : IDisposable
{
    private readonly TestWorkspace _workspace = new();

    public void Dispose() => _workspace.Dispose();

    [Fact]
    public void AnEmptyDeckIsAWellFormedPackage()
    {
        var path = _workspace.PathFor("empty.pptx");

        using (var presentation = new Presentation())
        {
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);
        PackageAssert.EveryPartResolvesAContentType(package);
        PackageAssert.EveryPartHasItsRequiredContentType(package);
        PackageAssert.NoOverrideNamesAMissingPart(package);
        PackageAssert.EveryInternalRelationshipTargetExists(package);
        PackageAssert.AllRelationshipReferencesResolve(package);
        SchemaValidation.HasNoSchemaErrors(path);
    }

    [Fact]
    public void AGraphicFrameLocksItsContentsThroughTheSchemaElementName()
    {
        var path = _workspace.PathFor("graphic-frame-locks.pptx");

        using (var presentation = new Presentation())
        {
            presentation.Slides[0].Shapes!.AddTable(50, 50, [100, 100], [30, 30]);
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);

        PackageAssert.SingleElement(package, "ppt/slides/slide1.xml", "//p:cNvGraphicFramePr/a:graphicFrameLocks");
        Assert.Empty(PackageAssert.Select(package, "ppt/slides/slide1.xml", "//a:graphicFrameLocking"));
    }

    [Fact]
    public void AConnectorWritesTheConnectionSitesItWasGiven()
    {
        var path = _workspace.PathFor("connector.pptx");

        using (var presentation = new Presentation())
        {
            var shapes = presentation.Slides[0].Shapes!;
            var start = shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 50);
            var end = shapes.AddAutoShape(ShapeType.Rectangle, 200, 10, 100, 50);
            var connector = shapes.AddConnector(ShapeType.BentConnector3, 0, 0, 1, 1);
            connector.StartShapeConnectedTo = start;
            connector.StartShapeConnectionSiteIndex = 3;
            connector.EndShapeConnectedTo = end;
            connector.EndShapeConnectionSiteIndex = 1;
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);

        PackageAssert.ElementWithAttributes(package, "ppt/slides/slide1.xml", "//p:cxnSp/p:nvCxnSpPr/p:cNvCxnSpPr/a:stCxn", ("idx", "3"));
        PackageAssert.ElementWithAttributes(package, "ppt/slides/slide1.xml", "//p:cxnSp/p:nvCxnSpPr/p:cNvCxnSpPr/a:endCxn", ("idx", "1"));
    }

    [Fact]
    public void ALineWidthInPointsIsWrittenInEnglishMetricUnits()
    {
        var path = _workspace.PathFor("line-width.pptx");

        using (var presentation = new Presentation())
        {
            var shape = presentation.Slides[0].Shapes!.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);
            shape.LineFormat.Width = 2.5f;
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);

        // 12,700 EMU to the point.
        PackageAssert.ElementWithAttributes(package, "ppt/slides/slide1.xml", "//p:spPr/a:ln", ("w", "31750"));
    }

    [Fact]
    public void EffectsAreWrittenInTheOrderTheEffectListRequires()
    {
        var path = _workspace.PathFor("effect-order.pptx");

        using (var presentation = new Presentation())
        {
            var shape = presentation.Slides[0].Shapes!.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);
            // Enabled in the reverse of schema order on purpose: the writer, not the caller, owns the order.
            shape.EffectFormat!.EnableSoftEdgeEffect();
            shape.EffectFormat.EnableOuterShadowEffect();
            shape.EffectFormat.EnableGlowEffect();
            shape.EffectFormat.EnableBlurEffect();
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);
        var effectList = PackageAssert.SingleElement(package, "ppt/slides/slide1.xml", "//a:effectLst");

        PackageAssert.ChildrenInSchemaOrder(
            effectList,
            Ns.A + "blur", Ns.A + "fillOverlay", Ns.A + "glow", Ns.A + "innerShdw",
            Ns.A + "outerShdw", Ns.A + "prstShdw", Ns.A + "reflection", Ns.A + "softEdge");
    }

    [Fact]
    public void MergedTableCellsAreWrittenAsASpanAndAContinuation()
    {
        var path = _workspace.PathFor("merged-cells.pptx");

        using (var presentation = new Presentation())
        {
            var table = presentation.Slides[0].Shapes!.AddTable(50, 50, [100, 100], [30, 30]);
            table.MergeCells(table.Rows[0][0], table.Rows[0][1], false);
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);

        PackageAssert.ElementWithAttributes(package, "ppt/slides/slide1.xml", "//a:tr[1]/a:tc[1]", ("gridSpan", "2"));
        PackageAssert.ElementWithAttributes(package, "ppt/slides/slide1.xml", "//a:tr[1]/a:tc[2]", ("hMerge", "1"));
    }
}
