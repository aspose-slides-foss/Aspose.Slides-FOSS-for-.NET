using Aspose.Slides.Foss.ConformanceTests.Harness;
using Aspose.Slides.Foss.Export;

namespace Aspose.Slides.Foss.ConformanceTests;

/// <summary>
/// Table markup is where a name that is nearly right costs the user the whole feature: PowerPoint
/// discards an element it does not recognise without a word, so a misspelled style reference makes
/// every table fall back to the default style and nothing anywhere reports it.
/// </summary>
public sealed class TableConformanceTests : IDisposable
{
    private readonly TestWorkspace _workspace = new();

    public void Dispose() => _workspace.Dispose();

    [Fact]
    public void ATableReferencesItsStyleThroughTheSchemaElementName()
    {
        var path = WriteDeckWithATable("table-style-id.pptx");

        using var package = PptxPackage.Open(path);

        var wrongName = PackageAssert.Select(package, "ppt/slides/slide1.xml", "//a:tblPr/a:tblStyleId");
        Assert.True(wrongName.Count == 0,
            $"The table style is written as <a:tblStyleId>, which is not an element of CT_TableProperties. " +
            $"ECMA-376 names it <a:tableStyleId>.{Environment.NewLine}" +
            $"written: {string.Join(", ", wrongName)}");

        PackageAssert.SingleElement(package, "ppt/slides/slide1.xml", "//a:tblPr/a:tableStyleId");
    }

    [Fact]
    public void ATableProducesASchemaValidPackage()
    {
        var path = WriteDeckWithATable("table-valid.pptx");

        SchemaValidation.HasNoSchemaErrors(path);
    }

    /// <summary>
    /// CT_TableProperties is a sequence: the style reference comes before <c>a:extLst</c>. A deck
    /// authored elsewhere may already carry an extension list on its table, and appending the style
    /// after it makes the table invalid while every element in it is legal — the failure mode a
    /// hand-read of the XML is worst at spotting.
    /// </summary>
    [Fact]
    public void SettingTheStyleOfATableThatCarriesAnExtensionListKeepsSchemaOrder()
    {
        var path = _workspace.PathFor("table-style-order.pptx");

        using (var presentation = new Presentation(TestFixtures.TableWithExtension))
        {
            var table = (ITable)presentation.Slides[0].Shapes![0];
            table.StylePreset = TableStylePreset.MediumStyle2Accent3;
            presentation.Save(path, SaveFormat.Pptx);
        }

        using var package = PptxPackage.Open(path);

        var properties = PackageAssert.SingleElement(package, "ppt/slides/slide1.xml", "//a:tbl/a:tblPr");
        PackageAssert.ChildrenInSchemaOrder(properties, Ns.A + "tableStyleId", Ns.A + "extLst");

        SchemaValidation.HasNoSchemaErrors(path);
    }

    private string WriteDeckWithATable(string fileName)
    {
        var path = _workspace.PathFor(fileName);

        using var presentation = new Presentation();
        presentation.Slides[0].Shapes!.AddTable(50, 50, [100, 100], [30, 30]);
        presentation.Save(path, SaveFormat.Pptx);

        return path;
    }
}
