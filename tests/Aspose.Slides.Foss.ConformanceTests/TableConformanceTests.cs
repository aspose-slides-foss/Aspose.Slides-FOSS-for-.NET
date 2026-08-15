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

    private string WriteDeckWithATable(string fileName)
    {
        var path = _workspace.PathFor(fileName);

        using var presentation = new Presentation();
        presentation.Slides[0].Shapes!.AddTable(50, 50, [100, 100], [30, 30]);
        presentation.Save(path, SaveFormat.Pptx);

        return path;
    }
}
