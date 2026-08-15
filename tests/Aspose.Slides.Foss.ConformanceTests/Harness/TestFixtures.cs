namespace Aspose.Slides.Foss.ConformanceTests.Harness;

/// <summary>
/// The input decks these tests open. They are copied next to the test assembly at build time.
/// </summary>
internal static class TestFixtures
{
    /// <summary>A minimal deck: one slide, one layout, one master, one theme.</summary>
    internal static string SimpleDeck => PathTo("Presentation.pptx");

    /// <summary>
    /// A 46-part deck authored by PowerPoint: two slides, twelve layouts, a notes master with a
    /// theme part of its own, a notes slide, <c>presProps</c>, <c>viewProps</c>, <c>tableStyles</c>
    /// and a thumbnail — the parts a round-trip is most likely to lose, and none of which this
    /// library writes itself. A fixture written by this library cannot stand in for it.
    /// </summary>
    internal static string PowerPointDeck => PathTo("PowerPointDeck.pptx");

    /// <summary>
    /// A one-table deck whose <c>a:tblPr</c> already carries an <c>a:extLst</c>, as a table authored
    /// by an application that stores extensions on it does. Nothing this library writes produces
    /// one, so it has to be opened rather than built.
    /// </summary>
    internal static string TableWithExtension => PathTo("TableWithExtension.pptx");

    private static string PathTo(string fileName)
    {
        var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test_data", fileName);
        Assert.True(File.Exists(path), $"The fixture '{fileName}' was not copied to '{path}'.");
        return path;
    }
}
