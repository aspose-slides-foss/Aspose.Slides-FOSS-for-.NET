namespace Aspose.Slides.Foss.ConformanceTests.Harness;

/// <summary>
/// The input decks these tests open. They are copied next to the test assembly at build time.
/// </summary>
internal static class TestFixtures
{
    /// <summary>A minimal deck: one slide, one layout, one master, one theme.</summary>
    internal static string SimpleDeck => PathTo("Presentation.pptx");

    private static string PathTo(string fileName)
    {
        var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test_data", fileName);
        Assert.True(File.Exists(path), $"The fixture '{fileName}' was not copied to '{path}'.");
        return path;
    }
}
