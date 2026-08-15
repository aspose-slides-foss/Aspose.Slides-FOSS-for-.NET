namespace Aspose.Slides.Foss.ConformanceTests.Harness;

/// <summary>
/// A temporary directory that produced packages are written to, and the small fixtures the tests
/// need. Deleted when the test class is disposed.
/// </summary>
internal sealed class TestWorkspace : IDisposable
{
    private readonly string _root;

    internal TestWorkspace()
    {
        _root = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "slides-foss-conformance", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_root);
    }

    /// <summary>Returns a path inside the workspace for a file that a test is about to write.</summary>
    internal string PathFor(string fileName) => System.IO.Path.Combine(_root, fileName);

    /// <summary>
    /// A valid 1×1 PNG. Small on purpose: these tests are about what the package says about the
    /// image, never about the pixels.
    /// </summary>
    internal static byte[] Png() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAAD0lEQVR4nGP8z4AAjEgcAB3+AP/7q6L4AAAAAElFTkSuQmCC");

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_root))
                Directory.Delete(_root, recursive: true);
        }
        catch (IOException)
        {
            // A leftover temp directory is not worth failing a test run over.
        }
    }
}
