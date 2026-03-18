using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.IntegrationTests;

/// <summary>
/// Tests for DocumentProperties: core + custom properties.
/// </summary>
public sealed class DocumentPropertiesTests : IDisposable
{
    private readonly string _tempDir;

    public DocumentPropertiesTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, recursive: true);
    }

    [Fact]
    public void TestCoreProperties()
    {
        // Core properties persist after save/reload.
        using var pres = new Presentation();
        var props = pres.DocumentProperties;
        props.Title = "My Presentation";
        props.Subject = "Demo Subject";
        props.Author = "John Doe";
        props.Keywords = "demo, test";
        props.Category = "Examples";

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var p2 = pres2.DocumentProperties;
        p2.Title.Should().Be("My Presentation");
        p2.Subject.Should().Be("Demo Subject");
        p2.Author.Should().Be("John Doe");
        p2.Keywords.Should().Be("demo, test");
        p2.Category.Should().Be("Examples");
    }

    [Fact]
    public void TestCustomStringProperty()
    {
        // Custom string properties persist.
        using var pres = new Presentation();
        pres.DocumentProperties.SetCustomPropertyValue("MyProp", "hello");

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var value = pres2.DocumentProperties.GetCustomPropertyValue("MyProp");
        value.Should().Be("hello");
    }

    [Fact]
    public void TestCustomIntProperty()
    {
        // Custom integer properties persist.
        using var pres = new Presentation();
        pres.DocumentProperties.SetCustomPropertyValue("Count", 42);

        using var pres2 = TestHelpers.SaveAndReopen(pres, _tempDir);
        var value = pres2.DocumentProperties.GetCustomPropertyValue("Count");
        value.Should().Be(42);
    }

    [Fact]
    public void TestRemoveCustomProperty()
    {
        // Removing a custom property decreases count.
        using var pres = new Presentation();
        var props = pres.DocumentProperties;
        props.SetCustomPropertyValue("A", "val");
        props.SetCustomPropertyValue("B", "val");
        props.CountOfCustomProperties.Should().Be(2);

        props.RemoveCustomProperty("A");
        props.CountOfCustomProperties.Should().Be(1);
        props.ContainsCustomProperty("A").Should().BeFalse();
        props.ContainsCustomProperty("B").Should().BeTrue();
    }
}
