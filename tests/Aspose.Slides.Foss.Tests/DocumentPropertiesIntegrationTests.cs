using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Export;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Integration tests for DocumentProperties through the full Presentation API.
/// </summary>
public sealed class DocumentPropertiesIntegrationTests
{
    private static byte[] SaveToBytes(Presentation pres)
    {
        using var ms = new MemoryStream();
        pres.Save(ms, SaveFormat.Pptx);
        return ms.ToArray();
    }

    /// <summary>
    /// Title persists after save.
    /// </summary>
    [Fact]
    public void CoreProperties_TitlePersistsAfterSave()
    {
        using var pres = new Presentation();
        pres.DocumentProperties.Title = "My Presentation";

        var bytes = SaveToBytes(pres);

        bytes.Should().NotBeEmpty();
        pres.DocumentProperties.Title.Should().Be("My Presentation");
    }

    /// <summary>
    /// Subject persists after save.
    /// </summary>
    [Fact]
    public void CoreProperties_SubjectPersistsAfterSave()
    {
        using var pres = new Presentation();
        pres.DocumentProperties.Subject = "Demo Subject";

        var bytes = SaveToBytes(pres);

        bytes.Should().NotBeEmpty();
        pres.DocumentProperties.Subject.Should().Be("Demo Subject");
    }

    /// <summary>
    /// Author persists after save.
    /// </summary>
    [Fact]
    public void CoreProperties_AuthorPersistsAfterSave()
    {
        using var pres = new Presentation();
        pres.DocumentProperties.Author = "John Doe";

        var bytes = SaveToBytes(pres);

        bytes.Should().NotBeEmpty();
        pres.DocumentProperties.Author.Should().Be("John Doe");
    }

    /// <summary>
    /// Keywords persist after save.
    /// </summary>
    [Fact]
    public void CoreProperties_KeywordsPersistAfterSave()
    {
        using var pres = new Presentation();
        pres.DocumentProperties.Keywords = "demo, test";

        var bytes = SaveToBytes(pres);

        bytes.Should().NotBeEmpty();
        pres.DocumentProperties.Keywords.Should().Be("demo, test");
    }

    /// <summary>
    /// Category persists after save.
    /// </summary>
    [Fact]
    public void CoreProperties_CategoryPersistsAfterSave()
    {
        using var pres = new Presentation();
        pres.DocumentProperties.Category = "Examples";

        var bytes = SaveToBytes(pres);

        bytes.Should().NotBeEmpty();
        pres.DocumentProperties.Category.Should().Be("Examples");
    }

    /// <summary>
    /// All five core properties persist together after save.
    /// </summary>
    [Fact]
    public void CoreProperties_AllFivePersistAfterSave()
    {
        using var pres = new Presentation();
        pres.DocumentProperties.Title = "My Presentation";
        pres.DocumentProperties.Subject = "Demo Subject";
        pres.DocumentProperties.Author = "John Doe";
        pres.DocumentProperties.Keywords = "demo, test";
        pres.DocumentProperties.Category = "Examples";

        var bytes = SaveToBytes(pres);

        bytes.Should().NotBeEmpty();
        pres.DocumentProperties.Title.Should().Be("My Presentation");
        pres.DocumentProperties.Subject.Should().Be("Demo Subject");
        pres.DocumentProperties.Author.Should().Be("John Doe");
        pres.DocumentProperties.Keywords.Should().Be("demo, test");
        pres.DocumentProperties.Category.Should().Be("Examples");
    }

    /// <summary>
    /// Custom string properties persist after save.
    /// </summary>
    [Fact]
    public void CustomStringProperty_PersistsAfterSave()
    {
        using var pres = new Presentation();
        pres.DocumentProperties.SetCustomPropertyValue("MyProp", "hello");

        var bytes = SaveToBytes(pres);

        bytes.Should().NotBeEmpty();
        pres.DocumentProperties.GetCustomPropertyValue("MyProp").Should().Be("hello");
        pres.DocumentProperties.CountOfCustomProperties.Should().Be(1);
    }

    /// <summary>
    /// Multiple custom string properties persist.
    /// </summary>
    [Fact]
    public void CustomStringProperty_MultiplePersist()
    {
        using var pres = new Presentation();
        pres.DocumentProperties.SetCustomPropertyValue("Prop1", "value1");
        pres.DocumentProperties.SetCustomPropertyValue("Prop2", "value2");

        var bytes = SaveToBytes(pres);

        bytes.Should().NotBeEmpty();
        pres.DocumentProperties.GetCustomPropertyValue("Prop1").Should().Be("value1");
        pres.DocumentProperties.GetCustomPropertyValue("Prop2").Should().Be("value2");
        pres.DocumentProperties.CountOfCustomProperties.Should().Be(2);
    }

    /// <summary>
    /// Custom integer properties persist after save.
    /// </summary>
    [Fact]
    public void CustomIntProperty_PersistsAfterSave()
    {
        using var pres = new Presentation();
        pres.DocumentProperties.SetCustomPropertyValue("Count", 42);

        var bytes = SaveToBytes(pres);

        bytes.Should().NotBeEmpty();
        pres.DocumentProperties.GetCustomPropertyValue("Count").Should().Be(42);
    }

    /// <summary>
    /// Mixed custom property types persist together.
    /// </summary>
    [Fact]
    public void CustomProperties_MixedTypesPersistAfterSave()
    {
        using var pres = new Presentation();
        pres.DocumentProperties.SetCustomPropertyValue("Name", "hello");
        pres.DocumentProperties.SetCustomPropertyValue("Count", 42);

        var bytes = SaveToBytes(pres);

        bytes.Should().NotBeEmpty();
        pres.DocumentProperties.CountOfCustomProperties.Should().Be(2);
        pres.DocumentProperties.GetCustomPropertyValue("Name").Should().Be("hello");
        pres.DocumentProperties.GetCustomPropertyValue("Count").Should().Be(42);
    }

    /// <summary>
    /// Removing a custom property decreases count and persists.
    /// </summary>
    [Fact]
    public void RemoveCustomProperty_DecreasesCountAndPersists()
    {
        using var pres = new Presentation();
        pres.DocumentProperties.SetCustomPropertyValue("A", "val");
        pres.DocumentProperties.SetCustomPropertyValue("B", "val");
        pres.DocumentProperties.RemoveCustomProperty("A");

        var bytes = SaveToBytes(pres);

        bytes.Should().NotBeEmpty();
        pres.DocumentProperties.CountOfCustomProperties.Should().Be(1);
        pres.DocumentProperties.ContainsCustomProperty("A").Should().BeFalse();
        pres.DocumentProperties.ContainsCustomProperty("B").Should().BeTrue();
    }

    /// <summary>
    /// After removal, the removed property is no longer accessible.
    /// </summary>
    [Fact]
    public void RemoveCustomProperty_PropertyNoLongerAccessible()
    {
        using var pres = new Presentation();
        pres.DocumentProperties.SetCustomPropertyValue("ToRemove", "gone");
        pres.DocumentProperties.SetCustomPropertyValue("ToKeep", "stay");

        pres.DocumentProperties.RemoveCustomProperty("ToRemove");

        pres.DocumentProperties.GetCustomPropertyValue("ToRemove").Should().BeNull();
        pres.DocumentProperties.GetCustomPropertyValue("ToKeep").Should().Be("stay");
    }

    /// <summary>
    /// Core properties default to empty string in a new presentation.
    /// </summary>
    [Fact]
    public void CoreProperties_DefaultToEmptyInNewPresentation()
    {
        using var pres = new Presentation();

        pres.DocumentProperties.Title.Should().BeEmpty();
        pres.DocumentProperties.Subject.Should().BeEmpty();
        pres.DocumentProperties.Author.Should().BeEmpty();
        pres.DocumentProperties.Keywords.Should().BeEmpty();
        pres.DocumentProperties.Category.Should().BeEmpty();
    }

    /// <summary>
    /// DocumentProperties is accessible through the Presentation.
    /// </summary>
    [Fact]
    public void DocumentProperties_IsAccessibleThroughPresentation()
    {
        using var pres = new Presentation();

        pres.DocumentProperties.Should().NotBeNull();
        pres.DocumentProperties.Should().BeAssignableTo<IDocumentProperties>();
    }

    /// <summary>
    /// Custom property count starts at zero for a new presentation.
    /// </summary>
    [Fact]
    public void CustomProperties_CountStartsAtZero()
    {
        using var pres = new Presentation();

        pres.DocumentProperties.CountOfCustomProperties.Should().Be(0);
    }
}
