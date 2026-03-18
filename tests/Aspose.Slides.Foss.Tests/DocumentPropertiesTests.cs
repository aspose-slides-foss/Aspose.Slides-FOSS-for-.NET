using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Internal;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests DocumentProperties: core + custom properties.
/// </summary>
public sealed class DocumentPropertiesTests
{
    /// <summary>
    /// Creates a DocumentProperties backed by an empty OpcPackage.
    /// </summary>
    private static DocumentProperties CreateDocumentProperties()
    {
        var package = new OpcPackage();
        var props = new DocumentProperties();
        props.InitInternal(package);
        return props;
    }

    // ── Core properties (test_core_properties) ──────────────────

    /// <summary>
    /// Title can be set and retrieved.
    /// </summary>
    [Fact]
    public void CoreProperties_TitleCanBeSetAndRetrieved()
    {
        var props = CreateDocumentProperties();

        props.Title = "My Presentation";

        props.Title.Should().Be("My Presentation");
    }

    /// <summary>
    /// Subject can be set and retrieved.
    /// </summary>
    [Fact]
    public void CoreProperties_SubjectCanBeSetAndRetrieved()
    {
        var props = CreateDocumentProperties();

        props.Subject = "Demo Subject";

        props.Subject.Should().Be("Demo Subject");
    }

    /// <summary>
    /// Author can be set and retrieved.
    /// </summary>
    [Fact]
    public void CoreProperties_AuthorCanBeSetAndRetrieved()
    {
        var props = CreateDocumentProperties();

        props.Author = "John Doe";

        props.Author.Should().Be("John Doe");
    }

    /// <summary>
    /// Keywords can be set and retrieved.
    /// </summary>
    [Fact]
    public void CoreProperties_KeywordsCanBeSetAndRetrieved()
    {
        var props = CreateDocumentProperties();

        props.Keywords = "demo, test";

        props.Keywords.Should().Be("demo, test");
    }

    /// <summary>
    /// Category can be set and retrieved.
    /// </summary>
    [Fact]
    public void CoreProperties_CategoryCanBeSetAndRetrieved()
    {
        var props = CreateDocumentProperties();

        props.Category = "Examples";

        props.Category.Should().Be("Examples");
    }

    /// <summary>
    /// All five core properties can be set simultaneously.
    /// </summary>
    [Fact]
    public void CoreProperties_AllFivePersistTogether()
    {
        var props = CreateDocumentProperties();

        props.Title = "My Presentation";
        props.Subject = "Demo Subject";
        props.Author = "John Doe";
        props.Keywords = "demo, test";
        props.Category = "Examples";

        props.Title.Should().Be("My Presentation");
        props.Subject.Should().Be("Demo Subject");
        props.Author.Should().Be("John Doe");
        props.Keywords.Should().Be("demo, test");
        props.Category.Should().Be("Examples");
    }

    /// <summary>
    /// Core string properties default to empty string.
    /// </summary>
    [Fact]
    public void CoreProperties_DefaultToEmptyString()
    {
        var props = CreateDocumentProperties();

        props.Title.Should().BeEmpty();
        props.Subject.Should().BeEmpty();
        props.Author.Should().BeEmpty();
        props.Keywords.Should().BeEmpty();
        props.Category.Should().BeEmpty();
    }

    // ── Custom string property (test_custom_string_property) ────

    /// <summary>
    /// Custom string properties can be set and retrieved.
    /// </summary>
    [Fact]
    public void CustomStringProperty_CanBeSetAndRetrieved()
    {
        var props = CreateDocumentProperties();

        props.SetCustomPropertyValue("MyProp", "hello");

        props.GetCustomPropertyValue("MyProp").Should().Be("hello");
    }

    /// <summary>
    /// Setting a custom property increases count.
    /// </summary>
    [Fact]
    public void CustomStringProperty_IncreasesCount()
    {
        var props = CreateDocumentProperties();

        props.CountOfCustomProperties.Should().Be(0);

        props.SetCustomPropertyValue("MyProp", "hello");

        props.CountOfCustomProperties.Should().Be(1);
    }

    // ── Custom int property (test_custom_int_property) ──────────

    /// <summary>
    /// Custom integer properties can be set and retrieved.
    /// </summary>
    [Fact]
    public void CustomIntProperty_CanBeSetAndRetrieved()
    {
        var props = CreateDocumentProperties();

        props.SetCustomPropertyValue("Count", 42);

        props.GetCustomPropertyValue("Count").Should().Be(42);
    }

    /// <summary>
    /// Custom int and string properties can coexist.
    /// </summary>
    [Fact]
    public void CustomProperties_MixedTypesCoexist()
    {
        var props = CreateDocumentProperties();

        props.SetCustomPropertyValue("Name", "hello");
        props.SetCustomPropertyValue("Count", 42);

        props.CountOfCustomProperties.Should().Be(2);
        props.GetCustomPropertyValue("Name").Should().Be("hello");
        props.GetCustomPropertyValue("Count").Should().Be(42);
    }

    // ── Remove custom property (test_remove_custom_property) ────

    /// <summary>
    /// Removing a custom property decreases count.
    /// </summary>
    [Fact]
    public void RemoveCustomProperty_DecreasesCount()
    {
        var props = CreateDocumentProperties();
        props.SetCustomPropertyValue("A", "val");
        props.SetCustomPropertyValue("B", "val");
        props.CountOfCustomProperties.Should().Be(2);

        props.RemoveCustomProperty("A");

        props.CountOfCustomProperties.Should().Be(1);
    }

    /// <summary>
    /// After removal, ContainsCustomProperty returns false for removed property.
    /// </summary>
    [Fact]
    public void RemoveCustomProperty_ContainsReturnsFalse()
    {
        var props = CreateDocumentProperties();
        props.SetCustomPropertyValue("A", "val");
        props.SetCustomPropertyValue("B", "val");

        props.RemoveCustomProperty("A");

        props.ContainsCustomProperty("A").Should().BeFalse();
        props.ContainsCustomProperty("B").Should().BeTrue();
    }

    /// <summary>
    /// RemoveCustomProperty returns true when property existed.
    /// </summary>
    [Fact]
    public void RemoveCustomProperty_ReturnsTrueWhenExists()
    {
        var props = CreateDocumentProperties();
        props.SetCustomPropertyValue("A", "val");

        props.RemoveCustomProperty("A").Should().BeTrue();
    }

    /// <summary>
    /// RemoveCustomProperty returns false when property does not exist.
    /// </summary>
    [Fact]
    public void RemoveCustomProperty_ReturnsFalseWhenNotExists()
    {
        var props = CreateDocumentProperties();

        props.RemoveCustomProperty("NonExistent").Should().BeFalse();
    }

    // ── Additional DocumentProperties contract tests ────────────

    /// <summary>
    /// DocumentProperties can be instantiated.
    /// </summary>
    [Fact]
    public void DocumentProperties_CanBeInstantiated()
    {
        var props = new DocumentProperties();
        props.Should().NotBeNull();
    }

    /// <summary>
    /// DocumentProperties implements IDocumentProperties.
    /// </summary>
    [Fact]
    public void DocumentProperties_ImplementsInterface()
    {
        var props = CreateDocumentProperties();
        props.Should().BeAssignableTo<IDocumentProperties>();
    }

    /// <summary>
    /// GetCustomPropertyValue returns null for nonexistent property.
    /// </summary>
    [Fact]
    public void GetCustomPropertyValue_ReturnsNullForNonexistent()
    {
        var props = CreateDocumentProperties();

        props.GetCustomPropertyValue("DoesNotExist").Should().BeNull();
    }

    /// <summary>
    /// ContainsCustomProperty returns false when no custom properties set.
    /// </summary>
    [Fact]
    public void ContainsCustomProperty_ReturnsFalseForEmpty()
    {
        var props = CreateDocumentProperties();

        props.ContainsCustomProperty("Anything").Should().BeFalse();
    }

    /// <summary>
    /// ContainsCustomProperty returns true after setting a property.
    /// </summary>
    [Fact]
    public void ContainsCustomProperty_ReturnsTrueAfterSet()
    {
        var props = CreateDocumentProperties();

        props.SetCustomPropertyValue("Exists", "yes");

        props.ContainsCustomProperty("Exists").Should().BeTrue();
    }

    /// <summary>
    /// SetCustomPropertyValue overwrites an existing property.
    /// </summary>
    [Fact]
    public void SetCustomPropertyValue_OverwritesExisting()
    {
        var props = CreateDocumentProperties();

        props.SetCustomPropertyValue("Key", "first");
        props.SetCustomPropertyValue("Key", "second");

        props.GetCustomPropertyValue("Key").Should().Be("second");
        props.CountOfCustomProperties.Should().Be(1);
    }

    /// <summary>
    /// ClearCustomProperties removes all custom properties.
    /// </summary>
    [Fact]
    public void ClearCustomProperties_RemovesAll()
    {
        var props = CreateDocumentProperties();
        props.SetCustomPropertyValue("A", "val");
        props.SetCustomPropertyValue("B", 42);

        props.ClearCustomProperties();

        props.CountOfCustomProperties.Should().Be(0);
    }

    /// <summary>
    /// ClearBuiltInProperties resets core properties to defaults.
    /// </summary>
    [Fact]
    public void ClearBuiltInProperties_ResetsCoreProperties()
    {
        var props = CreateDocumentProperties();
        props.Title = "Something";
        props.Author = "Someone";

        props.ClearBuiltInProperties();

        props.Title.Should().BeEmpty();
        props.Author.Should().BeEmpty();
    }

    /// <summary>
    /// GetCustomPropertyName returns the name by index.
    /// </summary>
    [Fact]
    public void GetCustomPropertyName_ReturnsCorrectName()
    {
        var props = CreateDocumentProperties();
        props.SetCustomPropertyValue("Alpha", "a");
        props.SetCustomPropertyValue("Beta", "b");

        props.GetCustomPropertyName(0).Should().Be("Alpha");
        props.GetCustomPropertyName(1).Should().Be("Beta");
    }

    /// <summary>
    /// Custom bool property can be set and retrieved.
    /// </summary>
    [Fact]
    public void CustomBoolProperty_CanBeSetAndRetrieved()
    {
        var props = CreateDocumentProperties();

        props.SetCustomPropertyValue("Flag", true);

        props.GetCustomPropertyValue("Flag").Should().Be(true);
    }

    /// <summary>
    /// Custom double property can be set and retrieved.
    /// </summary>
    [Fact]
    public void CustomDoubleProperty_CanBeSetAndRetrieved()
    {
        var props = CreateDocumentProperties();

        props.SetCustomPropertyValue("Score", 3.14);

        props.GetCustomPropertyValue("Score").Should().Be(3.14);
    }

    /// <summary>
    /// Core property: ContentStatus can be set and retrieved.
    /// </summary>
    [Fact]
    public void CoreProperties_ContentStatusCanBeSetAndRetrieved()
    {
        var props = CreateDocumentProperties();

        props.ContentStatus = "Draft";

        props.ContentStatus.Should().Be("Draft");
    }

    /// <summary>
    /// Core property: Comments can be set and retrieved.
    /// </summary>
    [Fact]
    public void CoreProperties_CommentsCanBeSetAndRetrieved()
    {
        var props = CreateDocumentProperties();

        props.Comments = "Review needed";

        props.Comments.Should().Be("Review needed");
    }

    /// <summary>
    /// Core property: LastSavedBy can be set and retrieved.
    /// </summary>
    [Fact]
    public void CoreProperties_LastSavedByCanBeSetAndRetrieved()
    {
        var props = CreateDocumentProperties();

        props.LastSavedBy = "Admin";

        props.LastSavedBy.Should().Be("Admin");
    }

    /// <summary>
    /// Core property: RevisionNumber can be set and retrieved.
    /// </summary>
    [Fact]
    public void CoreProperties_RevisionNumberCanBeSetAndRetrieved()
    {
        var props = CreateDocumentProperties();

        props.RevisionNumber = 5;

        props.RevisionNumber.Should().Be(5);
    }

    /// <summary>
    /// Core property: RevisionNumber defaults to zero.
    /// </summary>
    [Fact]
    public void CoreProperties_RevisionNumberDefaultsToZero()
    {
        var props = CreateDocumentProperties();

        props.RevisionNumber.Should().Be(0);
    }

    /// <summary>
    /// App property: Company can be set and retrieved.
    /// </summary>
    [Fact]
    public void AppProperties_CompanyCanBeSetAndRetrieved()
    {
        var props = CreateDocumentProperties();

        props.Company = "Acme Corp";

        props.Company.Should().Be("Acme Corp");
    }

    /// <summary>
    /// App property: NameOfApplication can be set and retrieved.
    /// </summary>
    [Fact]
    public void AppProperties_NameOfApplicationCanBeSetAndRetrieved()
    {
        var props = CreateDocumentProperties();

        props.NameOfApplication = "TestApp";

        props.NameOfApplication.Should().Be("TestApp");
    }

    /// <summary>
    /// App property: Manager can be set and retrieved.
    /// </summary>
    [Fact]
    public void AppProperties_ManagerCanBeSetAndRetrieved()
    {
        var props = CreateDocumentProperties();

        props.Manager = "Jane Smith";

        props.Manager.Should().Be("Jane Smith");
    }

    /// <summary>
    /// App property: SharedDoc defaults to false.
    /// </summary>
    [Fact]
    public void AppProperties_SharedDocDefaultsToFalse()
    {
        var props = CreateDocumentProperties();

        props.SharedDoc.Should().BeFalse();
    }

    /// <summary>
    /// App property: SharedDoc can be set and retrieved.
    /// </summary>
    [Fact]
    public void AppProperties_SharedDocCanBeSetAndRetrieved()
    {
        var props = CreateDocumentProperties();

        props.SharedDoc = true;

        props.SharedDoc.Should().BeTrue();
    }

    /// <summary>
    /// App property: HyperlinkBase can be set and retrieved.
    /// </summary>
    [Fact]
    public void AppProperties_HyperlinkBaseCanBeSetAndRetrieved()
    {
        var props = CreateDocumentProperties();

        props.HyperlinkBase = "/base/path";

        props.HyperlinkBase.Should().Be("/base/path");
    }

    /// <summary>
    /// App property: PresentationFormat can be set and retrieved.
    /// </summary>
    [Fact]
    public void AppProperties_PresentationFormatCanBeSetAndRetrieved()
    {
        var props = CreateDocumentProperties();

        props.PresentationFormat = "On-screen Show (4:3)";

        props.PresentationFormat.Should().Be("On-screen Show (4:3)");
    }

    /// <summary>
    /// App read-only int properties default to zero.
    /// </summary>
    [Fact]
    public void AppProperties_ReadOnlyIntsDefaultToZero()
    {
        var props = CreateDocumentProperties();

        props.Slides.Should().Be(0);
        props.HiddenSlides.Should().Be(0);
        props.Notes.Should().Be(0);
        props.Paragraphs.Should().Be(0);
        props.Words.Should().Be(0);
        props.MultimediaClips.Should().Be(0);
    }

    /// <summary>
    /// AppVersion defaults to empty string.
    /// </summary>
    [Fact]
    public void AppProperties_AppVersionDefaultsToEmpty()
    {
        var props = CreateDocumentProperties();

        props.AppVersion.Should().BeEmpty();
    }

    /// <summary>
    /// HeadingPairs defaults to empty list.
    /// </summary>
    [Fact]
    public void AppProperties_HeadingPairsDefaultsToEmpty()
    {
        var props = CreateDocumentProperties();

        props.HeadingPairs.Should().BeEmpty();
    }

    /// <summary>
    /// TitlesOfParts defaults to empty list.
    /// </summary>
    [Fact]
    public void AppProperties_TitlesOfPartsDefaultsToEmpty()
    {
        var props = CreateDocumentProperties();

        props.TitlesOfParts.Should().BeEmpty();
    }

    /// <summary>
    /// Multiple custom property types coexist correctly.
    /// </summary>
    [Theory]
    [InlineData("StrProp", "hello")]
    [InlineData("IntProp", 42)]
    [InlineData("BoolProp", true)]
    public void CustomProperties_VariousTypesCanBeSet(string name, object value)
    {
        var props = CreateDocumentProperties();

        props.SetCustomPropertyValue(name, value);

        props.ContainsCustomProperty(name).Should().BeTrue();
        props.GetCustomPropertyValue(name).Should().Be(value);
    }
}
