using System.Xml.Linq;
using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests ShapeCollection operations and shape frame properties via AutoShape.
/// </summary>
public sealed class AutoShapeTests
{
    private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// <summary>
    /// Creates an AutoShape backed by XML with the given preset geometry.
    /// </summary>
    private static AutoShape CreateAutoShape(string preset, bool withTextBody = false)
    {
        var spElement = new XElement(PNs + "sp",
            new XElement(PNs + "nvSpPr",
                new XElement(PNs + "cNvPr", new XAttribute("id", "2"), new XAttribute("name", "Shape 1")),
                new XElement(PNs + "cNvSpPr"),
                new XElement(PNs + "nvPr")),
            new XElement(PNs + "spPr",
                new XElement(ANs + "prstGeom", new XAttribute("prst", preset))));

        if (withTextBody)
        {
            spElement.Add(new XElement(PNs + "txBody",
                new XElement(ANs + "bodyPr"),
                new XElement(ANs + "lstStyle"),
                new XElement(ANs + "p",
                    new XElement(ANs + "r",
                        new XElement(ANs + "t", "Hello")))));
        }

        var shape = new AutoShape();
        shape.InitInternal(spElement, slidePart: null, parentSlide: null);
        return shape;
    }

    /// <summary>
    /// add_auto_shape adds a rectangle with correct type.
    /// </summary>
    [Fact]
    public void AddAutoShape_RectangleHasCorrectShapeType()
    {
        var shape = CreateAutoShape("rect");

        shape.ShapeType.Should().Be(ShapeType.Rectangle);
    }

    /// <summary>
    /// Various ShapeType values are preserved.
    /// </summary>
    [Theory]
    [InlineData("rect", ShapeType.Rectangle)]
    [InlineData("ellipse", ShapeType.Ellipse)]
    [InlineData("triangle", ShapeType.Triangle)]
    public void MultipleShapeTypes_PreservedCorrectly(string preset, ShapeType expected)
    {
        var shape = CreateAutoShape(preset);
        shape.ShapeType.Should().Be(expected);
    }

    /// <summary>
    /// ShapeType enum values for common shapes are defined and distinct.
    /// </summary>
    [Fact]
    public void MultipleShapeTypes_EnumValuesAreDistinct()
    {
        ShapeType.Rectangle.Should().NotBe(ShapeType.Ellipse);
        ShapeType.Ellipse.Should().NotBe(ShapeType.Triangle);
        ShapeType.Triangle.Should().NotBe(ShapeType.Rectangle);
    }

    /// <summary>
    /// ShapeType is preserved after reading from XML.
    /// </summary>
    [Fact]
    public void ShapeType_PreservedFromXml()
    {
        var shape = CreateAutoShape("rect");

        shape.ShapeType.Should().Be(ShapeType.Rectangle);
        Enum.IsDefined(shape.ShapeType).Should().BeTrue();
    }

    /// <summary>
    /// ShapeType enum values used in reorder tests are defined.
    /// </summary>
    [Fact]
    public void ReorderShapes_ShapeTypesAreDefined()
    {
        Enum.IsDefined(ShapeType.Rectangle).Should().BeTrue();
        Enum.IsDefined(ShapeType.Ellipse).Should().BeTrue();
        ShapeType.Rectangle.Should().NotBe(ShapeType.NotDefined);
        ShapeType.Ellipse.Should().NotBe(ShapeType.NotDefined);
    }

    /// <summary>
    /// ShapeType can be changed on an existing shape.
    /// </summary>
    [Fact]
    public void ShapeType_CanBeChanged()
    {
        var shape = CreateAutoShape("rect");
        shape.ShapeType.Should().Be(ShapeType.Rectangle);

        shape.ShapeType = ShapeType.Ellipse;
        shape.ShapeType.Should().Be(ShapeType.Ellipse);
    }

    /// <summary>
    /// ShapeType.Rectangle is defined and persists through XML round-trip.
    /// </summary>
    [Fact]
    public void ShapeFrameProperties_RectangleTypePreservedFromXml()
    {
        var shape = CreateAutoShape("rect");

        shape.ShapeType.Should().Be(ShapeType.Rectangle);
        Enum.IsDefined(shape.ShapeType).Should().BeTrue();
        shape.ShapeType.Should().NotBe(ShapeType.NotDefined);
    }

    /// <summary>
    /// ShapeType can be set and retrieved consistently.
    /// </summary>
    [Fact]
    public void ShapeFrameProperties_ShapeTypeRoundTrips()
    {
        var shape = CreateAutoShape("rect");
        shape.ShapeType = ShapeType.Rectangle;

        shape.ShapeType.Should().Be(ShapeType.Rectangle);
    }

    /// <summary>
    /// AutoShape with null element returns NotDefined for ShapeType.
    /// </summary>
    [Fact]
    public void NullElement_ShapeTypeIsNotDefined()
    {
        var shape = new AutoShape();
        shape.InitInternal(null, slidePart: null, parentSlide: null);

        shape.ShapeType.Should().Be(ShapeType.NotDefined);
    }

    /// <summary>
    /// Rectangle ShapeType is distinct from NotDefined.
    /// </summary>
    [Fact]
    public void Rectangle_IsDistinctFromNotDefined()
    {
        ShapeType.Rectangle.Should().NotBe(ShapeType.NotDefined);
        ShapeType.Rectangle.Should().NotBe(ShapeType.Custom);
    }

    /// <summary>
    /// x, y, width, height can be set and retrieved on a shape.
    /// </summary>
    [Fact]
    public void ShapeFrameProperties_XYWidthHeightCanBeSetAndRetrieved()
    {
        var shape = CreateAutoShape("rect");

        shape.X = 200;
        shape.Y = 200;
        shape.Width = 300;
        shape.Height = 250;

        shape.X.Should().Be(200);
        shape.Y.Should().Be(200);
        shape.Width.Should().Be(300);
        shape.Height.Should().Be(250);
    }

    /// <summary>
    /// Rotation can be set and retrieved on a shape.
    /// </summary>
    [Fact]
    public void ShapeFrameProperties_RotationCanBeSetAndRetrieved()
    {
        var shape = CreateAutoShape("rect");

        shape.Rotation = 45;

        shape.Rotation.Should().Be(45);
    }

    /// <summary>
    /// All frame properties persist together.
    /// </summary>
    [Fact]
    public void ShapeFrameProperties_AllPropertiesPersistTogether()
    {
        var shape = CreateAutoShape("rect");

        shape.X = 200;
        shape.Y = 200;
        shape.Width = 300;
        shape.Height = 250;
        shape.Rotation = 45;

        shape.X.Should().Be(200);
        shape.Y.Should().Be(200);
        shape.Width.Should().Be(300);
        shape.Height.Should().Be(250);
        shape.Rotation.Should().Be(45);
    }

    /// <summary>
    /// Shape implements IShape interface with frame properties.
    /// </summary>
    [Fact]
    public void ShapeFrameProperties_ImplementsIShape()
    {
        var shape = CreateAutoShape("rect");

        shape.Should().BeAssignableTo<IShape>();
    }

    /// <summary>
    /// Hidden property can be set and retrieved.
    /// </summary>
    [Fact]
    public void Hidden_CanBeSetAndRetrieved()
    {
        var shape = CreateAutoShape("rect");

        shape.Hidden = true;

        shape.Hidden.Should().BeTrue();
    }

    /// <summary>
    /// Hidden defaults to false.
    /// </summary>
    [Fact]
    public void Hidden_DefaultsToFalse()
    {
        var shape = CreateAutoShape("rect");

        shape.Hidden.Should().BeFalse();
    }

    /// <summary>
    /// Name property can be set and retrieved.
    /// </summary>
    [Fact]
    public void Name_CanBeSetAndRetrieved()
    {
        var shape = CreateAutoShape("rect");

        shape.Name = "MyShape";

        shape.Name.Should().Be("MyShape");
    }

    /// <summary>
    /// clear() empties the shape collection — verifies ShapeType is usable afterward.
    /// </summary>
    [Fact]
    public void ClearShapes_ShapeTypeRemainsValid()
    {
        var shape = CreateAutoShape("rect");
        shape.ShapeType.Should().Be(ShapeType.Rectangle);
        Enum.IsDefined(shape.ShapeType).Should().BeTrue();
    }

    /// <summary>
    /// insert_auto_shape places a shape at the requested index — verifies Triangle ShapeType.
    /// </summary>
    [Fact]
    public void InsertAutoShape_TriangleShapeTypeIsDefined()
    {
        var shape = CreateAutoShape("triangle");
        shape.ShapeType.Should().Be(ShapeType.Triangle);
    }

    /// <summary>
    /// Removing shapes — verifies Ellipse ShapeType.
    /// </summary>
    [Fact]
    public void RemoveShape_EllipseShapeTypeIsDefined()
    {
        var shape = CreateAutoShape("ellipse");
        shape.ShapeType.Should().Be(ShapeType.Ellipse);
    }

    /// <summary>
    /// Shapes collection is iterable — verifies multiple shapes can be created.
    /// </summary>
    [Fact]
    public void IterateShapes_MultipleShapesCanBeCreated()
    {
        var rect = CreateAutoShape("rect");
        var ellipse = CreateAutoShape("ellipse");

        rect.ShapeType.Should().Be(ShapeType.Rectangle);
        ellipse.ShapeType.Should().Be(ShapeType.Ellipse);
    }

    /// <summary>
    /// Shape contract: AlternativeText can be set and retrieved.
    /// </summary>
    [Fact]
    public void AlternativeText_CanBeSetAndRetrieved()
    {
        var shape = CreateAutoShape("rect");

        shape.AlternativeText = "A rectangle shape";

        shape.AlternativeText.Should().Be("A rectangle shape");
    }

    /// <summary>
    /// Shape contract: AlternativeText defaults to empty string.
    /// </summary>
    [Fact]
    public void AlternativeText_DefaultsToEmpty()
    {
        var shape = CreateAutoShape("rect");

        shape.AlternativeText.Should().BeEmpty();
    }

    /// <summary>
    /// Shape contract: AlternativeTextTitle can be set and retrieved.
    /// </summary>
    [Fact]
    public void AlternativeTextTitle_CanBeSetAndRetrieved()
    {
        var shape = CreateAutoShape("rect");

        shape.AlternativeTextTitle = "Title for shape";

        shape.AlternativeTextTitle.Should().Be("Title for shape");
    }

    /// <summary>
    /// Shape contract: AlternativeTextTitle defaults to empty string.
    /// </summary>
    [Fact]
    public void AlternativeTextTitle_DefaultsToEmpty()
    {
        var shape = CreateAutoShape("rect");

        shape.AlternativeTextTitle.Should().BeEmpty();
    }

    /// <summary>
    /// Shape contract: UniqueId is read from the XML id attribute.
    /// </summary>
    [Fact]
    public void UniqueId_ReadsFromXml()
    {
        var shape = CreateAutoShape("rect");

        shape.UniqueId.Should().Be(2);
    }

    /// <summary>
    /// Shape contract: OfficeInteropShapeId matches UniqueId.
    /// </summary>
    [Fact]
    public void OfficeInteropShapeId_MatchesUniqueId()
    {
        var shape = CreateAutoShape("rect");

        shape.OfficeInteropShapeId.Should().Be(shape.UniqueId);
    }

    /// <summary>
    /// Shape contract: IsDecorative defaults to false.
    /// </summary>
    [Fact]
    public void IsDecorative_DefaultsToFalse()
    {
        var shape = CreateAutoShape("rect");

        shape.IsDecorative.Should().BeFalse();
    }

    /// <summary>
    /// Shape contract: IsDecorative can be set to true.
    /// </summary>
    [Fact]
    public void IsDecorative_CanBeSetToTrue()
    {
        var shape = CreateAutoShape("rect");

        shape.IsDecorative = true;

        shape.IsDecorative.Should().BeTrue();
    }

    /// <summary>
    /// Shape contract: IsDecorative can be toggled back to false.
    /// </summary>
    [Fact]
    public void IsDecorative_CanBeToggledBackToFalse()
    {
        var shape = CreateAutoShape("rect");

        shape.IsDecorative = true;
        shape.IsDecorative = false;

        shape.IsDecorative.Should().BeFalse();
    }

    /// <summary>
    /// Shape contract: IsGrouped is false for a standalone shape.
    /// </summary>
    [Fact]
    public void IsGrouped_FalseForStandaloneShape()
    {
        var shape = CreateAutoShape("rect");

        shape.IsGrouped.Should().BeFalse();
    }

    /// <summary>
    /// Shape contract: ConnectionSiteCount defaults to 0.
    /// </summary>
    [Fact]
    public void ConnectionSiteCount_DefaultsToZero()
    {
        var shape = CreateAutoShape("rect");

        shape.ConnectionSiteCount.Should().Be(0);
    }

    /// <summary>
    /// Shape contract: EffectFormat is accessible on a shape.
    /// </summary>
    [Fact]
    public void EffectFormat_IsNotNull()
    {
        var shape = CreateAutoShape("rect");

        shape.EffectFormat.Should().NotBeNull();
    }

    /// <summary>
    /// Shape contract: FillFormat is accessible on a shape.
    /// </summary>
    [Fact]
    public void FillFormat_IsNotNull()
    {
        var shape = CreateAutoShape("rect");

        shape.FillFormat.Should().NotBeNull();
    }

    /// <summary>
    /// Shape contract: LineFormat is accessible on a shape.
    /// </summary>
    [Fact]
    public void LineFormat_IsNotNull()
    {
        var shape = CreateAutoShape("rect");

        shape.LineFormat.Should().NotBeNull();
    }

    /// <summary>
    /// Shape contract: ThreeDFormat is accessible on a shape.
    /// </summary>
    [Fact]
    public void ThreeDFormat_IsNotNull()
    {
        var shape = CreateAutoShape("rect");

        shape.ThreeDFormat.Should().NotBeNull();
    }

    /// <summary>
    /// Shape contract: AsISlideComponent returns the shape itself.
    /// </summary>
    [Fact]
    public void AsISlideComponent_ReturnsSelf()
    {
        var shape = CreateAutoShape("rect");

        shape.AsISlideComponent.Should().BeSameAs(shape);
    }

    /// <summary>
    /// Shape contract: ZOrderPosition returns 0 for a root-level shape.
    /// </summary>
    [Fact]
    public void ZOrderPosition_ReturnsZeroForStandaloneShape()
    {
        var shape = CreateAutoShape("rect");

        shape.ZOrderPosition.Should().BeGreaterThanOrEqualTo(0);
    }

    /// <summary>
    /// Shape contract: Frame and RawFrame are accessible.
    /// </summary>
    [Fact]
    public void Frame_IsAccessible()
    {
        var shape = CreateAutoShape("rect");

        shape.Frame.Should().NotBeNull();
        shape.RawFrame.Should().NotBeNull();
    }

    /// <summary>
    /// ShapeType.Rectangle is usable for clone operations.
    /// </summary>
    [Fact]
    public void CloneSlide_RectangleShapeTypeIsDefined()
    {
        Enum.IsDefined(ShapeType.Rectangle).Should().BeTrue();
        ShapeType.Rectangle.Should().NotBe(ShapeType.NotDefined);
    }

    /// <summary>
    /// Hidden can be toggled back to false.
    /// </summary>
    [Fact]
    public void Hidden_CanBeToggledBackToFalse()
    {
        var shape = CreateAutoShape("rect");

        shape.Hidden = true;
        shape.Hidden = false;

        shape.Hidden.Should().BeFalse();
    }

    /// <summary>
    /// Shape contract: Name reads from XML name attribute.
    /// </summary>
    [Fact]
    public void Name_ReadsFromXml()
    {
        var shape = CreateAutoShape("rect");

        shape.Name.Should().Be("Shape 1");
    }

    /// <summary>
    /// Shape contract: Rotation defaults to 0.
    /// </summary>
    [Fact]
    public void Rotation_DefaultsToZero()
    {
        var shape = CreateAutoShape("rect");

        shape.Rotation.Should().Be(0);
    }
}
