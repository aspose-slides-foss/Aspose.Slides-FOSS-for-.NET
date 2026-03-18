using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using Aspose.Slides.Foss.Internal;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests for ITableFormat / TableFormat.
/// </summary>
public sealed class ITableFormatTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static readonly string SlideXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
               xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:cSld>
            <p:spTree>
              <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
              <p:grpSpPr/>
            </p:spTree>
          </p:cSld>
        </p:sld>
        """;

    private static (SlidePart slidePart, ShapeCollection shapes) CreateSlideWithShapes()
    {
        var slidePart = new SlidePart();
        slidePart.InitInternal("ppt/slides/slide1.xml");
        slidePart.Element = XDocument.Parse(SlideXml).Root;

        var shapes = new ShapeCollection();
        shapes.InitInternal(slidePart, null);
        return (slidePart, shapes);
    }

    private static ShapeCollection RoundTripShapes(SlidePart slidePart)
    {
        var xml = slidePart.Element!.ToString();
        var reloadedRoot = XDocument.Parse(xml).Root!;

        var newSlidePart = new SlidePart();
        newSlidePart.InitInternal("ppt/slides/slide1.xml");
        newSlidePart.Element = reloadedRoot;

        var shapes = new ShapeCollection();
        shapes.InitInternal(newSlidePart, null);
        return shapes;
    }

    /// <summary>
    /// ITableFormat exposes a FillFormat property.
    /// </summary>
    [Fact]
    public void TableFormat_FillFormat_ReturnsIFillFormat()
    {
        var (_, shapes) = CreateSlideWithShapes();
        var table = shapes.AddTable(50, 50, [100], [30]);

        table.TableFormat.FillFormat.Should().NotBeNull();
        table.TableFormat.FillFormat.Should().BeAssignableTo<IFillFormat>();
    }

    /// <summary>
    /// Table-level fill defaults to NotDefined when no fill is set.
    /// </summary>
    [Fact]
    public void TableFormat_FillFormat_DefaultsToNotDefined()
    {
        var (_, shapes) = CreateSlideWithShapes();
        var table = shapes.AddTable(50, 50, [100], [30]);

        table.TableFormat.FillFormat.FillType.Should().Be(FillType.NotDefined);
    }

    /// <summary>
    /// Table-level solid fill can be set and read back.
    /// </summary>
    [Fact]
    public void TableFormat_SolidFill_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var table = shapes.AddTable(50, 50, [100, 100], [30]);

        table.TableFormat.FillFormat.FillType = FillType.Solid;
        table.TableFormat.FillFormat.SolidFillColor.Color = Color.FromArgb(255, 0, 128, 255);

        var reloaded = RoundTripShapes(slidePart);
        var reloadedTable = (ITable)reloaded[0];

        reloadedTable.TableFormat.FillFormat.FillType.Should().Be(FillType.Solid);
        var c = reloadedTable.TableFormat.FillFormat.SolidFillColor.Color;
        c.Should().NotBeNull();
        c!.R.Should().Be(0);
        c.G.Should().Be(128);
        c.B.Should().Be(255);
    }

    /// <summary>
    /// Table-level NoFill persists after round-trip.
    /// </summary>
    [Fact]
    public void TableFormat_NoFill_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var table = shapes.AddTable(50, 50, [100], [30]);

        table.TableFormat.FillFormat.FillType = FillType.NoFill;

        var reloaded = RoundTripShapes(slidePart);
        var reloadedTable = (ITable)reloaded[0];

        reloadedTable.TableFormat.FillFormat.FillType.Should().Be(FillType.NoFill);
    }

    /// <summary>
    /// Table-level gradient fill type persists after round-trip.
    /// </summary>
    [Fact]
    public void TableFormat_GradientFill_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var table = shapes.AddTable(50, 50, [100], [30]);

        table.TableFormat.FillFormat.FillType = FillType.Gradient;

        var reloaded = RoundTripShapes(slidePart);
        var reloadedTable = (ITable)reloaded[0];

        reloadedTable.TableFormat.FillFormat.FillType.Should().Be(FillType.Gradient);
    }

    /// <summary>
    /// Table-level pattern fill type persists after round-trip.
    /// </summary>
    [Fact]
    public void TableFormat_PatternFill_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var table = shapes.AddTable(50, 50, [100], [30]);

        table.TableFormat.FillFormat.FillType = FillType.Pattern;

        var reloaded = RoundTripShapes(slidePart);
        var reloadedTable = (ITable)reloaded[0];

        reloadedTable.TableFormat.FillFormat.FillType.Should().Be(FillType.Pattern);
    }

    /// <summary>
    /// Table-level picture fill type persists after round-trip.
    /// </summary>
    [Fact]
    public void TableFormat_PictureFill_PersistsAfterRoundTrip()
    {
        var (slidePart, shapes) = CreateSlideWithShapes();
        var table = shapes.AddTable(50, 50, [100], [30]);

        table.TableFormat.FillFormat.FillType = FillType.Picture;

        var reloaded = RoundTripShapes(slidePart);
        var reloadedTable = (ITable)reloaded[0];

        reloadedTable.TableFormat.FillFormat.FillType.Should().Be(FillType.Picture);
    }

    /// <summary>
    /// Cell fill and table fill are independent.
    /// </summary>
    [Fact]
    public void TableFormat_CellFillDoesNotAffectTableFill()
    {
        var (_, shapes) = CreateSlideWithShapes();
        var table = shapes.AddTable(50, 50, [100], [30]);

        var cellFormat = table.Rows[0][0].CellFormat;
        cellFormat.FillFormat.FillType = FillType.Solid;
        cellFormat.FillFormat.SolidFillColor.Color = Color.Red;

        table.TableFormat.FillFormat.FillType.Should().Be(FillType.NotDefined);
    }

    /// <summary>
    /// Cell border fill and table fill are independent.
    /// </summary>
    [Fact]
    public void TableFormat_CellBorderFillDoesNotAffectTableFill()
    {
        var (_, shapes) = CreateSlideWithShapes();
        var table = shapes.AddTable(50, 50, [100], [30]);

        var cellFormat = table.Rows[0][0].CellFormat;
        cellFormat.BorderTop.FillFormat.FillType = FillType.Solid;
        cellFormat.BorderTop.FillFormat.SolidFillColor.Color = Color.Red;
        cellFormat.BorderTop.Width = 3.0f;

        table.TableFormat.FillFormat.FillType.Should().Be(FillType.NotDefined);
    }

    /// <summary>
    /// TableFormat from an XML element with existing solid fill reads correctly.
    /// </summary>
    [Fact]
    public void TableFormat_ReadsExistingSolidFillFromXml()
    {
        var tblPr = new XElement(ANs + "tblPr",
            new XElement(ANs + "solidFill",
                new XElement(ANs + "srgbClr", new XAttribute("val", "FF0000"))));

        var tf = new TableFormat();
        tf.InitInternal(tblPr, slidePart: null, parentSlide: null);

        tf.FillFormat.FillType.Should().Be(FillType.Solid);
        var c = tf.FillFormat.SolidFillColor.Color;
        c.Should().NotBeNull();
        c!.R.Should().Be(255);
        c.G.Should().Be(0);
        c.B.Should().Be(0);
    }

    /// <summary>
    /// TableFormat with null element returns NotDefined fill type.
    /// </summary>
    [Fact]
    public void TableFormat_NullElement_ReturnsFillTypeNotDefined()
    {
        var tf = new TableFormat();

        tf.FillFormat.FillType.Should().Be(FillType.NotDefined);
    }
}
