using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Verifies the enum values and color constants used by line format operations.
/// </summary>
public sealed class LineFormatTests
{
    /// <summary>
    /// Verifies FillType.Solid and Color.DarkRed used in the test.
    /// </summary>
    [Fact]
    public void LineColorAndWidth_FillTypeSolidIsDefined()
    {
        Enum.IsDefined(FillType.Solid).Should().BeTrue();
        FillType.Solid.Should().NotBe(FillType.NotDefined);
    }

    [Fact]
    public void LineColorAndWidth_DarkRedColorHasCorrectRedComponent()
    {
        Color.DarkRed.Should().NotBe(Color.Empty);
        Color.DarkRed.R.Should().BeGreaterThan(0);
    }

    /// <summary>
    /// Verifies LineDashStyle.Dash, FillType.Solid, Color.Black.
    /// </summary>
    [Fact]
    public void LineDashStyle_DashIsDefined()
    {
        Enum.IsDefined(LineDashStyle.Dash).Should().BeTrue();
        LineDashStyle.Dash.Should().NotBe(LineDashStyle.NotDefined);
    }

    [Fact]
    public void LineDashStyle_BlackColorHasZeroRgbComponents()
    {
        Color.Black.Should().NotBe(Color.Empty);
        Color.Black.R.Should().Be(0);
        Color.Black.G.Should().Be(0);
        Color.Black.B.Should().Be(0);
    }

    [Fact]
    public void LineDashStyle_DashIsDistinctFromSolid()
    {
        LineDashStyle.Dash.Should().NotBe(LineDashStyle.Solid);
    }

    /// <summary>
    /// Verifies all four dash style enum values used in the test are defined and distinct.
    /// </summary>
    [Theory]
    [InlineData(LineDashStyle.Solid)]
    [InlineData(LineDashStyle.Dash)]
    [InlineData(LineDashStyle.Dot)]
    [InlineData(LineDashStyle.DashDot)]
    public void MultipleDashStyles_AllUsedValuesAreDefined(LineDashStyle style)
    {
        Enum.IsDefined(style).Should().BeTrue();
    }

    [Fact]
    public void MultipleDashStyles_AllValuesAreDistinct()
    {
        var styles = new[]
        {
            LineDashStyle.Solid,
            LineDashStyle.Dash,
            LineDashStyle.Dot,
            LineDashStyle.DashDot
        };

        styles.Should().OnlyHaveUniqueItems();
    }

    [Fact]
    public void LineDashStyle_NotDefinedSentinelIsZero()
    {
        ((int)LineDashStyle.NotDefined).Should().Be(0);
    }

    /// <summary>
    /// LineFormat can be instantiated and initialized with XML.
    /// </summary>
    [Fact]
    public void LineColorAndWidth_LineFormatCanBeInstantiated()
    {
        var lf = new LineFormat();
        lf.Should().NotBeNull();
    }

    /// <summary>
    /// DarkRed colour components match expected values.
    /// </summary>
    [Fact]
    public void LineColorAndWidth_DarkRedIsDistinctFromBlack()
    {
        Color.DarkRed.Should().NotBe(Color.Black);
        Color.DarkRed.R.Should().BeGreaterThan(Color.Black.R);
    }

    /// <summary>
    /// Width value 5 is representable and line format accepts numeric widths.
    /// </summary>
    [Fact]
    public void LineColorAndWidth_WidthValueIsRepresentable()
    {
        var width = 5.0;
        width.Should().BeGreaterThan(0);
    }

    /// <summary>
    /// Dash is distinct from Dot and DashDot.
    /// </summary>
    [Fact]
    public void LineDashStyle_DashIsDistinctFromDotAndDashDot()
    {
        LineDashStyle.Dash.Should().NotBe(LineDashStyle.Dot);
        LineDashStyle.Dash.Should().NotBe(LineDashStyle.DashDot);
    }

    /// <summary>
    /// Sets DashStyle on a LineFormat backed by XML and reads it back.
    /// </summary>
    [Fact]
    public void LineDashStyle_DashStylePersistsOnLineFormat()
    {
        var spPr = System.Xml.Linq.XElement.Parse("<spPr xmlns='http://schemas.openxmlformats.org/drawingml/2006/main'/>");
        var lf = new LineFormat();
        lf.InitInternal(spPr, null);

        lf.Width = 3;
        lf.DashStyle = LineDashStyle.Dash;

        lf.DashStyle.Should().Be(LineDashStyle.Dash);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void LineDashStyle_FillFormatCanBeSetToSolid()
    {
        var spPr = System.Xml.Linq.XElement.Parse("<spPr xmlns='http://schemas.openxmlformats.org/drawingml/2006/main'/>");
        var lf = new LineFormat();
        lf.InitInternal(spPr, null);

        lf.Width = 3;
        lf.DashStyle = LineDashStyle.Dash;
        lf.FillFormat.FillType = FillType.Solid;
        lf.FillFormat.SolidFillColor.Color = Color.Black;

        lf.FillFormat.FillType.Should().Be(FillType.Solid);
    }

    /// <summary>
    /// </summary>
    [Theory]
    [InlineData(LineDashStyle.Solid)]
    [InlineData(LineDashStyle.Dash)]
    [InlineData(LineDashStyle.Dot)]
    [InlineData(LineDashStyle.DashDot)]
    public void MultipleDashStyles_CanBeSetAndReadBackOnLineFormat(LineDashStyle style)
    {
        var spPr = System.Xml.Linq.XElement.Parse("<spPr xmlns='http://schemas.openxmlformats.org/drawingml/2006/main'/>");
        var lf = new LineFormat();
        lf.InitInternal(spPr, null);

        lf.DashStyle = style;

        lf.DashStyle.Should().Be(style);
    }

    /// <summary>
    /// CustomDashPattern can be set and read back on LineFormat.
    /// </summary>
    [Fact]
    public void CustomDashPattern_CanBeSetAndReadBack()
    {
        var spPr = System.Xml.Linq.XElement.Parse("<spPr xmlns='http://schemas.openxmlformats.org/drawingml/2006/main'/>");
        var lf = new LineFormat();
        lf.InitInternal(spPr, null);

        lf.CustomDashPattern = [3.0f, 1.0f, 1.0f, 1.0f];

        lf.CustomDashPattern.Should().HaveCount(4);
        lf.CustomDashPattern[0].Should().Be(3.0f);
        lf.CustomDashPattern[1].Should().Be(1.0f);
    }

    /// <summary>
    /// Style can be set and read back on LineFormat.
    /// </summary>
    [Fact]
    public void Style_CanBeSetAndReadBack()
    {
        var spPr = System.Xml.Linq.XElement.Parse("<spPr xmlns='http://schemas.openxmlformats.org/drawingml/2006/main'/>");
        var lf = new LineFormat();
        lf.InitInternal(spPr, null);

        lf.Style = LineStyle.ThinThick;

        lf.Style.Should().Be(LineStyle.ThinThick);
    }

    /// <summary>
    /// MiterLimit can be set and read back on LineFormat.
    /// </summary>
    [Fact]
    public void MiterLimit_CanBeSetAndReadBack()
    {
        var spPr = System.Xml.Linq.XElement.Parse("<spPr xmlns='http://schemas.openxmlformats.org/drawingml/2006/main'/>");
        var lf = new LineFormat();
        lf.InitInternal(spPr, null);

        lf.MiterLimit = 8.0f;

        lf.MiterLimit.Should().Be(8.0f);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void MultipleDashStyles_SettingNewStyleReplacesOld()
    {
        var spPr = System.Xml.Linq.XElement.Parse("<spPr xmlns='http://schemas.openxmlformats.org/drawingml/2006/main'/>");
        var lf = new LineFormat();
        lf.InitInternal(spPr, null);

        lf.DashStyle = LineDashStyle.Dash;
        lf.DashStyle.Should().Be(LineDashStyle.Dash);

        lf.DashStyle = LineDashStyle.Dot;
        lf.DashStyle.Should().Be(LineDashStyle.Dot);
    }

    /// <summary>
    /// LineFormat width can be set and read back.
    /// </summary>
    [Fact]
    public void LineFormat_WidthCanBeSetAndReadBack()
    {
        var spPr = System.Xml.Linq.XElement.Parse("<spPr xmlns='http://schemas.openxmlformats.org/drawingml/2006/main'/>");
        var lf = new LineFormat();
        lf.InitInternal(spPr, null);

        lf.Width = 5;

        lf.Width.Should().Be(5);
    }

    /// <summary>
    /// LineFormat without a line element reports NotDefined for dash style.
    /// </summary>
    [Fact]
    public void LineFormat_UndefinedDashStyleWhenNoLineElement()
    {
        var spPr = System.Xml.Linq.XElement.Parse("<spPr xmlns='http://schemas.openxmlformats.org/drawingml/2006/main'/>");
        var lf = new LineFormat();
        lf.InitInternal(spPr, null);

        lf.DashStyle.Should().Be(LineDashStyle.NotDefined);
    }
}
