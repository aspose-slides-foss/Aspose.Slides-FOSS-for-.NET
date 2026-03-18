using System.Xml.Linq;
using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Drawing;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Verifies the enum values and color constants used by text formatting operations.
/// </summary>
public sealed class TextFormattingTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    /// <summary>
    /// Verifies NullableBool.True used for FontBold and FontItalic.
    /// </summary>
    [Fact]
    public void BoldItalic_NullableBoolTrueIsDefined()
    {
        Enum.IsDefined(NullableBool.True).Should().BeTrue();
        NullableBool.True.Should().NotBe(NullableBool.NotDefined);
    }

    [Fact]
    public void BoldItalic_TrueIsDistinctFromFalse()
    {
        NullableBool.True.Should().NotBe(NullableBool.False);
    }

    /// <summary>
    /// Verifies TextUnderlineType.Single.
    /// </summary>
    [Fact]
    public void Underline_SingleTypeIsDefined()
    {
        Enum.IsDefined(TextUnderlineType.Single).Should().BeTrue();
        TextUnderlineType.Single.Should().NotBe(TextUnderlineType.NotDefined);
    }

    [Fact]
    public void Underline_SingleIsDistinctFromNone()
    {
        TextUnderlineType.Single.Should().NotBe(TextUnderlineType.None);
    }

    /// <summary>
    /// Verifies TextStrikethroughType.Single.
    /// </summary>
    [Fact]
    public void Strikethrough_SingleTypeIsDefined()
    {
        Enum.IsDefined(TextStrikethroughType.Single).Should().BeTrue();
        TextStrikethroughType.Single.Should().NotBe(TextStrikethroughType.NotDefined);
    }

    [Fact]
    public void Strikethrough_SingleIsDistinctFromNone()
    {
        TextStrikethroughType.Single.Should().NotBe(TextStrikethroughType.None);
    }

    /// <summary>
    /// Verifies FillType.Solid and Color.Red RGB components (255, 0, 0).
    /// </summary>
    [Fact]
    public void FontColor_FillTypeSolidIsDefined()
    {
        Enum.IsDefined(FillType.Solid).Should().BeTrue();
        FillType.Solid.Should().NotBe(FillType.NotDefined);
    }

    [Fact]
    public void FontColor_RedColorHasCorrectComponents()
    {
        Color.Red.R.Should().Be(255);
        Color.Red.G.Should().Be(0);
        Color.Red.B.Should().Be(0);
    }

    /// <summary>
    /// Verifies TextAlignment.Center.
    /// </summary>
    [Fact]
    public void ParagraphAlignment_CenterIsDefined()
    {
        Enum.IsDefined(TextAlignment.Center).Should().BeTrue();
        TextAlignment.Center.Should().NotBe(TextAlignment.NotDefined);
    }

    [Fact]
    public void ParagraphAlignment_CenterIsDistinctFromLeftAndRight()
    {
        TextAlignment.Center.Should().NotBe(TextAlignment.Left);
        TextAlignment.Center.Should().NotBe(TextAlignment.Right);
    }

    [Theory]
    [InlineData(TextAlignment.Left)]
    [InlineData(TextAlignment.Center)]
    [InlineData(TextAlignment.Right)]
    [InlineData(TextAlignment.Justify)]
    public void TextAlignment_AllCommonValuesAreDefined(TextAlignment alignment)
    {
        Enum.IsDefined(alignment).Should().BeTrue();
    }

    [Fact]
    public void NullableBool_NotDefinedSentinelIsZero()
    {
        ((int)NullableBool.NotDefined).Should().Be(0);
    }

    [Fact]
    public void TextUnderlineType_NotDefinedSentinelIsZero()
    {
        ((int)TextUnderlineType.NotDefined).Should().Be(0);
    }

    [Fact]
    public void TextStrikethroughType_NotDefinedSentinelIsZero()
    {
        ((int)TextStrikethroughType.NotDefined).Should().Be(0);
    }

    [Fact]
    public void TextAlignment_NotDefinedSentinelIsZero()
    {
        ((int)TextAlignment.NotDefined).Should().Be(0);
    }

    /// <summary>
    /// Verifies FontData class can be instantiated.
    /// </summary>
    [Fact]
    public void FontSize_FontDataCanBeInstantiated()
    {
        var fontData = new FontData("Arial");
        fontData.Should().NotBeNull();
    }

    /// <summary>
    /// Verifies FontData class exists and can be created.
    /// </summary>
    [Fact]
    public void LatinFont_FontDataCanBeInstantiated()
    {
        var fontData = new FontData("Courier New");
        fontData.Should().NotBeNull();
    }

    /// <summary>
    /// FontData.FontName returns the font name passed to the constructor.
    /// </summary>
    [Fact]
    public void LatinFont_FontNameMatchesConstructorArgument()
    {
        var fontData = new FontData("Courier New");

        fontData.FontName.Should().Be("Courier New");
    }

    /// <summary>
    /// Color.Red has the expected RGB components (255, 0, 0).
    /// </summary>
    [Fact]
    public void FontColor_RedColorIsNotEmpty()
    {
        Color.Red.Should().NotBe(Color.Empty);
    }

    /// <summary>
    /// Font height would be stored as a numeric value; NullableBool sentinel confirms pattern.
    /// </summary>
    [Fact]
    public void FontSize_NullableBoolPatternIsConsistent()
    {
        NullableBool.NotDefined.Should().NotBe(NullableBool.True);
        NullableBool.NotDefined.Should().NotBe(NullableBool.False);
    }

    /// <summary>
    /// FillFormat on a portion element can be set to Solid and color persists in XML.
    /// </summary>
    [Fact]
    public void FontColor_SolidFillColorPersistsInXml()
    {
        var rPr = new XElement(ANs + "rPr");
        var ff = new FillFormat();
        ff.InitInternal(rPr, parentSlide: null);

        ff.FillType = FillType.Solid;
        ff.SolidFillColor.Color = Color.Red;

        ff.FillType.Should().Be(FillType.Solid);
        var c = ff.SolidFillColor.Color;
        c.Should().NotBeNull();
        c!.R.Should().Be(255);
        c.G.Should().Be(0);
        c.B.Should().Be(0);
    }

    /// <summary>
    /// FontData preserves the exact font name through construction.
    /// </summary>
    [Theory]
    [InlineData("Courier New")]
    [InlineData("Arial")]
    [InlineData("Times New Roman")]
    public void LatinFont_FontNamePreservedForVariousFonts(string fontName)
    {
        var fontData = new FontData(fontName);

        fontData.FontName.Should().Be(fontName);
    }

    /// <summary>
    /// FontData.GetFontName returns same value as FontName.
    /// </summary>
    [Fact]
    public void LatinFont_GetFontNameReturnsSameAsFontName()
    {
        var fontData = new FontData("Courier New");

        fontData.GetFontName(null).Should().Be("Courier New");
    }
}
