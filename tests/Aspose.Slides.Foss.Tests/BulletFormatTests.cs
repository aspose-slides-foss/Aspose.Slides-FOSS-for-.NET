using System.Xml.Linq;
using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests for BulletFormat: type, char, font, height, color, numbered styles,
/// hard color/font flags, and schema-ordered element insertion.
/// </summary>
public sealed class BulletFormatTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static BulletFormat CreateBulletFormat(XElement? pPr = null)
    {
        pPr ??= new XElement(ANs + "pPr");
        var bf = new BulletFormat();
        bf.InitInternal(pPr, null, null);
        return bf;
    }

    [Fact]
    public void Type_DefaultIsNotDefined()
    {
        var bf = CreateBulletFormat();
        bf.Type.Should().Be(BulletType.NotDefined);
    }

    [Theory]
    [InlineData(BulletType.None)]
    [InlineData(BulletType.Symbol)]
    [InlineData(BulletType.Numbered)]
    [InlineData(BulletType.Picture)]
    public void Type_SetAndGet_RoundTrips(BulletType type)
    {
        var bf = CreateBulletFormat();
        bf.Type = type;
        bf.Type.Should().Be(type);
    }

    [Fact]
    public void Type_SetSymbol_CreatesDefaultBulletChar()
    {
        var pPr = new XElement(ANs + "pPr");
        var bf = CreateBulletFormat(pPr);
        bf.Type = BulletType.Symbol;

        var buChar = pPr.Element(ANs + "buChar");
        buChar.Should().NotBeNull();
        buChar!.Attribute("char")!.Value.Should().Be("\u2022");
    }

    [Fact]
    public void Type_SetNumbered_CreatesDefaultArabicPeriod()
    {
        var pPr = new XElement(ANs + "pPr");
        var bf = CreateBulletFormat(pPr);
        bf.Type = BulletType.Numbered;

        var buAutoNum = pPr.Element(ANs + "buAutoNum");
        buAutoNum.Should().NotBeNull();
        buAutoNum!.Attribute("type")!.Value.Should().Be("arabicPeriod");
    }

    [Fact]
    public void Type_SetReplacesPreviousBulletType()
    {
        var bf = CreateBulletFormat();
        bf.Type = BulletType.Symbol;
        bf.Type = BulletType.Numbered;
        bf.Type.Should().Be(BulletType.Numbered);
    }

    [Fact]
    public void Char_EmptyByDefault()
    {
        var bf = CreateBulletFormat();
        bf.Char.Should().BeEmpty();
    }

    [Fact]
    public void Char_GetReturnsCharAttribute()
    {
        var pPr = new XElement(ANs + "pPr",
            new XElement(ANs + "buChar", new XAttribute("char", "-")));
        var bf = CreateBulletFormat(pPr);
        bf.Char.Should().Be("-");
    }

    [Fact]
    public void Char_SetUpdatesExistingBuChar()
    {
        var pPr = new XElement(ANs + "pPr",
            new XElement(ANs + "buChar", new XAttribute("char", "-")));
        var bf = CreateBulletFormat(pPr);
        bf.Char = ">";
        bf.Char.Should().Be(">");
    }

    [Fact]
    public void Char_SetCreatesBuCharWhenNoneExists()
    {
        var bf = CreateBulletFormat();
        bf.Char = "*";
        bf.Char.Should().Be("*");
        bf.Type.Should().Be(BulletType.Symbol);
    }

    [Fact]
    public void Char_SetRemovesOtherBulletTypes()
    {
        var bf = CreateBulletFormat();
        bf.Type = BulletType.Numbered;
        bf.Char = ">";
        bf.Type.Should().Be(BulletType.Symbol);
    }

    [Fact]
    public void Font_NullByDefault()
    {
        var bf = CreateBulletFormat();
        bf.Font.Should().BeNull();
    }

    [Fact]
    public void Font_SetAndGet_RoundTrips()
    {
        var bf = CreateBulletFormat();
        bf.Font = new FontData("Arial");
        bf.Font.Should().NotBeNull();
        bf.Font!.FontName.Should().Be("Arial");
    }

    [Fact]
    public void Font_SetNull_RemovesBuFont()
    {
        var bf = CreateBulletFormat();
        bf.Font = new FontData("Arial");
        bf.Font = null;
        bf.Font.Should().BeNull();
    }

    [Fact]
    public void Font_SetOverwritesExisting()
    {
        var bf = CreateBulletFormat();
        bf.Font = new FontData("Arial");
        bf.Font = new FontData("Courier New");
        bf.Font!.FontName.Should().Be("Courier New");
    }

    [Fact]
    public void Height_NaNByDefault()
    {
        var bf = CreateBulletFormat();
        float.IsNaN(bf.Height).Should().BeTrue();
    }

    [Fact]
    public void Height_ReadsBuSzPct()
    {
        var pPr = new XElement(ANs + "pPr",
            new XElement(ANs + "buSzPct", new XAttribute("val", "75000")));
        var bf = CreateBulletFormat(pPr);
        bf.Height.Should().Be(75f);
    }

    [Fact]
    public void Height_ReadsBuSzPts()
    {
        var pPr = new XElement(ANs + "pPr",
            new XElement(ANs + "buSzPts", new XAttribute("val", "1200")));
        var bf = CreateBulletFormat(pPr);
        bf.Height.Should().Be(12f);
    }

    [Fact]
    public void Height_PctTakesPrecedenceOverPts()
    {
        var pPr = new XElement(ANs + "pPr",
            new XElement(ANs + "buSzPct", new XAttribute("val", "50000")),
            new XElement(ANs + "buSzPts", new XAttribute("val", "1200")));
        var bf = CreateBulletFormat(pPr);
        bf.Height.Should().Be(50f);
    }

    [Fact]
    public void Height_SetAndGet_RoundTrips()
    {
        var bf = CreateBulletFormat();
        bf.Height = 120f;
        bf.Height.Should().Be(120f);
    }

    [Fact]
    public void Height_SetNaN_RemovesElement()
    {
        var bf = CreateBulletFormat();
        bf.Height = 100f;
        bf.Height = float.NaN;
        float.IsNaN(bf.Height).Should().BeTrue();
    }

    [Fact]
    public void NumberedBulletStartWith_DefaultIs1()
    {
        var bf = CreateBulletFormat();
        bf.NumberedBulletStartWith.Should().Be(1);
    }

    [Fact]
    public void NumberedBulletStartWith_ReadsStartAt()
    {
        var pPr = new XElement(ANs + "pPr",
            new XElement(ANs + "buAutoNum",
                new XAttribute("type", "arabicPeriod"),
                new XAttribute("startAt", "5")));
        var bf = CreateBulletFormat(pPr);
        bf.NumberedBulletStartWith.Should().Be(5);
    }

    [Fact]
    public void NumberedBulletStartWith_SetUpdatesAttribute()
    {
        var bf = CreateBulletFormat();
        bf.Type = BulletType.Numbered;
        bf.NumberedBulletStartWith = 3;
        bf.NumberedBulletStartWith.Should().Be(3);
    }

    [Fact]
    public void NumberedBulletStyle_DefaultIsNotDefined()
    {
        var bf = CreateBulletFormat();
        bf.NumberedBulletStyle.Should().Be(NumberedBulletStyle.NotDefined);
    }

    [Theory]
    [InlineData("arabicPeriod", NumberedBulletStyle.BulletArabicPeriod)]
    [InlineData("romanUcPeriod", NumberedBulletStyle.BulletRomanUCPeriod)]
    [InlineData("alphaLcParenR", NumberedBulletStyle.BulletAlphaLCParenRight)]
    [InlineData("circleNumDbPlain", NumberedBulletStyle.BulletCircleNumDBPlain)]
    [InlineData("ea1ChsPeriod", NumberedBulletStyle.BulletSimpChinPeriod)]
    [InlineData("thaiAlphaPeriod", NumberedBulletStyle.BulletThaiAlphaPeriod)]
    [InlineData("hindiNumPeriod", NumberedBulletStyle.BulletHindiNumPeriod)]
    [InlineData("arabic1Minus", NumberedBulletStyle.BulletArabicAlphaDash)]
    [InlineData("hebrew2Minus", NumberedBulletStyle.BulletHebrewDash)]
    public void NumberedBulletStyle_ReadsFromXml(string xmlType, NumberedBulletStyle expected)
    {
        var pPr = new XElement(ANs + "pPr",
            new XElement(ANs + "buAutoNum", new XAttribute("type", xmlType)));
        var bf = CreateBulletFormat(pPr);
        bf.NumberedBulletStyle.Should().Be(expected);
    }

    [Fact]
    public void NumberedBulletStyle_SetAndGet_RoundTrips()
    {
        var bf = CreateBulletFormat();
        bf.Type = BulletType.Numbered;
        bf.NumberedBulletStyle = NumberedBulletStyle.BulletRomanUCPeriod;
        bf.NumberedBulletStyle.Should().Be(NumberedBulletStyle.BulletRomanUCPeriod);
    }

    [Fact]
    public void IsBulletHardColor_DefaultIsNotDefined()
    {
        var bf = CreateBulletFormat();
        bf.IsBulletHardColor.Should().Be(NullableBool.NotDefined);
    }

    [Fact]
    public void IsBulletHardColor_TrueWhenBuClrPresent()
    {
        var pPr = new XElement(ANs + "pPr",
            new XElement(ANs + "buClr"));
        var bf = CreateBulletFormat(pPr);
        bf.IsBulletHardColor.Should().Be(NullableBool.True);
    }

    [Fact]
    public void IsBulletHardColor_FalseWhenBuClrTxPresent()
    {
        var pPr = new XElement(ANs + "pPr",
            new XElement(ANs + "buClrTx"));
        var bf = CreateBulletFormat(pPr);
        bf.IsBulletHardColor.Should().Be(NullableBool.False);
    }

    [Theory]
    [InlineData(NullableBool.True)]
    [InlineData(NullableBool.False)]
    public void IsBulletHardColor_SetAndGet_RoundTrips(NullableBool value)
    {
        var bf = CreateBulletFormat();
        bf.IsBulletHardColor = value;
        bf.IsBulletHardColor.Should().Be(value);
    }

    [Fact]
    public void IsBulletHardFont_DefaultIsNotDefined()
    {
        var bf = CreateBulletFormat();
        bf.IsBulletHardFont.Should().Be(NullableBool.NotDefined);
    }

    [Fact]
    public void IsBulletHardFont_TrueWhenBuFontPresent()
    {
        var pPr = new XElement(ANs + "pPr",
            new XElement(ANs + "buFont"));
        var bf = CreateBulletFormat(pPr);
        bf.IsBulletHardFont.Should().Be(NullableBool.True);
    }

    [Fact]
    public void IsBulletHardFont_FalseWhenBuFontTxPresent()
    {
        var pPr = new XElement(ANs + "pPr",
            new XElement(ANs + "buFontTx"));
        var bf = CreateBulletFormat(pPr);
        bf.IsBulletHardFont.Should().Be(NullableBool.False);
    }

    [Theory]
    [InlineData(NullableBool.True)]
    [InlineData(NullableBool.False)]
    public void IsBulletHardFont_SetAndGet_RoundTrips(NullableBool value)
    {
        var bf = CreateBulletFormat();
        bf.IsBulletHardFont = value;
        bf.IsBulletHardFont.Should().Be(value);
    }

    [Fact]
    public void Picture_ReturnsNull()
    {
        var bf = CreateBulletFormat();
        bf.Picture.Should().BeNull();
    }

    [Fact]
    public void Color_ReturnsColorFormat()
    {
        var bf = CreateBulletFormat();
        bf.Color.Should().NotBeNull();
    }
}
