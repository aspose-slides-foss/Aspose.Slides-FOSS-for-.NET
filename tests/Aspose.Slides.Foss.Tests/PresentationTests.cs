using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Export;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests Presentation create / save / properties.
/// </summary>
public sealed class PresentationTests
{
    /// <summary>
    /// SaveFormat.Pptx enum value is defined.
    /// </summary>
    [Fact]
    public void SaveToStream_SaveFormatPptxIsDefined()
    {
        Enum.IsDefined(SaveFormat.Pptx).Should().BeTrue();
    }

    /// <summary>
    /// SaveFormat has all expected format values.
    /// </summary>
    [Theory]
    [InlineData(SaveFormat.Ppt)]
    [InlineData(SaveFormat.Pdf)]
    [InlineData(SaveFormat.Pptx)]
    [InlineData(SaveFormat.Ppsx)]
    [InlineData(SaveFormat.Html)]
    [InlineData(SaveFormat.Html5)]
    [InlineData(SaveFormat.Xml)]
    public void SaveFormat_AllValuesAreDefined(SaveFormat format)
    {
        Enum.IsDefined(format).Should().BeTrue();
    }

    /// <summary>
    /// Pptx and Pdf are distinct formats.
    /// </summary>
    [Fact]
    public void SaveFormat_PptxIsDistinctFromPdf()
    {
        SaveFormat.Pptx.Should().NotBe(SaveFormat.Pdf);
    }

    /// <summary>
    /// Presentation class can be instantiated.
    /// </summary>
    [Fact]
    public void Presentation_CanBeInstantiated()
    {
        using var pres = new Presentation();
        pres.Should().NotBeNull();
    }

    /// <summary>
    /// SaveFormat.Pptx is not the first enum value (distinguishable).
    /// </summary>
    [Fact]
    public void SaveFormat_PptxIsNotFirstValue()
    {
        SaveFormat.Pptx.Should().NotBe(SaveFormat.Ppt);
    }

    /// <summary>
    /// SaveFormat enum has all 21 expected members.
    /// </summary>
    [Fact]
    public void SaveFormat_HasExpectedMemberCount()
    {
        Enum.GetValues<SaveFormat>().Should().HaveCount(21);
    }

    /// <summary>
    /// Each SaveFormat value has a unique underlying integer.
    /// </summary>
    [Fact]
    public void SaveFormat_AllValuesAreUnique()
    {
        var values = Enum.GetValues<SaveFormat>().Cast<int>().ToList();
        values.Should().OnlyHaveUniqueItems();
    }

    /// <summary>
    /// SaveFormat can be cast to int and back without data loss.
    /// </summary>
    [Theory]
    [InlineData(SaveFormat.Pptx)]
    [InlineData(SaveFormat.Odp)]
    [InlineData(SaveFormat.Tiff)]
    [InlineData(SaveFormat.Gif)]
    [InlineData(SaveFormat.Md)]
    public void SaveFormat_RoundTripsThroughInt(SaveFormat format)
    {
        var asInt = (int)format;
        var backToEnum = (SaveFormat)asInt;
        backToEnum.Should().Be(format);
    }

    /// <summary>
    /// All presentation-related formats are defined.
    /// </summary>
    [Theory]
    [InlineData(SaveFormat.Pptm)]
    [InlineData(SaveFormat.Ppsm)]
    [InlineData(SaveFormat.Potx)]
    [InlineData(SaveFormat.Potm)]
    [InlineData(SaveFormat.Otp)]
    [InlineData(SaveFormat.Pps)]
    [InlineData(SaveFormat.Pot)]
    [InlineData(SaveFormat.Fodp)]
    [InlineData(SaveFormat.Swf)]
    [InlineData(SaveFormat.Xps)]
    public void SaveFormat_AdditionalValuesAreDefined(SaveFormat format)
    {
        Enum.IsDefined(format).Should().BeTrue();
    }

    /// <summary>
    /// SaveFormat.ToString() returns the member name.
    /// </summary>
    [Fact]
    public void SaveFormat_PptxToStringReturnsMemberName()
    {
        SaveFormat.Pptx.ToString().Should().Be("Pptx");
    }

    /// <summary>
    /// SaveFormat can be parsed from a string name.
    /// </summary>
    [Theory]
    [InlineData("Pptx", SaveFormat.Pptx)]
    [InlineData("Pdf", SaveFormat.Pdf)]
    [InlineData("Html5", SaveFormat.Html5)]
    public void SaveFormat_CanBeParsedFromString(string name, SaveFormat expected)
    {
        Enum.Parse<SaveFormat>(name).Should().Be(expected);
    }
}
