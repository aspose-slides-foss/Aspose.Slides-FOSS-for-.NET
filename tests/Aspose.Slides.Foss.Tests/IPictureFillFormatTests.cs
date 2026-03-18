using System.Xml.Linq;
using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Verifies picture fill format properties on the <see cref="PictureFillFormat"/> class.
/// </summary>
public sealed class IPictureFillFormatTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static PictureFillFormat CreatePictureFillFormat(XElement? blipFill = null)
    {
        var el = blipFill ?? new XElement(ANs + "blipFill",
            new XElement(ANs + "blip"),
            new XElement(ANs + "stretch",
                new XElement(ANs + "fillRect")));
        var pff = new PictureFillFormat();
        pff.InitInternal(el, slidePart: null);
        return pff;
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void PictureFillMode_StretchCanBeSetAndRead()
    {
        var pff = CreatePictureFillFormat();
        pff.PictureFillMode = PictureFillMode.Stretch;

        pff.PictureFillMode.Should().Be(PictureFillMode.Stretch);
    }

    [Fact]
    public void PictureFillMode_TileCanBeSetAndRead()
    {
        var pff = CreatePictureFillFormat();
        pff.PictureFillMode = PictureFillMode.Tile;

        pff.PictureFillMode.Should().Be(PictureFillMode.Tile);
    }

    [Fact]
    public void PictureFillMode_ChangingModeReplacesElement()
    {
        var pff = CreatePictureFillFormat();
        pff.PictureFillMode = PictureFillMode.Stretch;
        pff.PictureFillMode.Should().Be(PictureFillMode.Stretch);

        pff.PictureFillMode = PictureFillMode.Tile;
        pff.PictureFillMode.Should().Be(PictureFillMode.Tile);
    }

    [Fact]
    public void Picture_IsNotNull()
    {
        var pff = CreatePictureFillFormat();
        pff.Picture.Should().NotBeNull();
    }

    [Fact]
    public void Dpi_DefaultIsZero()
    {
        var pff = CreatePictureFillFormat();
        pff.Dpi.Should().Be(0);
    }

    [Fact]
    public void Dpi_CanBeSetAndRead()
    {
        var pff = CreatePictureFillFormat();
        pff.Dpi = 300;

        pff.Dpi.Should().Be(300);
    }

    [Fact]
    public void CropLeft_DefaultIsZero()
    {
        var pff = CreatePictureFillFormat();
        pff.CropLeft.Should().Be(0f);
    }

    [Fact]
    public void CropLeft_CanBeSetAndRead()
    {
        var pff = CreatePictureFillFormat();
        pff.CropLeft = 10.5f;

        pff.CropLeft.Should().Be(10.5f);
    }

    [Theory]
    [InlineData(0f)]
    [InlineData(25.0f)]
    [InlineData(50.0f)]
    public void CropTop_CanBeSetAndRead(float value)
    {
        var pff = CreatePictureFillFormat();
        pff.CropTop = value;

        pff.CropTop.Should().Be(value);
    }

    [Fact]
    public void CropRight_CanBeSetAndRead()
    {
        var pff = CreatePictureFillFormat();
        pff.CropRight = 15f;

        pff.CropRight.Should().Be(15f);
    }

    [Fact]
    public void CropBottom_CanBeSetAndRead()
    {
        var pff = CreatePictureFillFormat();
        pff.CropBottom = 20f;

        pff.CropBottom.Should().Be(20f);
    }

    [Fact]
    public void StretchOffsetLeft_DefaultIsZero()
    {
        var pff = CreatePictureFillFormat();
        pff.StretchOffsetLeft.Should().Be(0f);
    }

    [Fact]
    public void StretchOffsetLeft_CanBeSetAndRead()
    {
        var pff = CreatePictureFillFormat();
        pff.StretchOffsetLeft = 5f;

        pff.StretchOffsetLeft.Should().Be(5f);
    }

    [Fact]
    public void StretchOffsetTop_CanBeSetAndRead()
    {
        var pff = CreatePictureFillFormat();
        pff.StretchOffsetTop = 10f;

        pff.StretchOffsetTop.Should().Be(10f);
    }

    [Fact]
    public void StretchOffsetRight_CanBeSetAndRead()
    {
        var pff = CreatePictureFillFormat();
        pff.StretchOffsetRight = 15f;

        pff.StretchOffsetRight.Should().Be(15f);
    }

    [Fact]
    public void StretchOffsetBottom_CanBeSetAndRead()
    {
        var pff = CreatePictureFillFormat();
        pff.StretchOffsetBottom = 20f;

        pff.StretchOffsetBottom.Should().Be(20f);
    }

    [Fact]
    public void TileOffsetX_DefaultIsZero()
    {
        var pff = CreatePictureFillFormat();
        pff.TileOffsetX.Should().Be(0f);
    }

    [Fact]
    public void TileOffsetX_CanBeSetAndRead()
    {
        var pff = CreatePictureFillFormat();
        pff.TileOffsetX = 100f;

        pff.TileOffsetX.Should().Be(100f);
    }

    [Fact]
    public void TileOffsetY_CanBeSetAndRead()
    {
        var pff = CreatePictureFillFormat();
        pff.TileOffsetY = -50f;

        pff.TileOffsetY.Should().Be(-50f);
    }

    [Fact]
    public void TileScaleX_CanBeSetAndRead()
    {
        var pff = CreatePictureFillFormat();
        pff.TileScaleX = 150f;

        pff.TileScaleX.Should().Be(150f);
    }

    [Fact]
    public void TileScaleY_CanBeSetAndRead()
    {
        var pff = CreatePictureFillFormat();
        pff.TileScaleY = 75f;

        pff.TileScaleY.Should().Be(75f);
    }

    [Fact]
    public void TileAlignment_DefaultIsNotDefined()
    {
        var pff = CreatePictureFillFormat();
        pff.TileAlignment.Should().Be(RectangleAlignment.NotDefined);
    }

    [Theory]
    [InlineData(RectangleAlignment.TopLeft)]
    [InlineData(RectangleAlignment.Center)]
    [InlineData(RectangleAlignment.BottomRight)]
    public void TileAlignment_CanBeSetAndRead(RectangleAlignment alignment)
    {
        var pff = CreatePictureFillFormat();
        pff.TileAlignment = alignment;

        pff.TileAlignment.Should().Be(alignment);
    }

    [Fact]
    public void TileFlip_DefaultIsNotDefined()
    {
        var pff = CreatePictureFillFormat();
        pff.TileFlip.Should().Be(TileFlip.NotDefined);
    }

    [Theory]
    [InlineData(TileFlip.NoFlip)]
    [InlineData(TileFlip.FlipX)]
    [InlineData(TileFlip.FlipY)]
    [InlineData(TileFlip.FlipBoth)]
    public void TileFlip_CanBeSetAndRead(TileFlip flip)
    {
        var pff = CreatePictureFillFormat();
        pff.TileFlip = flip;

        pff.TileFlip.Should().Be(flip);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void PictureFill_EnumValuesAreDefined()
    {
        Enum.IsDefined(FillType.Picture).Should().BeTrue();
        Enum.IsDefined(PictureFillMode.Stretch).Should().BeTrue();
    }

    /// <summary>
    /// through FillFormat works end-to-end.
    /// </summary>
    [Fact]
    public void PictureFill_FillFormatCreatesPictureFillFormat()
    {
        var parent = new XElement(ANs + "spPr");
        var ff = new FillFormat();
        ff.InitInternal(parent, parentSlide: null);
        ff.FillType = FillType.Picture;

        ff.FillType.Should().Be(FillType.Picture);
        var pff = ff.PictureFillFormat;
        pff.Should().NotBeNull();
        pff.PictureFillMode = PictureFillMode.Stretch;
        pff.PictureFillMode.Should().Be(PictureFillMode.Stretch);
    }
}
