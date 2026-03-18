using Aspose.Slides.Foss.Drawing;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests for Size and SizeF: construction, dimension access, mutability, equality.
/// test_shapes (shape frame dimensions), and test_table (row heights, column widths).
/// </summary>
public sealed class SizeTests
{
    [Fact]
    public void Size_Constructor_SetsDimensions()
    {
        var size = new Size(300, 250);

        size.Width.Should().Be(300);
        size.Height.Should().Be(250);
    }

    [Fact]
    public void Size_Constructor_Defaults_AreZero()
    {
        var size = new Size();

        size.Width.Should().Be(0);
        size.Height.Should().Be(0);
    }

    [Fact]
    public void Size_Dimensions_AreMutable()
    {
        var size = new Size(100, 200);

        size.Width = 400;
        size.Height = 500;

        size.Width.Should().Be(400);
        size.Height.Should().Be(500);
    }

    [Fact]
    public void Size_PositiveDimensions()
    {
        // From test_notes_size: notes size has positive width and height
        var size = new Size(720, 540);

        size.Width.Should().BePositive();
        size.Height.Should().BePositive();
    }

    [Fact]
    public void Size_Equals_SameDimensions_ReturnsTrue()
    {
        var a = new Size(200, 100);
        var b = new Size(200, 100);

        a.Should().Be(b);
    }

    [Fact]
    public void Size_Equals_DifferentDimensions_ReturnsFalse()
    {
        var a = new Size(200, 100);
        var b = new Size(200, 101);

        a.Should().NotBe(b);
    }

    [Fact]
    public void Size_GetHashCode_EqualSizes_SameHash()
    {
        var a = new Size(300, 250);
        var b = new Size(300, 250);

        a.GetHashCode().Should().Be(b.GetHashCode());
    }

    [Theory]
    [InlineData(100, 60)]
    [InlineData(200, 100)]
    [InlineData(300, 250)]
    public void Size_Constructor_RoundTrips(int width, int height)
    {
        // From test_shapes: shape frame width/height values
        var size = new Size(width, height);

        size.Width.Should().Be(width);
        size.Height.Should().Be(height);
    }

    [Fact]
    public void SizeF_Constructor_SetsDimensions()
    {
        var size = new SizeF(300.5f, 250.5f);

        size.Width.Should().Be(300.5f);
        size.Height.Should().Be(250.5f);
    }

    [Fact]
    public void SizeF_Constructor_Defaults_AreZero()
    {
        var size = new SizeF();

        size.Width.Should().Be(0f);
        size.Height.Should().Be(0f);
    }

    [Fact]
    public void SizeF_Dimensions_AreMutable()
    {
        var size = new SizeF(1.0f, 2.0f);

        size.Width = 10.5f;
        size.Height = 20.5f;

        size.Width.Should().Be(10.5f);
        size.Height.Should().Be(20.5f);
    }

    [Fact]
    public void SizeF_Equals_SameDimensions_ReturnsTrue()
    {
        var a = new SizeF(100.5f, 200.5f);
        var b = new SizeF(100.5f, 200.5f);

        a.Should().Be(b);
    }

    [Fact]
    public void SizeF_Equals_DifferentDimensions_ReturnsFalse()
    {
        var a = new SizeF(100.5f, 200.5f);
        var b = new SizeF(100.5f, 200.6f);

        a.Should().NotBe(b);
    }

    [Fact]
    public void SizeF_GetHashCode_EqualSizes_SameHash()
    {
        var a = new SizeF(10.5f, 20.5f);
        var b = new SizeF(10.5f, 20.5f);

        a.GetHashCode().Should().Be(b.GetHashCode());
    }
}
