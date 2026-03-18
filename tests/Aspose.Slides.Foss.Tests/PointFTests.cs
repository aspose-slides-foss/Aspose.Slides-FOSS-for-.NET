using Aspose.Slides.Foss.Drawing;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests for PointF: construction, coordinate access, mutability, equality.
/// test_connector (connector endpoints), and test_shapes (shape frame positions).
/// </summary>
public sealed class PointFTests
{
    [Fact]
    public void Constructor_SetsCoordinates()
    {
        // From test_add_comment: PointF(2.0, 3.0)
        var point = new PointF(2.0f, 3.0f);

        point.X.Should().Be(2.0f);
        point.Y.Should().Be(3.0f);
    }

    [Fact]
    public void Constructor_Defaults_AreZero()
    {
        var point = new PointF();

        point.X.Should().Be(0f);
        point.Y.Should().Be(0f);
    }

    [Fact]
    public void Constructor_VariousPositions()
    {
        // From test_get_slide_comments: multiple comments at different positions
        var p1 = new PointF(1, 1);
        var p2 = new PointF(2, 2);

        p1.X.Should().Be(1f);
        p1.Y.Should().Be(1f);
        p2.X.Should().Be(2f);
        p2.Y.Should().Be(2f);
    }

    [Fact]
    public void Coordinates_AreMutable()
    {
        var point = new PointF(1.0f, 1.0f);

        point.X = 5.0f;
        point.Y = 7.0f;

        point.X.Should().Be(5.0f);
        point.Y.Should().Be(7.0f);
    }

    [Fact]
    public void Equals_SameCoordinates_ReturnsTrue()
    {
        var a = new PointF(2.0f, 3.0f);
        var b = new PointF(2.0f, 3.0f);

        a.Should().Be(b);
    }

    [Fact]
    public void Equals_DifferentCoordinates_ReturnsFalse()
    {
        var a = new PointF(1.0f, 1.0f);
        var b = new PointF(2.0f, 2.0f);

        a.Should().NotBe(b);
    }

    [Fact]
    public void GetHashCode_EqualPoints_SameHash()
    {
        var a = new PointF(3.5f, 7.5f);
        var b = new PointF(3.5f, 7.5f);

        a.GetHashCode().Should().Be(b.GetHashCode());
    }

    [Theory]
    [InlineData(1f, 1f)]
    [InlineData(2f, 3f)]
    [InlineData(0f, 0f)]
    [InlineData(100.5f, 200.5f)]
    public void Constructor_RoundTrips_AllValues(float x, float y)
    {
        // From test_comments: various PointF values used for comment positions
        var point = new PointF(x, y);

        point.X.Should().Be(x);
        point.Y.Should().Be(y);
    }
}
