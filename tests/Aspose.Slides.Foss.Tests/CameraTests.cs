using System.Xml.Linq;
using Aspose.Slides.Foss;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Unit tests for the Camera class.
/// </summary>
public sealed class CameraTests
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static Camera CreateCamera(XElement? cameraElement = null)
    {
        var scene3d = new XElement(ANs + "scene3d");
        if (cameraElement is not null)
            scene3d.Add(cameraElement);
        else
            scene3d.Add(new XElement(ANs + "camera", new XAttribute("prst", "orthographicFront")));

        var spPr = new XElement(ANs + "spPr", scene3d);
        var tdf = new ThreeDFormat();
        tdf.InitInternal(spPr, null);
        return (Camera)tdf.Camera;
    }

    [Fact]
    public void CameraType_DefaultIsOrthographicFront()
    {
        var cam = CreateCamera();
        cam.CameraType.Should().Be(CameraPresetType.OrthographicFront);
    }

    [Theory]
    [InlineData(CameraPresetType.PerspectiveAbove)]
    [InlineData(CameraPresetType.PerspectiveFront)]
    [InlineData(CameraPresetType.IsometricLeftDown)]
    [InlineData(CameraPresetType.ObliqueTopLeft)]
    [InlineData(CameraPresetType.LegacyPerspectiveFront)]
    public void CameraType_SetAndReadBack(CameraPresetType preset)
    {
        var cam = CreateCamera();
        cam.CameraType = preset;
        cam.CameraType.Should().Be(preset);
    }

    [Fact]
    public void CameraType_SetNotDefined_RemovesPresetAttribute()
    {
        var cam = CreateCamera();
        cam.CameraType = CameraPresetType.NotDefined;
        cam.CameraType.Should().Be(CameraPresetType.NotDefined);
    }

    [Fact]
    public void FieldOfViewAngle_DefaultIsZero()
    {
        var cam = CreateCamera();
        cam.FieldOfViewAngle.Should().Be(0f);
    }

    [Fact]
    public void FieldOfViewAngle_SetAndReadBack()
    {
        var cam = CreateCamera();
        cam.FieldOfViewAngle = 90f;
        cam.FieldOfViewAngle.Should().BeApproximately(90f, 0.01f);
    }

    [Fact]
    public void Zoom_DefaultIs100()
    {
        var cam = CreateCamera();
        cam.Zoom.Should().Be(100f);
    }

    [Fact]
    public void Zoom_SetAndReadBack()
    {
        var cam = CreateCamera();
        cam.Zoom = 150f;
        cam.Zoom.Should().BeApproximately(150f, 0.01f);
    }

    [Fact]
    public void SetRotation_GetRotation_RoundTrip()
    {
        var cam = CreateCamera();
        cam.SetRotation(45f, 90f, 180f);
        var rot = cam.GetRotation();
        rot.Should().HaveCount(3);
        rot[0].Should().BeApproximately(45f, 0.01f);
        rot[1].Should().BeApproximately(90f, 0.01f);
        rot[2].Should().BeApproximately(180f, 0.01f);
    }

    [Fact]
    public void GetRotation_DefaultIsAllZeros()
    {
        var camElement = new XElement(ANs + "camera", new XAttribute("prst", "orthographicFront"));
        var cam = CreateCamera(camElement);
        var rot = cam.GetRotation();
        rot.Should().Equal([0f, 0f, 0f]);
    }

    /// <summary>
    /// </summary>
    [Fact]
    public void CameraType_PerspectiveAbove_PersistsViaThreeDFormat()
    {
        var spPr = XElement.Parse("<spPr xmlns='http://schemas.openxmlformats.org/drawingml/2006/main'/>");
        var tdf = new ThreeDFormat();
        tdf.InitInternal(spPr, null);

        tdf.Camera.CameraType = CameraPresetType.PerspectiveAbove;

        var tdf2 = new ThreeDFormat();
        tdf2.InitInternal(spPr, null);
        tdf2.Camera.CameraType.Should().Be(CameraPresetType.PerspectiveAbove);
    }
}
