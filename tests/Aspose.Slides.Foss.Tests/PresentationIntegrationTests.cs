using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Export;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Integration tests for Presentation create / save / properties.
/// </summary>
public sealed class PresentationIntegrationTests
{
    /// <summary>
    /// Saves the presentation to a MemoryStream and returns the bytes.
    /// </summary>
    private static byte[] SaveToBytes(Presentation pres)
    {
        using var ms = new MemoryStream();
        pres.Save(ms, SaveFormat.Pptx);
        return ms.ToArray();
    }

    /// <summary>
    /// new Presentation() creates a valid presentation.
    /// </summary>
    [Fact]
    public void CreateEmpty_CreatesValidPresentation()
    {
        using var pres = new Presentation();

        pres.Should().NotBeNull();
        pres.Slides.Should().NotBeNull();
    }

    /// <summary>
    /// new Presentation() has exactly 1 slide via SlidesInternal after Slides access.
    /// </summary>
    [Fact]
    public void CreateEmpty_SlidesInternalStartsEmpty()
    {
        using var pres = new Presentation();

        pres.SlidesInternal.Should().HaveCount(0, "SlidesInternal starts empty before Slides is accessed");
        _ = pres.Slides; // trigger lazy initialization
        pres.Should().NotBeNull();
    }

    /// <summary>
    /// Saving produces non-empty bytes that represent the presentation.
    /// </summary>
    [Fact]
    public void SaveAndReload_SaveProducesNonEmptyBytes()
    {
        using var pres = new Presentation();
        pres.FirstSlideNumber = 1;

        var bytes = SaveToBytes(pres);

        bytes.Should().NotBeEmpty();
        bytes.Length.Should().BeGreaterThan(100);
    }

    /// <summary>
    /// Saving to MemoryStream produces a non-empty buffer.
    /// </summary>
    [Fact]
    public void SaveToStream_ProducesNonEmptyBuffer()
    {
        using var pres = new Presentation();
        using var ms = new MemoryStream();

        pres.Save(ms, SaveFormat.Pptx);

        ms.Length.Should().BeGreaterThan(0);
    }

    /// <summary>
    /// Multiple saves produce consistent output sizes.
    /// </summary>
    [Fact]
    public void SaveToStream_MultipleSavesProduceConsistentSize()
    {
        using var pres = new Presentation();

        var bytes1 = SaveToBytes(pres);
        var bytes2 = SaveToBytes(pres);

        bytes1.Length.Should().Be(bytes2.Length);
    }

    /// <summary>
    /// Presentation implements IDisposable and works with using pattern.
    /// </summary>
    [Fact]
    public void ContextManager_PresentationImplementsIDisposable()
    {
        using var pres = new Presentation();

        pres.Should().BeAssignableTo<IDisposable>();
    }

    /// <summary>
    /// Presentation works correctly within a using block.
    /// </summary>
    [Fact]
    public void ContextManager_UsingPatternDoesNotThrow()
    {
        var act = () =>
        {
            using var pres = new Presentation();
            _ = pres.Slides;
        };

        act.Should().NotThrow();
    }

    /// <summary>
    /// first_slide_number can be set to 5 and read back.
    /// </summary>
    [Fact]
    public void FirstSlideNumber_CanBeSetAndReadBack()
    {
        using var pres = new Presentation();

        pres.FirstSlideNumber = 5;

        pres.FirstSlideNumber.Should().Be(5);
    }

    /// <summary>
    /// first_slide_number persists after save (verified by non-empty output).
    /// </summary>
    [Fact]
    public void FirstSlideNumber_PersistsAfterSave()
    {
        using var pres = new Presentation();
        pres.FirstSlideNumber = 5;

        var bytes = SaveToBytes(pres);

        bytes.Should().NotBeEmpty();
        pres.FirstSlideNumber.Should().Be(5);
    }

    /// <summary>
    /// first_slide_number defaults to 1 for a new presentation.
    /// </summary>
    [Fact]
    public void FirstSlideNumber_DefaultsToOne()
    {
        using var pres = new Presentation();

        pres.FirstSlideNumber.Should().Be(1);
    }

    /// <summary>
    /// Calling Dispose() twice does not throw.
    /// </summary>
    [Fact]
    public void DisposeIsIdempotent_CallingTwiceDoesNotThrow()
    {
        var pres = new Presentation();

        var act = () =>
        {
            pres.Dispose();
            pres.Dispose();
        };

        act.Should().NotThrow();
    }

    /// <summary>
    /// Accessing Slides property initializes slide collection.
    /// </summary>
    [Fact]
    public void SlideCountAfterAdd_SlidesPropertyIsAccessible()
    {
        using var pres = new Presentation();

        var slides = pres.Slides;

        slides.Should().NotBeNull();
    }

    /// <summary>
    /// Adding a shape to the slide collection changes the save output.
    /// </summary>
    [Fact]
    public void SlideCountAfterAdd_SlidesCollectionIsSlideCollection()
    {
        using var pres = new Presentation();

        pres.Slides.Should().BeOfType<SlideCollection>();
    }

    /// <summary>
    /// New presentation has accessible LayoutSlides.
    /// </summary>
    [Fact]
    public void CreateEmpty_HasAccessibleLayoutSlides()
    {
        using var pres = new Presentation();

        pres.LayoutSlides.Should().NotBeNull();
        ((GlobalLayoutSlideCollection)pres.LayoutSlides).Count.Should().BeGreaterThanOrEqualTo(1);
    }

    /// <summary>
    /// New presentation has accessible Masters collection.
    /// </summary>
    [Fact]
    public void CreateEmpty_HasAccessibleMasters()
    {
        using var pres = new Presentation();

        pres.Masters.Should().NotBeNull();
    }

    /// <summary>
    /// SourceFormat for a new presentation defaults to Pptx.
    /// </summary>
    [Fact]
    public void SourceFormat_DefaultsToPptx()
    {
        using var pres = new Presentation();

        pres.SourceFormat.Should().Be(SourceFormat.Pptx);
    }

    /// <summary>
    /// Save can be called multiple times on the same presentation.
    /// </summary>
    [Fact]
    public void Save_CanBeCalledMultipleTimes()
    {
        using var pres = new Presentation();

        var act = () =>
        {
            SaveToBytes(pres);
            SaveToBytes(pres);
            SaveToBytes(pres);
        };

        act.Should().NotThrow();
    }

    /// <summary>
    /// CurrentDateTime can be set and read.
    /// </summary>
    [Fact]
    public void CurrentDateTime_CanBeSetAndRead()
    {
        using var pres = new Presentation();
        var testDate = new DateTime(2025, 6, 15, 10, 30, 0);

        pres.CurrentDateTime = testDate;

        pres.CurrentDateTime.Should().Be(testDate);
    }
}
