using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Internal;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests NotesSlideHeaderFooterManager visibility and text operations,
/// plus cast properties to base interfaces.
/// </summary>
public sealed class NotesSlideHeaderFooterTests
{
    /// <summary>
    /// Creates a NotesSlideHeaderFooterManager backed by a fresh NotesSlidePart.
    /// </summary>
    private static NotesSlideHeaderFooterManager CreateManager()
    {
        var notesPart = new NotesSlidePart();
        var manager = new NotesSlideHeaderFooterManager();
        manager.InitInternal(notesPart);
        return manager;
    }

    /// <summary>
    /// Footer visibility can be set to true and is reflected in IsFooterVisible.
    /// </summary>
    [Fact]
    public void SetFooterVisibility_True_MakesFooterVisible()
    {
        var hfm = CreateManager();

        hfm.IsFooterVisible.Should().BeFalse();

        hfm.SetFooterVisibility(true);

        hfm.IsFooterVisible.Should().BeTrue();
    }

    /// <summary>
    /// Footer visibility can be toggled off after being set on.
    /// </summary>
    [Fact]
    public void SetFooterVisibility_False_RemovesFooter()
    {
        var hfm = CreateManager();
        hfm.SetFooterVisibility(true);
        hfm.IsFooterVisible.Should().BeTrue();

        hfm.SetFooterVisibility(false);

        hfm.IsFooterVisible.Should().BeFalse();
    }

    /// <summary>
    /// Footer text can be set via SetFooterText.
    /// </summary>
    [Fact]
    public void SetFooterText_SetsText()
    {
        var hfm = CreateManager();
        hfm.SetFooterVisibility(true);

        hfm.SetFooterText("Confidential");

        hfm.IsFooterVisible.Should().BeTrue();
    }

    /// <summary>
    /// Slide number visibility can be set to true.
    /// </summary>
    [Fact]
    public void SetSlideNumberVisibility_True_MakesSlideNumberVisible()
    {
        var hfm = CreateManager();

        hfm.IsSlideNumberVisible.Should().BeFalse();

        hfm.SetSlideNumberVisibility(true);

        hfm.IsSlideNumberVisible.Should().BeTrue();
    }

    /// <summary>
    /// Both footer and slide number visibility persist together.
    /// </summary>
    [Fact]
    public void FooterAndSlideNumber_BothVisibleTogether()
    {
        var hfm = CreateManager();

        hfm.SetFooterVisibility(true);
        hfm.SetFooterText("Confidential");
        hfm.SetSlideNumberVisibility(true);

        hfm.IsFooterVisible.Should().BeTrue();
        hfm.IsSlideNumberVisible.Should().BeTrue();
    }

    /// <summary>
    /// DateTime visibility can be set and read.
    /// </summary>
    [Fact]
    public void SetDateTimeVisibility_True_MakesDateTimeVisible()
    {
        var hfm = CreateManager();

        hfm.IsDateTimeVisible.Should().BeFalse();

        hfm.SetDateTimeVisibility(true);

        hfm.IsDateTimeVisible.Should().BeTrue();
    }

    /// <summary>
    /// DateTime text can be set via SetDateTimeText.
    /// </summary>
    [Fact]
    public void SetDateTimeText_SetsText()
    {
        var hfm = CreateManager();
        hfm.SetDateTimeVisibility(true);

        hfm.SetDateTimeText("2024-01-01");

        hfm.IsDateTimeVisible.Should().BeTrue();
    }

    /// <summary>
    /// Header visibility can be set and read.
    /// </summary>
    [Fact]
    public void SetHeaderVisibility_True_MakesHeaderVisible()
    {
        var hfm = CreateManager();

        hfm.IsHeaderVisible.Should().BeFalse();

        hfm.SetHeaderVisibility(true);

        hfm.IsHeaderVisible.Should().BeTrue();
    }

    /// <summary>
    /// Header text can be set via SetHeaderText.
    /// </summary>
    [Fact]
    public void SetHeaderText_SetsText()
    {
        var hfm = CreateManager();
        hfm.SetHeaderVisibility(true);

        hfm.SetHeaderText("My Header");

        hfm.IsHeaderVisible.Should().BeTrue();
    }

    /// <summary>
    /// AsIBaseHandoutNotesSlideHeaderFooterManag returns the manager itself.
    /// </summary>
    [Fact]
    public void AsIBaseHandoutNotesSlideHeaderFooterManag_ReturnsSelf()
    {
        var hfm = CreateManager();

        IBaseHandoutNotesSlideHeaderFooterManager cast = hfm.AsIBaseHandoutNotesSlideHeaderFooterManag;

        cast.Should().BeSameAs(hfm);
    }

    /// <summary>
    /// AsIBaseSlideHeaderFooterManager returns the manager itself.
    /// </summary>
    [Fact]
    public void AsIBaseSlideHeaderFooterManager_ReturnsSelf()
    {
        var hfm = CreateManager();

        IBaseSlideHeaderFooterManager cast = hfm.AsIBaseSlideHeaderFooterManager;

        cast.Should().BeSameAs(hfm);
    }

    /// <summary>
    /// AsIBaseHeaderFooterManager returns the manager itself.
    /// </summary>
    [Fact]
    public void AsIBaseHeaderFooterManager_ReturnsSelf()
    {
        var hfm = CreateManager();

        IBaseHeaderFooterManager cast = hfm.AsIBaseHeaderFooterManager;

        cast.Should().BeSameAs(hfm);
    }
}
