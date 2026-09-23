namespace Aspose.Slides.Foss.ConformanceTests.Harness;

/// <summary>
/// The input decks these tests open. They are copied next to the test assembly at build time.
/// </summary>
internal static class TestFixtures
{
    /// <summary>A minimal deck: one slide, one layout, one master, one theme.</summary>
    internal static string SimpleDeck => PathTo("Presentation.pptx");

    /// <summary>
    /// A 46-part deck authored by PowerPoint: two slides, twelve layouts, a notes master with a
    /// theme part of its own, a notes slide, <c>presProps</c>, <c>viewProps</c>, <c>tableStyles</c>
    /// and a thumbnail — the parts a round-trip is most likely to lose, and none of which this
    /// library writes itself. A fixture written by this library cannot stand in for it.
    /// </summary>
    internal static string PowerPointDeck => PathTo("PowerPointDeck.pptx");

    /// <summary>
    /// A one-table deck whose <c>a:tblPr</c> already carries an <c>a:extLst</c>, as a table authored
    /// by an application that stores extensions on it does. Nothing this library writes produces
    /// one, so it has to be opened rather than built.
    /// </summary>
    internal static string TableWithExtension => PathTo("TableWithExtension.pptx");

    /// <summary>
    /// A deck authored in PowerPoint with sections: three slides whose rectangles read
    /// <c>Slide A</c>, <c>Slide B</c> and <c>Slide C</c>, in the sections <c>Intro</c> (A) and
    /// <c>Body</c> (B, C). The section list is PowerPoint's own, so it is the shape a user's file
    /// actually has.
    /// </summary>
    internal static string PowerPointSections => PathTo("PowerPointSections.pptx");

    /// <summary>
    /// What PowerPoint itself saved after deleting slide B from <see cref="PowerPointSections"/>:
    /// slide ids 256 and 258, with <c>Intro</c> (A) and <c>Body</c> (C). PowerPoint does not
    /// renumber the slides that remain, so the ids have a gap, as they do in most real decks.
    /// </summary>
    internal static string PowerPointSectionsWithIdGap => PathTo("PowerPointSectionsWithIdGap.pptx");

    /// <summary>
    /// A one-slide deck authored in PowerPoint with one comment and one reply to it, both by
    /// <c>Alice Ref</c>: <c>Parent comment</c> at 10 pt from the slide's top-left corner and
    /// <c>Reply to parent</c> below it. The comments are PowerPoint's own classic ones: each
    /// <c>&lt;p:cm&gt;</c> carries a <c>p15:threadingInfo</c> extension, the reply's names its parent
    /// with <c>p15:parentCm</c>, and <c>p:pos</c> is in PowerPoint's unit — <c>x="80" y="80"</c>
    /// for 10 pt.
    /// </summary>
    internal static string PowerPointCommentReply => PathTo("PowerPointCommentReply.pptx");

    /// <summary>
    /// The same thread as <see cref="PowerPointCommentReply"/>, with the reply written by a second
    /// author, <c>Bob Ref</c>. Comment indices are counted per author, so both comments have
    /// <c>idx="1"</c> and only the author tells the parent from the reply.
    /// </summary>
    internal static string PowerPointCommentReplyTwoAuthors => PathTo("PowerPointCommentReplyTwoAuthors.pptx");

    private static string PathTo(string fileName)
    {
        var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test_data", fileName);
        Assert.True(File.Exists(path), $"The fixture '{fileName}' was not copied to '{path}'.");
        return path;
    }
}
