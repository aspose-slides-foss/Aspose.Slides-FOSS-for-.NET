namespace Aspose.Slides.Foss.Internal;

/// <summary>
/// The content types this library declares for the parts it creates.
/// </summary>
/// <remarks>
/// Per ISO/IEC 29500-2 §10.1.2 a part's content type is its identity, and it is resolved from
/// <c>[Content_Types].xml</c> alone — an <c>Override</c> for the part, or a <c>Default</c> for its
/// extension. A slide part written with no <c>Override</c> falls through to
/// <c>&lt;Default Extension="xml"/&gt;</c> and is an <c>application/xml</c> part to every consumer,
/// whatever its markup says. PowerPoint guesses from the relationship type and opens it anyway;
/// stricter readers refuse the whole package. So every part created here has to be declared here too.
/// </remarks>
internal static class PartContentTypes
{
    internal const string Slide =
        "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";

    internal const string NotesSlide =
        "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";

    internal const string NotesMaster =
        "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml";

    internal const string Comments =
        "application/vnd.openxmlformats-officedocument.presentationml.comments+xml";

    internal const string CommentAuthors =
        "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml";

    internal const string ExtendedProperties =
        "application/vnd.openxmlformats-officedocument.extended-properties+xml";
}
