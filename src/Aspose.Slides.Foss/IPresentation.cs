using Aspose.Slides.Foss.Export;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a presentation document.
/// </summary>
public abstract class IPresentation : IPresentationComponent
{
    /// <inheritdoc />
    public override IPresentation? Presentation => this;
    /// <summary>
    /// Gets or sets the date and time used to substitute content of datetime fields.
    /// Defaults to the time the <see cref="IPresentation"/> object was created.
    /// </summary>
    public abstract DateTime CurrentDateTime { get; set; }

    /// <summary>
    /// Gets the collection of all slides defined in the presentation.
    /// </summary>
    public abstract ISlideCollection Slides { get; }

    /// <summary>
    /// Gets the notes slide size object for this presentation.
    /// </summary>
    public abstract INotesSize NotesSize { get; }

    /// <summary>
    /// Gets the collection of all layout slides defined in the presentation.
    /// </summary>
    public abstract IGlobalLayoutSlideCollection LayoutSlides { get; }

    /// <summary>
    /// Gets the collection of all master slides defined in the presentation.
    /// </summary>
    public abstract IMasterSlideCollection Masters { get; }

    /// <summary>
    /// Gets the collection of all sections in the presentation.
    /// </summary>
    public abstract ISectionCollection Sections { get; }

    /// <summary>
    /// Gets the collection of comment authors.
    /// </summary>
    public abstract ICommentAuthorCollection CommentAuthors { get; }

    /// <summary>
    /// Gets the document properties (metadata) of the presentation.
    /// </summary>
    public abstract IDocumentProperties DocumentProperties { get; }

    /// <summary>
    /// Gets the collection of all images embedded in the presentation.
    /// </summary>
    public abstract IImageCollection Images { get; }

    /// <summary>
    /// Gets the format the presentation was loaded from.
    /// </summary>
    public abstract SourceFormat SourceFormat { get; }

    /// <summary>
    /// Gets or sets the first slide number in the presentation.
    /// </summary>
    public abstract int FirstSlideNumber { get; set; }

    /// <summary>
    /// Returns this instance as an <see cref="IPresentationComponent"/>.
    /// </summary>
    public abstract IPresentationComponent AsIPresentationComponent { get; }

    /// <summary>
    /// Saves the presentation to a file in the specified format.
    /// </summary>
    /// <param name="fname">The output file path.</param>
    /// <param name="format">The export format.</param>
    public abstract void Save(string fname, SaveFormat format);

    /// <summary>
    /// Saves the presentation to a stream in the specified format.
    /// </summary>
    /// <param name="stream">The output stream.</param>
    /// <param name="format">The export format.</param>
    public abstract void Save(Stream stream, SaveFormat format);

    /// <summary>
    /// Saves the presentation to a file with the specified format and options.
    /// </summary>
    /// <param name="fname">The output file path.</param>
    /// <param name="format">The export format.</param>
    /// <param name="options">Additional export options.</param>
    public abstract void Save(string fname, SaveFormat format, ISaveOptions options);

    /// <summary>
    /// Saves the presentation to a stream with the specified format and options.
    /// </summary>
    /// <param name="stream">The output stream.</param>
    /// <param name="format">The export format.</param>
    /// <param name="options">Additional export options.</param>
    public abstract void Save(Stream stream, SaveFormat format, ISaveOptions options);

    /// <summary>
    /// Saves specific slides to a file in the specified format.
    /// </summary>
    /// <param name="fname">The output file path.</param>
    /// <param name="slides">The indices of slides to export.</param>
    /// <param name="format">The export format.</param>
    public abstract void Save(string fname, int[] slides, SaveFormat format);

    /// <summary>
    /// Saves specific slides to a file with the specified format and options.
    /// </summary>
    /// <param name="fname">The output file path.</param>
    /// <param name="slides">The indices of slides to export.</param>
    /// <param name="format">The export format.</param>
    /// <param name="options">Additional export options.</param>
    public abstract void Save(string fname, int[] slides, SaveFormat format, ISaveOptions options);

    /// <summary>
    /// Saves specific slides to a stream in the specified format.
    /// </summary>
    /// <param name="stream">The output stream.</param>
    /// <param name="slides">The indices of slides to export.</param>
    /// <param name="format">The export format.</param>
    public abstract void Save(Stream stream, int[] slides, SaveFormat format);

    /// <summary>
    /// Saves specific slides to a stream with the specified format and options.
    /// </summary>
    /// <param name="stream">The output stream.</param>
    /// <param name="slides">The indices of slides to export.</param>
    /// <param name="format">The export format.</param>
    /// <param name="options">Additional export options.</param>
    public abstract void Save(Stream stream, int[] slides, SaveFormat format, ISaveOptions options);

    /// <summary>
    /// Saves the presentation using the specified options object.
    /// </summary>
    /// <param name="options">Export options controlling format and destination.</param>
    public abstract void Save(ISaveOptions options);
}
