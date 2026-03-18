using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a notes slide in a presentation.
/// </summary>
public sealed class NotesSlide : BaseSlide, INotesSlide
{
    private NotesSlidePart? _notesPart;
    private OpcPackage? _package;
    private NotesSlideHeaderFooterManager? _headerFooterManager;
    private ITextFrame? _notesTextFrame;
    private ISlide? _parentSlide;

    /// <summary>
    /// Initializes internal state for a notes slide.
    /// </summary>
    /// <param name="presentation">The parent Presentation object.</param>
    /// <param name="package">The OPC package.</param>
    /// <param name="partName">The part name of this notes slide.</param>
    /// <param name="notesPart">The parsed NotesSlidePart.</param>
    /// <param name="parentSlide">The slide that owns this notes slide.</param>
    internal void InitInternal(
        IPresentation presentation,
        OpcPackage package,
        string partName,
        NotesSlidePart notesPart,
        ISlide parentSlide)
    {
        _presentationRef = presentation;
        _package = package;
        _partName = partName;
        _notesPart = notesPart;
        _parentSlide = parentSlide;

        var hfm = new NotesSlideHeaderFooterManager();
        hfm.InitInternal(notesPart);
        _headerFooterManager = hfm;
    }

    /// <summary>
    /// Flushes the notes slide state (XML + placeholders) to the OPC package.
    /// </summary>
    internal void FlushInternal()
    {
        _notesPart?.Flush();
    }

    /// <summary>
    /// Sets the notes text frame. Called internally during initialization.
    /// </summary>
    /// <param name="textFrame">The text frame containing notes text.</param>
    internal void SetNotesTextFrame(ITextFrame textFrame)
    {
        _notesTextFrame = textFrame;
    }

    /// <inheritdoc/>
    public INotesSlideHeaderFooterManager HeaderFooterManager =>
        _headerFooterManager ?? throw new InvalidOperationException("NotesSlide not initialized.");

    /// <inheritdoc/>
    public ITextFrame NotesTextFrame =>
        _notesTextFrame ?? throw new InvalidOperationException("NotesSlide not initialized.");

    /// <inheritdoc/>
    public ISlide ParentSlide =>
        _parentSlide ?? throw new InvalidOperationException("NotesSlide not initialized.");

    /// <inheritdoc/>
    public IBaseSlide AsIBaseSlide => this;
}
