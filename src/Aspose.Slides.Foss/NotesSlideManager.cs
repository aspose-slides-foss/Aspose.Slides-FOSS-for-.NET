using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Manages notes slide operations for a slide.
/// </summary>
public sealed class NotesSlideManager : INotesSlideManager
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private ISlide? _parentSlide;
    private OpcPackage? _package;
    private SlidePart? _slidePart;
    private NotesSlide? _notesSlideCache;

    /// <summary>
    /// Initializes internal state. Called after default construction to inject dependencies.
    /// </summary>
    internal void InitInternal(ISlide parentSlide, OpcPackage? package, SlidePart? slidePart)
    {
        _parentSlide = parentSlide;
        _package = package;
        _slidePart = slidePart;
    }

    /// <inheritdoc/>
    public INotesSlide? NotesSlide
    {
        get
        {
            if (_notesSlideCache is not null)
                return _notesSlideCache;

            var partName = GetNotesPartName();
            if (partName is null)
                return null;

            return LoadNotesSlide(partName);
        }
    }

    /// <inheritdoc/>
    public INotesSlide AddNotesSlide()
    {
        var existing = NotesSlide;
        if (existing is not null)
            return existing;

        if (_package is null || _slidePart is null)
            throw new InvalidOperationException("NotesSlideManager not properly initialized.");

        var slidePartName = _slidePart.PartName;

        // Create the notes slide XML part in the OPC package. The presentation part comes along
        // because a notes slide needs a notes master, and a notes master has to be registered in
        // ppt/presentation.xml to be one.
        var presentationPart = (_parentSlide?.Presentation as Presentation)?.PresentationPartInternal;
        var notesPart = NotesSlidePart.CreateEmpty(_package, slidePartName, presentationPart);

        // Add relationship from slide → notes slide
        var relativeTarget = NotesSlidePart.ComputeRelativeTarget(slidePartName, notesPart.PartName);
        _slidePart.RelsManager.Add(NotesSlidePart.RelType, relativeTarget);
        PersistSlideRels();

        return LoadNotesSlide(notesPart.PartName);
    }

    /// <inheritdoc/>
    public void RemoveNotesSlide()
    {
        var partName = GetNotesPartName();
        if (partName is null)
            return;

        if (_package is null || _slidePart is null)
            return;

        // Remove all notes slide relationships from the slide
        var rels = _slidePart.RelsManager.FindByType(NotesSlidePart.RelType).ToList();
        foreach (var rel in rels)
            _slidePart.RelsManager.Remove(rel.Id);
        PersistSlideRels();

        // Delete notes slide part files from OPC
        NotesSlidePart.Delete(_package, partName);

        _notesSlideCache = null;
    }

    /// <summary>
    /// Flushes the notes slide state to the OPC package.
    /// </summary>
    internal void Flush()
    {
        _notesSlideCache?.FlushInternal();
    }

    // ── Private helpers ──────────────────────────────────────

    private string? GetNotesPartName()
    {
        if (_slidePart is null)
            return null;

        var rels = _slidePart.RelsManager.FindByType(NotesSlidePart.RelType);
        var notesRel = rels.FirstOrDefault();
        if (notesRel is null)
            return null;

        return ResolvePartName(GetDirectoryPath(_slidePart.PartName), notesRel.Target);
    }

    private NotesSlide LoadNotesSlide(string partName)
    {
        var notesPart = NotesSlidePart.Load(_package!, partName);
        notesPart ??= new NotesSlidePart();

        var notesSlide = new NotesSlide();
        notesSlide.InitInternal(_parentSlide!.Presentation!, _package!, partName, notesPart, _parentSlide!);

        // Set up the text frame from the notes XML body placeholder
        var txBody = notesPart.GetNotesTxBody();
        if (txBody is not null)
        {
            var textFrame = new TextFrame();
            textFrame.InitInternal(txBody, slidePart: null, parentSlide: notesSlide, parentShape: null);
            notesSlide.SetNotesTextFrame(textFrame);
        }
        else
        {
            // Fallback: create a minimal txBody
            var txBodyElement = new XElement(ANs + "txBody",
                new XElement(ANs + "p",
                    new XElement(ANs + "r",
                        new XElement(ANs + "t", ""))));
            var textFrame = new TextFrame();
            textFrame.InitInternal(txBodyElement, slidePart: null, parentSlide: notesSlide, parentShape: null);
            notesSlide.SetNotesTextFrame(textFrame);
        }

        _notesSlideCache = notesSlide;
        return notesSlide;
    }

    private void PersistSlideRels()
    {
        if (_slidePart is null || _package is null)
            return;

        var relsPath = GetRelsPath(_slidePart.PartName);
        _package.SetPart(relsPath, _slidePart.RelsManager.ToBytes());
    }

    private static string ResolvePartName(string basePath, string target)
    {
        var combined = basePath + target;
        var segments = combined.Split('/');
        var resolved = new Stack<string>();

        foreach (var segment in segments)
        {
            if (segment == "..")
            {
                if (resolved.Count > 0)
                    resolved.Pop();
            }
            else if (segment != "." && segment.Length > 0)
            {
                resolved.Push(segment);
            }
        }

        return string.Join("/", resolved.Reverse());
    }

    private static string GetDirectoryPath(string partName)
    {
        var lastSlash = partName.LastIndexOf('/');
        return lastSlash >= 0 ? partName[..(lastSlash + 1)] : "";
    }

    private static string GetRelsPath(string partName)
    {
        var dir = GetDirectoryPath(partName);
        var file = partName[dir.Length..];
        return $"{dir}_rels/{file}.rels";
    }
}
