using System.Xml.Linq;
using Aspose.Slides.Foss.Export;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents a Microsoft PowerPoint presentation document.
/// Supports loading from files, streams, or creating new empty presentations.
/// </summary>
public sealed class Presentation : IPresentation, IDisposable
{
    private static readonly string SlideLayoutRelType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";

    private OpcPackage? _opcPackage;
    private PresentationPart? _presentationPart;
    private SourceFormat _sourceFormat = SourceFormat.Pptx;
    private DateTime _currentDateTime = DateTime.Now;
    private int _firstSlideNumber = 1;
    private ILoadOptions? _loadOptions;

    // Lazy-initialized collections (null means not yet created)
    private SlideCollection? _slides;
    private GlobalLayoutSlideCollection? _layoutSlidesCollection;
    private MasterSlideCollection? _mastersCollection;
    private NotesSize? _notesSize;
    private SectionCollection? _sections;
    private CommentAuthorCollection? _commentAuthors;
    private CommentAuthorsPart? _commentAuthorsPart;
    private DocumentProperties? _documentProperties;
    private ImageCollection? _imagesCollection;

    // Maps populated by EnsureLayoutSlidesParsed
    private Dictionary<string, IMasterSlide>? _masterSlidesMap;
    private Dictionary<string, ILayoutSlide>? _layoutSlidesMap;

    /// <summary>
    /// Creates a new empty presentation from a minimal PPTX template.
    /// </summary>
    public Presentation()
    {
        InitEmpty();
    }

    /// <summary>
    /// Creates a new empty presentation with the specified load options.
    /// </summary>
    /// <param name="loadOptions">Options controlling load behavior.</param>
    public Presentation(ILoadOptions loadOptions)
    {
        _loadOptions = loadOptions;
        InitEmpty();
    }

    /// <summary>
    /// Loads a presentation from the specified stream.
    /// </summary>
    /// <param name="stream">A stream containing the presentation data.</param>
    public Presentation(Stream stream)
    {
        InitFromStream(stream);
    }

    /// <summary>
    /// Loads a presentation from the specified stream with load options.
    /// </summary>
    /// <param name="stream">A stream containing the presentation data.</param>
    /// <param name="loadOptions">Options controlling load behavior.</param>
    public Presentation(Stream stream, ILoadOptions loadOptions)
    {
        _loadOptions = loadOptions;
        InitFromStream(stream);
    }

    /// <summary>
    /// Loads a presentation from the specified file path.
    /// </summary>
    /// <param name="file">The path to the presentation file.</param>
    public Presentation(string file)
    {
        InitFromFile(file);
    }

    /// <summary>
    /// Loads a presentation from the specified file path with load options.
    /// </summary>
    /// <param name="file">The path to the presentation file.</param>
    /// <param name="loadOptions">Options controlling load behavior.</param>
    public Presentation(string file, ILoadOptions loadOptions)
    {
        _loadOptions = loadOptions;
        InitFromFile(file);
    }

    // ── Properties ────────────────────────────────────────────

    /// <inheritdoc />
    public override DateTime CurrentDateTime
    {
        get => _currentDateTime;
        set => _currentDateTime = value;
    }

    /// <inheritdoc />
    public override ISlideCollection Slides
    {
        get
        {
            if (_slides is null)
            {
                EnsureLayoutSlidesParsed();
                _slides = new SlideCollection();
                _slides.InitInternal(this, _opcPackage!, _presentationPart!, ResolveLayoutSlide);
            }
            return _slides;
        }
    }

    /// <inheritdoc />
    public override INotesSize NotesSize
    {
        get
        {
            if (_notesSize is null)
            {
                _notesSize = new NotesSize();
                _notesSize.InitInternal(_presentationPart!);
            }
            return _notesSize;
        }
    }

    /// <inheritdoc />
    public override IGlobalLayoutSlideCollection LayoutSlides
    {
        get
        {
            if (_layoutSlidesCollection is null)
            {
                EnsureLayoutSlidesParsed();
                var layouts = _layoutSlidesMap!.Values.ToList();
                _layoutSlidesCollection = new GlobalLayoutSlideCollection();
                _layoutSlidesCollection.InitInternal(layouts);
            }
            return _layoutSlidesCollection;
        }
    }

    /// <inheritdoc />
    public override IMasterSlideCollection Masters
    {
        get
        {
            if (_mastersCollection is null)
            {
                EnsureLayoutSlidesParsed();
                _mastersCollection = new MasterSlideCollection();
                _mastersCollection.InitInternal(this, _opcPackage!, _presentationPart!, _masterSlidesMap!.Values.ToList());
            }
            return _mastersCollection;
        }
    }

    /// <inheritdoc />
    public override ISectionCollection Sections => _sections ??= new SectionCollection();

    /// <inheritdoc />
    public override ICommentAuthorCollection CommentAuthors
    {
        get
        {
            if (_commentAuthors is null)
            {
                _commentAuthorsPart = LoadCommentAuthorsPart(_opcPackage!);
                _commentAuthors = new CommentAuthorCollection();
                _commentAuthors.InitInternal(_commentAuthorsPart, _opcPackage!, this);
            }
            return _commentAuthors;
        }
    }

    /// <inheritdoc />
    public override IDocumentProperties DocumentProperties
    {
        get
        {
            if (_documentProperties is null)
            {
                _documentProperties = new DocumentProperties();
                _documentProperties.InitInternal(_opcPackage!, new WeakReference<Presentation>(this));
            }
            return _documentProperties;
        }
    }

    /// <inheritdoc />
    public override IImageCollection Images
    {
        get
        {
            if (_imagesCollection is null)
            {
                _imagesCollection = new ImageCollection();
                _imagesCollection.InitInternal(_opcPackage!, new ContentTypesManager());
            }
            return _imagesCollection;
        }
    }

    /// <inheritdoc />
    public override SourceFormat SourceFormat => _sourceFormat;

    /// <inheritdoc />
    public override int FirstSlideNumber
    {
        get => _firstSlideNumber;
        set
        {
            _firstSlideNumber = value;
            _presentationPart?.SetFirstSlideNumber(value);
        }
    }

    /// <inheritdoc />
    public override IPresentationComponent AsIPresentationComponent => this;

    /// <summary>
    /// Gets the image collection for internal use by picture-related classes.
    /// </summary>
    internal ImageCollection? ImagesInternal => _imagesCollection;

    /// <summary>
    /// Gets the list of slides in this presentation for internal use.
    /// </summary>
    internal List<ISlide> SlidesInternal { get; } = [];

    /// <summary>
    /// Gets the backing OPC package.
    /// </summary>
    internal OpcPackage? OpcPackageInternal => _opcPackage;

    /// <summary>
    /// Gets the presentation part.
    /// </summary>
    internal PresentationPart? PresentationPartInternal => _presentationPart;

    /// <summary>
    /// Gets the master slides map for internal use.
    /// </summary>
    internal Dictionary<string, IMasterSlide>? MasterSlidesMapInternal => _masterSlidesMap;

    /// <summary>
    /// Gets the layout slides map for internal use.
    /// </summary>
    internal Dictionary<string, ILayoutSlide>? LayoutSlidesMapInternal => _layoutSlidesMap;

    // ── Save methods ──────────────────────────────────────────

    /// <inheritdoc />
    public override void Save(string fname, SaveFormat format)
    {
        using var stream = File.Create(fname);
        Save(stream, format);
    }

    /// <inheritdoc />
    public override void Save(Stream stream, SaveFormat format)
    {
        _presentationPart?.Flush();
        _documentProperties?.Save();
        FlushComments();
        FlushNotesSlides();
        FlushSlides();
        _opcPackage?.SaveToStream(stream);
    }

    /// <inheritdoc />
    public override void Save(string fname, SaveFormat format, ISaveOptions options)
    {
        Save(fname, format);
    }

    /// <inheritdoc />
    public override void Save(Stream stream, SaveFormat format, ISaveOptions options)
    {
        Save(stream, format);
    }

    /// <inheritdoc />
    public override void Save(string fname, int[] slides, SaveFormat format)
    {
        Save(fname, format);
    }

    /// <inheritdoc />
    public override void Save(string fname, int[] slides, SaveFormat format, ISaveOptions options)
    {
        Save(fname, format);
    }

    /// <inheritdoc />
    public override void Save(Stream stream, int[] slides, SaveFormat format)
    {
        Save(stream, format);
    }

    /// <inheritdoc />
    public override void Save(Stream stream, int[] slides, SaveFormat format, ISaveOptions options)
    {
        Save(stream, format);
    }

    /// <inheritdoc />
    public override void Save(ISaveOptions options)
    {
        throw new InvalidOperationException("A file path or stream is required. Use Save(string, SaveFormat) or Save(Stream, SaveFormat) instead.");
    }

    // ── IDisposable ───────────────────────────────────────────

    /// <summary>
    /// Releases resources used by this presentation.
    /// </summary>
    public void Dispose()
    {
        _opcPackage = null;
        _presentationPart = null;
        _slides = null;
        _layoutSlidesCollection = null;
        _mastersCollection = null;
        _notesSize = null;
        _sections = null;
        _commentAuthors = null;
        _commentAuthorsPart = null;
        _documentProperties = null;
        _imagesCollection = null;
        _masterSlidesMap = null;
        _layoutSlidesMap = null;
    }

    // ── Private initialization ────────────────────────────────

    private void InitEmpty()
    {
        _opcPackage = OpcPackage.CreateNew();
        MinimalPptxCreator.Populate(_opcPackage);
        FinishInit();
    }

    private void InitFromFile(string path)
    {
        _opcPackage = OpcPackage.Open(path);
        DetectSourceFormat(path);
        FinishInit();
    }

    private void InitFromStream(Stream stream)
    {
        _opcPackage = OpcPackage.Open(stream);
        FinishInit();
    }

    private void FinishInit()
    {
        _presentationPart = PresentationPart.CreateFromPackage(_opcPackage!);
        _firstSlideNumber = _presentationPart.ReadFirstSlideNumber();
    }

    private void DetectSourceFormat(string path)
    {
        var ext = Path.GetExtension(path);
        if (string.IsNullOrEmpty(ext))
            return;

        _sourceFormat = ext.ToLowerInvariant() switch
        {
            ".pptx" or ".pptm" or ".ppsx" or ".potx" => SourceFormat.Pptx,
            ".ppt" => SourceFormat.Ppt,
            ".odp" => SourceFormat.Odp,
            _ => _sourceFormat,
        };
    }

    private void EnsureLayoutSlidesParsed()
    {
        if (_masterSlidesMap is not null)
            return;

        _masterSlidesMap = new Dictionary<string, IMasterSlide>(StringComparer.OrdinalIgnoreCase);
        _layoutSlidesMap = new Dictionary<string, ILayoutSlide>(StringComparer.OrdinalIgnoreCase);

        var masterEntries = _presentationPart!.GetMasterEntries();
        foreach (var (_, relId) in masterEntries)
        {
            var rel = _presentationPart.Rels.GetById(relId);
            if (rel is null)
                continue;

            var masterPartName = ResolvePartName("ppt/", rel.Target);
            var master = new MasterSlide();
            master.PartName = masterPartName;
            master.PackageInternal = _opcPackage;
            master.SetPresentation(this);
            _masterSlidesMap[masterPartName] = master;

            // Parse layout slides referenced by this master
            var masterRelsName = GetRelsPath(masterPartName);
            var masterRelsData = _opcPackage!.GetPart(masterRelsName);
            if (masterRelsData is null)
                continue;

            var masterRels = new RelsManager();
            masterRels.Load(masterRelsData);

            var masterLayouts = new List<ILayoutSlide>();
            foreach (var layoutRel in masterRels.FindByType(SlideLayoutRelType))
            {
                var layoutPartName = ResolvePartName(
                    GetDirectoryPath(masterPartName),
                    layoutRel.Target);
                if (!_layoutSlidesMap.ContainsKey(layoutPartName))
                {
                    var layout = new LayoutSlide();
                    _layoutSlidesMap[layoutPartName] = layout;
                }
                masterLayouts.Add(_layoutSlidesMap[layoutPartName]);
            }

            var layoutCollection = new MasterLayoutSlideCollection();
            layoutCollection.InitInternal(masterLayouts);
            master.SetLayoutSlides(layoutCollection);
        }
    }

    private ILayoutSlide? ResolveLayoutSlide(string partName)
    {
        EnsureLayoutSlidesParsed();
        return _layoutSlidesMap!.TryGetValue(partName, out var layout) ? layout : null;
    }

    /// <summary>
    /// Resolves a master slide part name to its <see cref="IMasterSlide"/> object.
    /// </summary>
    /// <param name="partName">The OPC part name of the master slide.</param>
    /// <returns>The master slide, or <c>null</c> if not found.</returns>
    internal IMasterSlide? ResolveMasterSlide(string partName)
    {
        EnsureLayoutSlidesParsed();
        return _masterSlidesMap!.TryGetValue(partName, out var master) ? master : null;
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
        var file = partName[(dir.Length)..];
        return $"{dir}_rels/{file}.rels";
    }

    private void FlushComments()
    {
        if (_commentAuthorsPart is null || _opcPackage is null)
            return;

        _commentAuthorsPart.Flush(_opcPackage);

        // Only register OPC metadata if there are cached comment parts (i.e. comments were added).
        if (_commentAuthorsPart.CpCache.Count == 0)
            return;

        // Register commentAuthors.xml content type + relationship from presentation.
        OpcRegistration.AddContentTypeOverride(_opcPackage, "ppt/commentAuthors.xml",
            "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml");
        OpcRegistration.EnsureRelationship(_opcPackage, "ppt/presentation.xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors",
            "commentAuthors.xml");

        // Register each comments part content type + relationship from its slide.
        foreach (var commentsPartName in _commentAuthorsPart.CpCache.Keys)
        {
            OpcRegistration.AddContentTypeOverride(_opcPackage, commentsPartName,
                "application/vnd.openxmlformats-officedocument.presentationml.comments+xml");

            // Derive slide part name from comments part name (ppt/comments/comment1.xml → ppt/slides/slide1.xml).
            var slidePartName = commentsPartName
                .Replace("ppt/comments/comment", "ppt/slides/slide", StringComparison.Ordinal);
            OpcRegistration.EnsureRelationship(_opcPackage, slidePartName,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
                "../comments/" + commentsPartName[(commentsPartName.LastIndexOf('/') + 1)..]);
        }
    }

    private void FlushSlides()
    {
        foreach (var slide in SlidesInternal)
        {
            if (slide is Slide s)
            {
                s.GetSlidePartInternal()?.Save();
            }
        }
    }

    private void FlushNotesSlides()
    {
        foreach (var slide in SlidesInternal)
        {
            if (slide is Slide s)
            {
                var mgr = s.NotesSlideManager as NotesSlideManager;
                mgr?.Flush();
            }
        }
    }

    private static CommentAuthorsPart LoadCommentAuthorsPart(OpcPackage package)
    {
        const string partName = "ppt/commentAuthors.xml";
        var data = package.GetPart(partName);
        XElement root;
        if (data is not null)
        {
            using var ms = new MemoryStream(data);
            var doc = XDocument.Load(ms);
            root = doc.Root ?? new XElement(
                XName.Get("cmAuthorLst", "http://schemas.openxmlformats.org/presentationml/2006/main"));
        }
        else
        {
            root = new XElement(
                XName.Get("cmAuthorLst", "http://schemas.openxmlformats.org/presentationml/2006/main"));
        }
        return new CommentAuthorsPart(root);
    }
}
