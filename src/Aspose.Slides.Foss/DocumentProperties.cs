using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents the metadata properties of a presentation, wrapping OPC core, app, and custom property parts
/// with lazy initialization.
/// </summary>
public sealed class DocumentProperties : IDocumentProperties
{
    private OpcPackage _package = null!;
    private WeakReference<Presentation>? _presentationRef;

    private CorePropertiesPart? _corePart;
    private AppPropertiesPart? _appPart;
    private CustomPropertiesPart? _customPart;

    /// <summary>
    /// Initializes the document properties with the backing OPC package.
    /// </summary>
    internal void InitInternal(OpcPackage package, WeakReference<Presentation>? presentationRef = null)
    {
        _package = package;
        _presentationRef = presentationRef;
        _corePart = null;
        _appPart = null;
        _customPart = null;
    }

    private CorePropertiesPart EnsureCore()
    {
        return _corePart ??= CorePropertiesPart.CreateFromPackage(_package);
    }

    private AppPropertiesPart EnsureApp()
    {
        return _appPart ??= AppPropertiesPart.CreateFromPackage(_package);
    }

    private CustomPropertiesPart EnsureCustom()
    {
        return _customPart ??= CustomPropertiesPart.CreateFromPackage(_package);
    }

    // ── Core string properties ────────────────────────────────

    /// <inheritdoc />
    public string Title
    {
        get => EnsureCore().Title ?? string.Empty;
        set { EnsureCore().Title = value; EnsureCore().MarkDirty(); }
    }

    /// <inheritdoc />
    public string Subject
    {
        get => EnsureCore().Subject ?? string.Empty;
        set { EnsureCore().Subject = value; EnsureCore().MarkDirty(); }
    }

    /// <inheritdoc />
    public string Author
    {
        get => EnsureCore().Creator ?? string.Empty;
        set { EnsureCore().Creator = value; EnsureCore().MarkDirty(); }
    }

    /// <inheritdoc />
    public string Keywords
    {
        get => EnsureCore().Keywords ?? string.Empty;
        set { EnsureCore().Keywords = value; EnsureCore().MarkDirty(); }
    }

    /// <inheritdoc />
    public string Comments
    {
        get => EnsureCore().Description ?? string.Empty;
        set { EnsureCore().Description = value; EnsureCore().MarkDirty(); }
    }

    /// <inheritdoc />
    public string Category
    {
        get => EnsureCore().Category ?? string.Empty;
        set { EnsureCore().Category = value; EnsureCore().MarkDirty(); }
    }

    /// <inheritdoc />
    public string ContentStatus
    {
        get => EnsureCore().ContentStatus ?? string.Empty;
        set { EnsureCore().ContentStatus = value; EnsureCore().MarkDirty(); }
    }

    /// <inheritdoc />
    public string ContentType
    {
        get => EnsureCore().ContentType ?? string.Empty;
        set { EnsureCore().ContentType = value; EnsureCore().MarkDirty(); }
    }

    /// <inheritdoc />
    public string LastSavedBy
    {
        get => EnsureCore().LastModifiedBy ?? string.Empty;
        set { EnsureCore().LastModifiedBy = value; EnsureCore().MarkDirty(); }
    }

    // ── Core non-string properties ────────────────────────────

    /// <inheritdoc />
    public int RevisionNumber
    {
        get
        {
            var text = EnsureCore().Revision;
            if (string.IsNullOrEmpty(text))
                return 0;
            return int.TryParse(text, out var v) ? v : 0;
        }
        set
        {
            EnsureCore().Revision = value.ToString();
            EnsureCore().MarkDirty();
        }
    }

    /// <inheritdoc />
    public DateTime? CreatedTime
    {
        get => EnsureCore().Created;
        set { EnsureCore().Created = value; EnsureCore().MarkDirty(); }
    }

    /// <inheritdoc />
    public DateTime? LastSavedTime
    {
        get => EnsureCore().Modified;
        set { EnsureCore().Modified = value; EnsureCore().MarkDirty(); }
    }

    /// <inheritdoc />
    public DateTime? LastPrinted
    {
        get => EnsureCore().LastPrinted;
        set { EnsureCore().LastPrinted = value; EnsureCore().MarkDirty(); }
    }

    // ── App string properties ─────────────────────────────────

    /// <inheritdoc />
    public string AppVersion => EnsureApp().AppVersion ?? string.Empty;

    /// <inheritdoc />
    public string NameOfApplication
    {
        get => EnsureApp().Application ?? string.Empty;
        set { EnsureApp().Application = value; EnsureApp().MarkDirty(); }
    }

    /// <inheritdoc />
    public string Company
    {
        get => EnsureApp().Company ?? string.Empty;
        set { EnsureApp().Company = value; EnsureApp().MarkDirty(); }
    }

    /// <inheritdoc />
    public string Manager
    {
        get => EnsureApp().Manager ?? string.Empty;
        set { EnsureApp().Manager = value; EnsureApp().MarkDirty(); }
    }

    /// <inheritdoc />
    public string PresentationFormat
    {
        get => EnsureApp().PresentationFormat ?? string.Empty;
        set { EnsureApp().PresentationFormat = value; EnsureApp().MarkDirty(); }
    }

    /// <inheritdoc />
    public string ApplicationTemplate
    {
        get => EnsureApp().Template ?? string.Empty;
        set { EnsureApp().Template = value; EnsureApp().MarkDirty(); }
    }

    /// <inheritdoc />
    public string HyperlinkBase
    {
        get => EnsureApp().HyperlinkBase ?? string.Empty;
        set { EnsureApp().HyperlinkBase = value; EnsureApp().MarkDirty(); }
    }

    // ── App non-string properties ─────────────────────────────

    /// <inheritdoc />
    public TimeSpan TotalEditingTime
    {
        get => EnsureApp().TotalTime;
        set { EnsureApp().TotalTime = value; EnsureApp().MarkDirty(); }
    }

    /// <inheritdoc />
    public bool SharedDoc
    {
        get => EnsureApp().SharedDoc;
        set { EnsureApp().SharedDoc = value; EnsureApp().MarkDirty(); }
    }

    /// <inheritdoc />
    public bool ScaleCrop
    {
        get => EnsureApp().ScaleCrop;
        set { EnsureApp().ScaleCrop = value; EnsureApp().MarkDirty(); }
    }

    /// <inheritdoc />
    public bool LinksUpToDate
    {
        get => EnsureApp().LinksUpToDate;
        set { EnsureApp().LinksUpToDate = value; EnsureApp().MarkDirty(); }
    }

    /// <inheritdoc />
    public bool HyperlinksChanged
    {
        get => EnsureApp().HyperlinksChanged;
        set { EnsureApp().HyperlinksChanged = value; EnsureApp().MarkDirty(); }
    }

    // ── App read-only int properties ──────────────────────────

    /// <summary>
    /// Recomputes the counts <c>docProps/app.xml</c> reports about the deck.
    /// </summary>
    internal void RefreshDerivedProperties(DeckStatistics statistics)
    {
        EnsureApp().SetDerivedCounts(statistics);
    }

    /// <inheritdoc />
    public int Slides => EnsureApp().Slides;

    /// <inheritdoc />
    public int HiddenSlides => EnsureApp().HiddenSlides;

    /// <inheritdoc />
    public int Notes => EnsureApp().Notes;

    /// <inheritdoc />
    public int Paragraphs => EnsureApp().Paragraphs;

    /// <inheritdoc />
    public int Words => EnsureApp().Words;

    /// <inheritdoc />
    public int MultimediaClips => EnsureApp().MultimediaClips;

    /// <inheritdoc />
    public IReadOnlyList<IHeadingPair> HeadingPairs
    {
        get
        {
            var data = EnsureApp().HeadingPairs;
            var result = new List<IHeadingPair>(data.Count);
            foreach (var hp in data)
                result.Add(new HeadingPair(hp.Name, hp.Count));
            return result;
        }
    }

    /// <inheritdoc />
    public IReadOnlyList<string> TitlesOfParts => EnsureApp().TitlesOfParts;

    // ── Custom properties ─────────────────────────────────────

    /// <inheritdoc />
    public int CountOfCustomProperties => EnsureCustom().Count;

    /// <inheritdoc />
    public object? GetCustomPropertyValue(string name) => EnsureCustom().GetValue(name);

    /// <inheritdoc />
    public void SetCustomPropertyValue(string name, object value) => EnsureCustom().SetValue(name, value);

    /// <inheritdoc />
    public string GetCustomPropertyName(int index) => EnsureCustom().GetNameAt(index);

    /// <inheritdoc />
    public bool RemoveCustomProperty(string name) => EnsureCustom().Remove(name);

    /// <inheritdoc />
    public bool ContainsCustomProperty(string name) => EnsureCustom().Contains(name);

    /// <inheritdoc />
    public void ClearCustomProperties() => EnsureCustom().Clear();

    // ── Save ───────────────────────────────────────────────────

    /// <summary>
    /// Serializes all dirty property parts back to the OPC package.
    /// Only parts that were previously accessed and modified are written.
    /// </summary>
    internal void Save()
    {
        // Recording when the file was last written is part of writing it, and it must not depend on
        // the caller having touched any core property. Without this, dcterms:modified stayed at
        // whatever date the deck was created, which is worse than absent: it is a wrong answer.
        EnsureCore().Modified = DateTime.UtcNow;
        EnsureCore().MarkDirty();

        if (_corePart is { IsDirty: true })
            SavePart("docProps/core.xml", _corePart.Root);
        if (_appPart is { IsDirty: true })
            SavePart("docProps/app.xml", _appPart.Root);
        if (_customPart is { IsDirty: true })
            SavePart("docProps/custom.xml", _customPart.Root);
    }

    private void SavePart(string partPath, System.Xml.Linq.XElement root)
    {
        using var ms = new MemoryStream();
        root.Save(ms);
        _package.SetPart(partPath, ms.ToArray());
    }

    // ── Clear ──────────────────────────────────────────────────

    /// <inheritdoc />
    public void ClearBuiltInProperties()
    {
        // Reset core properties
        var core = EnsureCore();
        core.Title = string.Empty;
        core.Subject = string.Empty;
        core.Creator = string.Empty;
        core.Keywords = string.Empty;
        core.Description = string.Empty;
        core.Category = string.Empty;
        core.ContentStatus = string.Empty;
        core.ContentType = string.Empty;
        core.LastModifiedBy = string.Empty;
        core.Revision = "0";
        core.Created = null;
        core.Modified = null;
        core.LastPrinted = null;
        core.MarkDirty();

        // Reset app properties
        var app = EnsureApp();
        app.ClearAll();
        app.MarkDirty();
    }
}
