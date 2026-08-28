using System.Text.RegularExpressions;
using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Represents the text body of a shape.
/// </summary>
public sealed class TextFrame : ISlideComponent, ITextFrame
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XName AP = ANs + "p";
    private static readonly XName AR = ANs + "r";
    private static readonly XName AT = ANs + "t";

    private XElement? _txBody;
    private SlidePart? _slidePart;
    private IBaseSlide? _parentSlide;
    private IShape? _parentShape;
    private ICell? _parentCell;

    /// <summary>
    /// Initializes internal state. Called after default construction to inject dependencies.
    /// </summary>
    internal void InitInternal(XElement txBody, SlidePart? slidePart, IBaseSlide? parentSlide, IShape? parentShape)
    {
        _txBody = txBody;
        _slidePart = slidePart;
        _parentSlide = parentSlide;
        _parentShape = parentShape;
    }

    /// <summary>
    /// Initializes internal state with a parent cell reference.
    /// </summary>
    internal void InitInternal(XElement txBody, SlidePart? slidePart, IBaseSlide? parentSlide, ICell? parentCell)
    {
        _txBody = txBody;
        _slidePart = slidePart;
        _parentSlide = parentSlide;
        _parentCell = parentCell;
    }

    /// <inheritdoc/>
    /// <remarks>
    /// The collection is bound to the <c>&lt;a:txBody&gt;</c> of this text frame, so it reads the
    /// paragraphs that are in the file and a paragraph added to it is written into the file. It was
    /// previously filled with copies instead, which made every addition a change to a list nobody
    /// read: the call returned, the save succeeded, and the text was not in the saved slide.
    /// </remarks>
    public IParagraphCollection Paragraphs
    {
        get
        {
            var coll = new ParagraphCollection();

            if (_txBody is not null)
                coll.InitInternal(_txBody, _slidePart, _parentSlide);
            else
                coll.InitInternal(_parentSlide);

            return coll;
        }
    }

    /// <inheritdoc/>
    public string Text
    {
        get
        {
            if (_txBody is null)
                return "";

            var parts = new List<string>();
            foreach (var pElem in _txBody.Elements(AP))
            {
                var pParts = new List<string>();
                foreach (var rElem in pElem.Elements(AR))
                {
                    var tElem = rElem.Element(AT);
                    if (tElem is not null && !string.IsNullOrEmpty(tElem.Value))
                        pParts.Add(tElem.Value);
                }
                parts.Add(string.Join("", pParts));
            }

            return string.Join("\n", parts);
        }
        set
        {
            if (_txBody is null)
                return;

            // Remove all existing paragraphs
            _txBody.Elements(AP).Remove();

            // Split on \r\n, \r, or \n to create separate paragraphs
            var lines = string.IsNullOrEmpty(value)
                ? [""]
                : Regex.Split(value, @"\r\n|\r|\n");

            foreach (var line in lines)
            {
                var pElem = new XElement(AP,
                    new XElement(AR,
                        new XElement(AT, line)));
                _txBody.Add(pElem);
            }

            _slidePart?.Save();
        }
    }

    /// <inheritdoc/>
    public ITextFrameFormat TextFrameFormat
    {
        get
        {
            var fmt = new TextFrameFormat();
            if (_txBody is not null)
                fmt.InitInternal(_txBody, _slidePart, _parentSlide);
            return fmt;
        }
    }

    /// <inheritdoc/>
    public IShape? ParentShape => _parentShape;

    /// <inheritdoc/>
    public ICell? ParentCell => _parentCell;

    /// <inheritdoc/>
    public override IBaseSlide? Slide => _parentSlide;

    /// <inheritdoc/>
    public override IPresentation? Presentation => _parentSlide?.Presentation;

    /// <inheritdoc/>
    public override IPresentationComponent AsIPresentationComponent => this;

    /// <summary>
    /// Returns this instance viewed as an <see cref="ISlideComponent"/>.
    /// </summary>
    public ISlideComponent AsISlideComponent => this;
}
