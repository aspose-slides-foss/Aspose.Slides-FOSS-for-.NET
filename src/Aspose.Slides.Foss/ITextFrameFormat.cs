namespace Aspose.Slides.Foss;

/// <summary>
/// Contains the TextFrame's formatting properties.
/// </summary>
public interface ITextFrameFormat
{
    /// <summary>
    /// Returns or sets the left margin (points) in a TextFrame.
    /// </summary>
    float MarginLeft { get; set; }

    /// <summary>
    /// Returns or sets the right margin (points) in a TextFrame.
    /// </summary>
    float MarginRight { get; set; }

    /// <summary>
    /// Returns or sets the top margin (points) in a TextFrame.
    /// </summary>
    float MarginTop { get; set; }

    /// <summary>
    /// Returns or sets the bottom margin (points) in a TextFrame.
    /// </summary>
    float MarginBottom { get; set; }

    /// <summary>
    /// True if text is wrapped at TextFrame's margins.
    /// </summary>
    NullableBool WrapText { get; set; }

    /// <summary>
    /// Returns or sets the text anchoring type.
    /// </summary>
    TextAnchorType AnchoringType { get; set; }

    /// <summary>
    /// Returns or sets whether text should be centered in the box horizontally.
    /// </summary>
    NullableBool CenterText { get; set; }

    /// <summary>
    /// Returns or sets the text vertical orientation type.
    /// </summary>
    TextVerticalType TextVerticalType { get; set; }

    /// <summary>
    /// Returns or sets the text autofit type.
    /// </summary>
    TextAutofitType AutofitType { get; set; }

    /// <summary>
    /// Returns or sets the number of columns in the text area.
    /// </summary>
    int ColumnCount { get; set; }

    /// <summary>
    /// Returns or sets the spacing between text columns in the text area (in points).
    /// </summary>
    float ColumnSpacing { get; set; }

    /// <summary>
    /// Returns or sets the rotation angle applied to text within the bounding box.
    /// </summary>
    float RotationAngle { get; set; }

    /// <summary>
    /// Returns or sets the text wrapping shape (preset text warp).
    /// </summary>
    TextShapeType Transform { get; set; }

    /// <summary>
    /// Returns or sets keeping text out of 3D scene entirely.
    /// </summary>
    bool KeepTextFlat { get; set; }

    /// <summary>
    /// Returns the <see cref="IThreeDFormat"/> for the text body.
    /// </summary>
    IThreeDFormat ThreeDFormat { get; }

    /// <summary>
    /// Returns the parent slide that owns this text frame format.
    /// </summary>
    IBaseSlide? Slide { get; }

    /// <summary>
    /// Returns the presentation that contains the text frame.
    /// </summary>
    IPresentation? Presentation { get; }

    /// <summary>
    /// Returns this instance as an <see cref="IPresentationComponent"/>.
    /// </summary>
    IPresentationComponent AsIPresentationComponent { get; }
}
