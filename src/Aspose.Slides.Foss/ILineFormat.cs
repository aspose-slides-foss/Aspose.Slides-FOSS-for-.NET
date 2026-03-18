namespace Aspose.Slides.Foss;

/// <summary>
/// Represents format of a line.
/// </summary>
public interface ILineFormat
{
    /// <summary>
    /// Returns true if line format is not defined (as just created, default). Read-only.
    /// </summary>
    bool IsFormatNotDefined { get; }

    /// <summary>
    /// Returns the fill format of a line. Read-only.
    /// </summary>
    ILineFillFormat FillFormat { get; }

    /// <summary>
    /// Returns or sets the width of a line. Read/write.
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// Returns or sets the line dash style. Read/write.
    /// </summary>
    LineDashStyle DashStyle { get; set; }

    /// <summary>
    /// Returns or sets the line cap style. Read/write.
    /// </summary>
    LineCapStyle CapStyle { get; set; }

    /// <summary>
    /// Returns or sets the line alignment. Read/write.
    /// </summary>
    LineAlignment Alignment { get; set; }

    /// <summary>
    /// Returns or sets the arrowhead style at the beginning of a line. Read/write.
    /// </summary>
    LineArrowheadStyle BeginArrowheadStyle { get; set; }

    /// <summary>
    /// Returns or sets the arrowhead style at the end of a line. Read/write.
    /// </summary>
    LineArrowheadStyle EndArrowheadStyle { get; set; }

    /// <summary>
    /// Returns or sets the arrowhead width at the beginning of a line. Read/write.
    /// </summary>
    LineArrowheadWidth BeginArrowheadWidth { get; set; }

    /// <summary>
    /// Returns or sets the arrowhead width at the end of a line. Read/write.
    /// </summary>
    LineArrowheadWidth EndArrowheadWidth { get; set; }

    /// <summary>
    /// Returns or sets the arrowhead length at the beginning of a line. Read/write.
    /// </summary>
    LineArrowheadLength BeginArrowheadLength { get; set; }

    /// <summary>
    /// Returns or sets the arrowhead length at the end of a line. Read/write.
    /// </summary>
    LineArrowheadLength EndArrowheadLength { get; set; }

    /// <summary>
    /// Returns or sets the custom dash pattern. Read/write.
    /// </summary>
    IList<float> CustomDashPattern { get; set; }

    /// <summary>
    /// Returns or sets the line compound style. Read/write.
    /// </summary>
    LineStyle Style { get; set; }

    /// <summary>
    /// Returns or sets the line join style. Read/write.
    /// </summary>
    LineJoinStyle JoinStyle { get; set; }

    /// <summary>
    /// Returns or sets the miter limit. Read/write.
    /// </summary>
    float MiterLimit { get; set; }
}
