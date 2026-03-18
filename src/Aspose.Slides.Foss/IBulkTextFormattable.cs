namespace Aspose.Slides.Foss;

/// <summary>
/// Represents an object that can apply text formatting in bulk to all contained text.
/// </summary>
public interface IBulkTextFormattable
{
    /// <summary>
    /// Applies the specified portion format to all text within this object.
    /// </summary>
    /// <param name="source">The portion format to apply.</param>
    void SetTextFormat(IBasePortionFormat source);

    /// <summary>
    /// Applies the specified paragraph format to all paragraphs within this object.
    /// </summary>
    /// <param name="source">The paragraph format to apply.</param>
    void SetTextFormat(IParagraphFormat source);

    /// <summary>
    /// Applies the specified text frame format to all text frames within this object.
    /// </summary>
    /// <param name="source">The text frame format to apply.</param>
    void SetTextFormat(ITextFrameFormat source);
}
