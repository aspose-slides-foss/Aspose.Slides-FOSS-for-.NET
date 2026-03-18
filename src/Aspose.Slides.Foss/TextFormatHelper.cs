using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Applies text formatting to a list of cells.
/// </summary>
internal static class TextFormatHelper
{
    /// <summary>
    /// Applies portion formatting to every cell in the list.
    /// </summary>
    internal static void ApplyTextFormat(List<ICell> cells, IBasePortionFormat source, SlidePart? slidePart)
    {
        // Formatting application will be implemented when ITextFrame supports portion format operations.
        slidePart?.Save();
    }

    /// <summary>
    /// Applies paragraph formatting to every cell in the list.
    /// </summary>
    internal static void ApplyTextFormat(List<ICell> cells, IParagraphFormat source, SlidePart? slidePart)
    {
        // Formatting application will be implemented when ITextFrame supports paragraph format operations.
        slidePart?.Save();
    }

    /// <summary>
    /// Applies text frame formatting to every cell in the list.
    /// </summary>
    internal static void ApplyTextFormat(List<ICell> cells, ITextFrameFormat source, SlidePart? slidePart)
    {
        // Formatting application will be implemented when ITextFrame supports text frame format operations.
        slidePart?.Save();
    }
}
