using System.Globalization;

namespace Aspose.Slides.Foss.Internal.Pptx.CommentsPart;

/// <summary>
/// Helpers for converting between OOXML datetime strings and <see cref="DateTime"/>.
/// </summary>
public static class DateTimeHelpers
{
    private static readonly string[] Formats =
    [
        "yyyy-MM-dd'T'HH:mm:ss.FFFFFFF",
        "yyyy-MM-dd'T'HH:mm:ss",
        "yyyy-MM-dd"
    ];

    /// <summary>
    /// Converts a <see cref="DateTime"/> to an OOXML datetime string.
    /// </summary>
    public static string DtToStr(DateTime dt)
    {
        int ms = dt.Millisecond;
        return dt.ToString($"yyyy-MM-dd'T'HH:mm:ss.{ms:D3}", CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Parses an OOXML datetime string to a <see cref="DateTime"/>,
    /// or returns <c>null</c> if the string is empty or unparseable.
    /// </summary>
    public static DateTime? StrToDt(string s)
    {
        if (string.IsNullOrEmpty(s))
            return null;

        if (DateTime.TryParseExact(s, Formats, CultureInfo.InvariantCulture,
                DateTimeStyles.None, out var result))
            return result;

        return null;
    }
}
