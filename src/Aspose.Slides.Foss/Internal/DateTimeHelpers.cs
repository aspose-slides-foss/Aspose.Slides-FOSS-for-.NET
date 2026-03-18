using System.Globalization;

namespace Aspose.Slides.Foss.Internal;

/// <summary>
/// Helpers for converting between OOXML datetime strings and <see cref="DateTime"/>.
/// </summary>
internal static class DateTimeHelpers
{
    private static readonly string[] Formats =
    [
        "yyyy-MM-dd'T'HH:mm:ss.FFFFFFF",
        "yyyy-MM-dd'T'HH:mm:ss",
        "yyyy-MM-dd"
    ];

    /// <summary>
    /// Parses an OOXML datetime string to a <see cref="DateTime"/>, or returns <c>null</c> if empty or unparseable.
    /// </summary>
    internal static DateTime? StrToDt(string s)
    {
        if (string.IsNullOrEmpty(s))
            return null;

        if (DateTime.TryParseExact(s, Formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var result))
            return result;

        return null;
    }

    /// <summary>
    /// Converts a <see cref="DateTime"/> to an OOXML datetime string.
    /// </summary>
    internal static string DtToStr(DateTime dt)
    {
        int ms = dt.Millisecond;
        return dt.ToString($"yyyy-MM-dd'T'HH:mm:ss.{ms:D3}", CultureInfo.InvariantCulture);
    }
}
