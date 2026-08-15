using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace Aspose.Slides.Foss.ConformanceTests.Harness;

/// <summary>
/// Schema validation of a produced package against ECMA-376, using the Open XML SDK's validator.
/// <para>
/// This catches what a hand-written package assertion will not: invented element and attribute names,
/// missing required attributes, and children in the wrong order. It is a strict reader — stricter
/// than PowerPoint, which silently discards markup it does not understand, so a file can validate
/// badly and still open.
/// </para>
/// </summary>
internal static class SchemaValidation
{
    /// <summary>How many validation errors a failure message prints before it truncates.</summary>
    private const int MaxReportedErrors = 12;

    /// <summary>
    /// Asserts that the package at <paramref name="path"/> produces zero validation errors against
    /// the Office 2019 schema set.
    /// </summary>
    internal static void HasNoSchemaErrors(string path)
    {
        var errors = Validate(path);

        Assert.True(errors.Count == 0,
            $"{Path.GetFileName(path)} has {errors.Count} schema validation error(s):{Environment.NewLine}" +
            string.Join(Environment.NewLine, errors.Take(MaxReportedErrors)) +
            (errors.Count > MaxReportedErrors ? $"{Environment.NewLine}… and {errors.Count - MaxReportedErrors} more" : string.Empty));
    }

    /// <summary>
    /// Validates the package and returns one formatted line per error. A package the SDK refuses to
    /// open at all is reported as a single error carrying the reason, because refusal by a strict
    /// consumer is exactly the outcome under test.
    /// </summary>
    internal static IReadOnlyList<string> Validate(string path)
    {
        var validator = new OpenXmlValidator(FileFormatVersions.Office2019);

        try
        {
            using var document = PresentationDocument.Open(path, isEditable: false);
            return validator.Validate(document)
                .Select(Format)
                .ToList();
        }
        catch (Exception ex) when (ex is OpenXmlPackageException or FileFormatException or InvalidDataException)
        {
            return [$"[package] {ex.GetType().Name}: {ex.Message.Trim()}"];
        }
    }

    private static string Format(ValidationErrorInfo error) =>
        $"[{error.ErrorType}] {error.Description}{Environment.NewLine}    at {error.Path?.XPath ?? "(unknown)"} in {error.Part?.Uri.ToString() ?? "(unknown part)"}";
}
