using System.Security.Cryptography;
using Aspose.Slides.Foss.ConformanceTests.Harness;
using Aspose.Slides.Foss.Export;

namespace Aspose.Slides.Foss.ConformanceTests;

/// <summary>
/// What the caller asks <see cref="Presentation.Save(string, SaveFormat)"/> for has to be what lands
/// on disk. A save that ignores its format argument and writes a PPTX under another extension is the
/// most damaging kind of wrong: it returns success, and the file is refused by the application the
/// user named it for.
/// </summary>
public sealed class SaveFormatConformanceTests : IDisposable
{
    private const string PresentationContentType =
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";

    private readonly TestWorkspace _workspace = new();

    public void Dispose() => _workspace.Dispose();

    /// <summary>Formats that are not an OPC presentation package and that this library cannot render.</summary>
    public static TheoryData<SaveFormat> FormatsThatCannotBeWritten() =>
    [
        SaveFormat.Ppt, SaveFormat.Pdf, SaveFormat.Xps, SaveFormat.Tiff, SaveFormat.Odp,
        SaveFormat.Html, SaveFormat.Swf, SaveFormat.Otp, SaveFormat.Pps, SaveFormat.Pot,
        SaveFormat.Fodp, SaveFormat.Gif, SaveFormat.Html5, SaveFormat.Md, SaveFormat.Xml,
    ];

    [Theory]
    [MemberData(nameof(FormatsThatCannotBeWritten))]
    public void RequestingAFormatThatCannotBeWrittenRaisesInsteadOfWritingAPowerPointPackage(SaveFormat format)
    {
        var path = _workspace.PathFor($"deck.{format.ToString().ToLowerInvariant()}");

        using var presentation = new Presentation();
        var thrown = Record.Exception(() => presentation.Save(path, format));

        Assert.True(thrown is NotSupportedException,
            $"Save(path, SaveFormat.{format}) {(thrown is null ? "returned successfully" : $"threw {thrown.GetType().Name}")}, " +
            $"expected NotSupportedException. {DescribeWhatWasWritten(path)}");
    }

    [Theory]
    [InlineData(SaveFormat.Pptx, "pptx", PresentationContentType)]
    [InlineData(SaveFormat.Potx, "potx", "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml")]
    [InlineData(SaveFormat.Ppsx, "ppsx", "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml")]
    public void TheRequestedFormatDecidesTheMainPartContentType(SaveFormat format, string extension, string expected)
    {
        var path = _workspace.PathFor($"deck.{extension}");

        using (var presentation = new Presentation())
        {
            presentation.Save(path, format);
        }

        using var package = PptxPackage.Open(path);
        var actual = package.ContentTypeOf("ppt/presentation.xml");

        Assert.True(expected == actual,
            $"Saved as .{extension}, but /ppt/presentation.xml is declared " +
            $"'{actual ?? "nothing"}'{Environment.NewLine}expected '{expected}'.");
    }

    [Theory]
    [InlineData(SaveFormat.Pptm, "pptm", "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml")]
    [InlineData(SaveFormat.Ppsm, "ppsm", "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml")]
    [InlineData(SaveFormat.Potm, "potm", "application/vnd.ms-powerpoint.template.macroEnabled.main+xml")]
    public void MacroEnabledFormatsAreNotWrittenAsAPlainPresentation(SaveFormat format, string extension, string macroContentType)
    {
        var path = _workspace.PathFor($"deck.{extension}");

        using var presentation = new Presentation();
        var thrown = Record.Exception(() => presentation.Save(path, format));
        if (thrown is NotSupportedException)
            return; // Refusing to write a macro-enabled package it cannot honour is an acceptable answer.

        Assert.Null(thrown);

        using var package = PptxPackage.Open(path);
        var actual = package.ContentTypeOf("ppt/presentation.xml");

        Assert.True(macroContentType == actual,
            $"Saved as .{extension} without raising, and /ppt/presentation.xml is declared " +
            $"'{actual ?? "nothing"}'.{Environment.NewLine}" +
            $"A macro-enabled package must either raise NotSupportedException or declare '{macroContentType}'.");
    }

    [Fact]
    public void EachRequestedFormatProducesADifferentPackage()
    {
        var written = new Dictionary<string, string>(StringComparer.Ordinal);

        foreach (var format in Enum.GetValues<SaveFormat>())
        {
            var path = _workspace.PathFor($"distinct-{format}.{format.ToString().ToLowerInvariant()}");
            using var presentation = new Presentation();
            try
            {
                presentation.Save(path, format);
            }
            catch (NotSupportedException)
            {
                continue;
            }

            written[format.ToString()] = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)));
        }

        var distinct = written.Values.Distinct(StringComparer.Ordinal).Count();

        Assert.True(distinct == written.Count,
            $"{written.Count} formats were written and only {distinct} distinct file(s) came out — " +
            $"the format argument made no difference to the bytes.{Environment.NewLine}" +
            string.Join(Environment.NewLine, written.Select(pair => $"  {pair.Key,-6} {pair.Value[..16]}")));
    }

    private static string DescribeWhatWasWritten(string path)
    {
        if (!File.Exists(path))
            return "No file was written.";

        var length = new FileInfo(path).Length;
        try
        {
            using var package = PptxPackage.Open(path);
            var mainType = package.ContentTypeOf("ppt/presentation.xml");
            return $"It wrote a {length}-byte OPC package with {package.PartNames.Count} parts " +
                   $"whose /ppt/presentation.xml is '{mainType ?? "undeclared"}'.";
        }
        catch (Exception)
        {
            return $"It wrote {length} bytes.";
        }
    }
}
