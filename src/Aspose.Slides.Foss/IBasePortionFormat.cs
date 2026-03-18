namespace Aspose.Slides.Foss;

/// <summary>
/// Defines common text run formatting properties.
/// </summary>
public interface IBasePortionFormat
{
    /// <summary>
    /// Gets the line format for text outline.
    /// </summary>
    ILineFormat? LineFormat { get; }

    /// <summary>
    /// Gets the fill format for text.
    /// </summary>
    IFillFormat? FillFormat { get; }

    /// <summary>
    /// Gets the effect format for text.
    /// </summary>
    IEffectFormat? EffectFormat { get; }

    /// <summary>
    /// Gets the highlight color format.
    /// </summary>
    IColorFormat? HighlightColor { get; }

    /// <summary>
    /// Gets the underline line format.
    /// </summary>
    ILineFormat? UnderlineLineFormat { get; }

    /// <summary>
    /// Gets the underline fill format.
    /// </summary>
    IFillFormat? UnderlineFillFormat { get; }

    /// <summary>
    /// Gets or sets whether text is bold.
    /// </summary>
    NullableBool FontBold { get; set; }

    /// <summary>
    /// Gets or sets whether text is italic.
    /// </summary>
    NullableBool FontItalic { get; set; }

    /// <summary>
    /// Gets or sets kumimoji rendering.
    /// </summary>
    NullableBool Kumimoji { get; set; }

    /// <summary>
    /// Gets or sets whether text height is normalized.
    /// </summary>
    NullableBool NormaliseHeight { get; set; }

    /// <summary>
    /// Gets or sets whether proofing is disabled.
    /// </summary>
    NullableBool ProofDisabled { get; set; }

    /// <summary>
    /// Gets or sets the font underline type.
    /// </summary>
    TextUnderlineType FontUnderline { get; set; }

    /// <summary>
    /// Gets or sets the text capitalization type.
    /// </summary>
    TextCapType TextCapType { get; set; }

    /// <summary>
    /// Gets or sets the strikethrough type.
    /// </summary>
    TextStrikethroughType StrikethroughType { get; set; }

    /// <summary>
    /// Gets or sets whether an explicit underline line element is present.
    /// </summary>
    NullableBool IsHardUnderlineLine { get; set; }

    /// <summary>
    /// Gets or sets whether an explicit underline fill element is present.
    /// </summary>
    NullableBool IsHardUnderlineFill { get; set; }

    /// <summary>
    /// Gets or sets the font height in points.
    /// </summary>
    float FontHeight { get; set; }

    /// <summary>
    /// Gets or sets the escapement (superscript/subscript) as a fraction.
    /// </summary>
    float Escapement { get; set; }

    /// <summary>
    /// Gets or sets the minimum font size for kerning, in points.
    /// </summary>
    float KerningMinimalSize { get; set; }

    /// <summary>
    /// Gets or sets the character spacing in points.
    /// </summary>
    float Spacing { get; set; }

    /// <summary>
    /// Gets or sets the Latin script font.
    /// </summary>
    IFontData? LatinFont { get; set; }

    /// <summary>
    /// Gets or sets the East Asian script font.
    /// </summary>
    IFontData? EastAsianFont { get; set; }

    /// <summary>
    /// Gets or sets the complex script font.
    /// </summary>
    IFontData? ComplexScriptFont { get; set; }

    /// <summary>
    /// Gets or sets the symbol font.
    /// </summary>
    IFontData? SymbolFont { get; set; }

    /// <summary>
    /// Gets or sets the language identifier (e.g. "en-US").
    /// </summary>
    string? LanguageId { get; set; }

    /// <summary>
    /// Gets or sets the alternative language identifier.
    /// </summary>
    string? AlternativeLanguageId { get; set; }

    /// <summary>
    /// Gets or sets the spell check flag.
    /// </summary>
    NullableBool SpellCheck { get; set; }
}
