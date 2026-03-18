namespace Aspose.Slides.Foss;

/// <summary>
/// Contains the paragraph formatting properties.
/// Unlike <see cref="IParagraph"/>, all properties of this class are writeable.
/// </summary>
public interface IParagraphFormat
{
    /// <summary>
    /// Returns bullet format of the paragraph. Read-only <see cref="IBulletFormat"/>.
    /// </summary>
    IBulletFormat Bullet { get; }

    /// <summary>
    /// Returns or sets depth of the paragraph.
    /// Value 0 means undefined value. Read/write.
    /// </summary>
    int Depth { get; set; }

    /// <summary>
    /// Returns or sets the text alignment in a paragraph with no inheritance.
    /// Read/write <see cref="TextAlignment"/>.
    /// </summary>
    TextAlignment Alignment { get; set; }

    /// <summary>
    /// Returns or sets the amount of space between base lines in a paragraph.
    /// Positive value means percentage, negative — size in points.
    /// No inheritance applied. Read/write.
    /// </summary>
    float SpaceWithin { get; set; }

    /// <summary>
    /// Returns or sets the amount of space before the first line in a paragraph with no inheritance.
    /// A positive value specifies the percentage of the font size that the white space should be.
    /// A negative value specifies the size of the white space in point size. Read/write.
    /// </summary>
    float SpaceBefore { get; set; }

    /// <summary>
    /// Returns or sets the amount of space after the last line in a paragraph with no inheritance.
    /// A positive value specifies the percentage of the font size that the white space should be.
    /// A negative value specifies the size of the white space in point size. Read/write.
    /// </summary>
    float SpaceAfter { get; set; }

    /// <summary>
    /// Determines whether the East Asian line break is used in a paragraph.
    /// No inheritance applied. Read/write <see cref="NullableBool"/>.
    /// </summary>
    NullableBool EastAsianLineBreak { get; set; }

    /// <summary>
    /// Determines whether the Right to Left writing is used in a paragraph.
    /// No inheritance applied. Read/write <see cref="NullableBool"/>.
    /// </summary>
    NullableBool RightToLeft { get; set; }

    /// <summary>
    /// Determines whether the Latin line break is used in a paragraph.
    /// No inheritance applied. Read/write <see cref="NullableBool"/>.
    /// </summary>
    NullableBool LatinLineBreak { get; set; }

    /// <summary>
    /// Determines whether the hanging punctuation is used in a paragraph.
    /// No inheritance applied. Read/write <see cref="NullableBool"/>.
    /// </summary>
    NullableBool HangingPunctuation { get; set; }

    /// <summary>
    /// Returns or sets the left margin in a paragraph with no inheritance. Read/write.
    /// </summary>
    float MarginLeft { get; set; }

    /// <summary>
    /// Returns or sets the right margin in a paragraph with no inheritance. Read/write.
    /// </summary>
    float MarginRight { get; set; }

    /// <summary>
    /// Returns or sets paragraph First Line Indent/Hanging Indent with no inheritance.
    /// Hanging Indent can be defined with negative values. Read/write.
    /// </summary>
    float Indent { get; set; }

    /// <summary>
    /// Returns or sets default tabulation size with no inheritance. Read/write.
    /// </summary>
    float DefaultTabSize { get; set; }

    /// <summary>
    /// Returns or sets a font alignment in a paragraph with no inheritance.
    /// Read/write <see cref="FontAlignment"/>.
    /// </summary>
    FontAlignment FontAlignment { get; set; }

    /// <summary>
    /// Returns default portion format of a paragraph.
    /// No inheritance applied. Read-only <see cref="PortionFormat"/>.
    /// </summary>
    PortionFormat DefaultPortionFormat { get; }
}
