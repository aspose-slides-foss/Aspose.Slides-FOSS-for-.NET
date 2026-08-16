namespace Aspose.Slides.Foss.Drawing;

/// <summary>
/// An immutable ARGB color. This is a reference type, but it has value semantics: two instances
/// with the same components are equal, and <c>==</c>, <c>!=</c>, <see cref="Equals(object?)"/> and
/// <see cref="GetHashCode"/> all compare the components rather than the reference.
/// </summary>
public sealed class Color
{
    private readonly int _a;
    private readonly int _r;
    private readonly int _g;
    private readonly int _b;

    /// <summary>
    /// Gets the alpha component.
    /// </summary>
    public int A => _a;

    /// <summary>
    /// Gets the red component.
    /// </summary>
    public int R => _r;

    /// <summary>
    /// Gets the green component.
    /// </summary>
    public int G => _g;

    /// <summary>
    /// Gets the blue component.
    /// </summary>
    public int B => _b;

    /// <summary>
    /// Initializes a new instance of <see cref="Color"/> with the specified ARGB components.
    /// </summary>
    public Color(int a = 255, int r = 0, int g = 0, int b = 0)
    {
        _a = a;
        _r = r;
        _g = g;
        _b = b;
    }

    /// <summary>
    /// Creates a new <see cref="Color"/> from the specified ARGB components.
    /// </summary>
    public static Color FromArgb(int a, int r, int g, int b) => new(a, r, g, b);

    /// <inheritdoc/>
    public override bool Equals(object? obj) =>
        obj is Color other && _a == other._a && _r == other._r && _g == other._g && _b == other._b;

    /// <inheritdoc/>
    public override int GetHashCode() => HashCode.Combine(_a, _r, _g, _b);

    /// <summary>
    /// Determines whether two colors have the same components. Two <c>null</c> references are
    /// equal; a <c>null</c> reference is not equal to any color.
    /// </summary>
    /// <param name="left">The first color to compare, which may be <c>null</c>.</param>
    /// <param name="right">The second color to compare, which may be <c>null</c>.</param>
    /// <returns><c>true</c> if the two are equal; otherwise <c>false</c>.</returns>
    public static bool operator ==(Color? left, Color? right) =>
        left is null ? right is null : left.Equals(right);

    /// <summary>
    /// Determines whether two colors differ in any component.
    /// </summary>
    /// <param name="left">The first color to compare, which may be <c>null</c>.</param>
    /// <param name="right">The second color to compare, which may be <c>null</c>.</param>
    /// <returns><c>true</c> if the two are not equal; otherwise <c>false</c>.</returns>
    public static bool operator !=(Color? left, Color? right) => !(left == right);

    /// <inheritdoc/>
    public override string ToString() =>
        _a == 255
            ? $"Color(r={_r}, g={_g}, b={_b})"
            : $"Color(a={_a}, r={_r}, g={_g}, b={_b})";

    // Named color constants
    /// <summary>Transparent white — alpha zero over white, ARGB <c>#00FFFFFF</c>.</summary>
    public static readonly Color Transparent = new(a: 0, r: 255, g: 255, b: 255);
    /// <summary>The color alice blue, ARGB <c>#FFF0F8FF</c>.</summary>
    public static readonly Color AliceBlue = new(r: 240, g: 248, b: 255);
    /// <summary>The color antique white, ARGB <c>#FFFAEBD7</c>.</summary>
    public static readonly Color AntiqueWhite = new(r: 250, g: 235, b: 215);
    /// <summary>The color aqua, ARGB <c>#FF00FFFF</c>.</summary>
    public static readonly Color Aqua = new(r: 0, g: 255, b: 255);
    /// <summary>The color aquamarine, ARGB <c>#FF7FFFD4</c>.</summary>
    public static readonly Color Aquamarine = new(r: 127, g: 255, b: 212);
    /// <summary>The color azure, ARGB <c>#FFF0FFFF</c>.</summary>
    public static readonly Color Azure = new(r: 240, g: 255, b: 255);
    /// <summary>The color beige, ARGB <c>#FFF5F5DC</c>.</summary>
    public static readonly Color Beige = new(r: 245, g: 245, b: 220);
    /// <summary>The color bisque, ARGB <c>#FFFFE4C4</c>.</summary>
    public static readonly Color Bisque = new(r: 255, g: 228, b: 196);
    /// <summary>The color black, ARGB <c>#FF000000</c>.</summary>
    public static readonly Color Black = new(r: 0, g: 0, b: 0);
    /// <summary>The color blanched almond, ARGB <c>#FFFFEBCD</c>.</summary>
    public static readonly Color BlanchedAlmond = new(r: 255, g: 235, b: 205);
    /// <summary>The color blue, ARGB <c>#FF0000FF</c>.</summary>
    public static readonly Color Blue = new(r: 0, g: 0, b: 255);
    /// <summary>The color blue violet, ARGB <c>#FF8A2BE2</c>.</summary>
    public static readonly Color BlueViolet = new(r: 138, g: 43, b: 226);
    /// <summary>The color brown, ARGB <c>#FFA52A2A</c>.</summary>
    public static readonly Color Brown = new(r: 165, g: 42, b: 42);
    /// <summary>The color burly wood, ARGB <c>#FFDEB887</c>.</summary>
    public static readonly Color BurlyWood = new(r: 222, g: 184, b: 135);
    /// <summary>The color cadet blue, ARGB <c>#FF5F9EA0</c>.</summary>
    public static readonly Color CadetBlue = new(r: 95, g: 158, b: 160);
    /// <summary>The color chartreuse, ARGB <c>#FF7FFF00</c>.</summary>
    public static readonly Color Chartreuse = new(r: 127, g: 255, b: 0);
    /// <summary>The color chocolate, ARGB <c>#FFD2691E</c>.</summary>
    public static readonly Color Chocolate = new(r: 210, g: 105, b: 30);
    /// <summary>The color coral, ARGB <c>#FFFF7F50</c>.</summary>
    public static readonly Color Coral = new(r: 255, g: 127, b: 80);
    /// <summary>The color cornflower blue, ARGB <c>#FF6495ED</c>.</summary>
    public static readonly Color CornflowerBlue = new(r: 100, g: 149, b: 237);
    /// <summary>The color cornsilk, ARGB <c>#FFFFF8DC</c>.</summary>
    public static readonly Color Cornsilk = new(r: 255, g: 248, b: 220);
    /// <summary>The color crimson, ARGB <c>#FFDC143C</c>.</summary>
    public static readonly Color Crimson = new(r: 220, g: 20, b: 60);
    /// <summary>The color cyan, ARGB <c>#FF00FFFF</c>.</summary>
    public static readonly Color Cyan = new(r: 0, g: 255, b: 255);
    /// <summary>The color dark blue, ARGB <c>#FF00008B</c>.</summary>
    public static readonly Color DarkBlue = new(r: 0, g: 0, b: 139);
    /// <summary>The color dark cyan, ARGB <c>#FF008B8B</c>.</summary>
    public static readonly Color DarkCyan = new(r: 0, g: 139, b: 139);
    /// <summary>The color dark goldenrod, ARGB <c>#FFB8860B</c>.</summary>
    public static readonly Color DarkGoldenrod = new(r: 184, g: 134, b: 11);
    /// <summary>The color dark gray, ARGB <c>#FFA9A9A9</c>.</summary>
    public static readonly Color DarkGray = new(r: 169, g: 169, b: 169);
    /// <summary>The color dark green, ARGB <c>#FF006400</c>.</summary>
    public static readonly Color DarkGreen = new(r: 0, g: 100, b: 0);
    /// <summary>The color dark khaki, ARGB <c>#FFBDB76B</c>.</summary>
    public static readonly Color DarkKhaki = new(r: 189, g: 183, b: 107);
    /// <summary>The color dark magenta, ARGB <c>#FF8B008B</c>.</summary>
    public static readonly Color DarkMagenta = new(r: 139, g: 0, b: 139);
    /// <summary>The color dark olive green, ARGB <c>#FF556B2F</c>.</summary>
    public static readonly Color DarkOliveGreen = new(r: 85, g: 107, b: 47);
    /// <summary>The color dark orange, ARGB <c>#FFFF8C00</c>.</summary>
    public static readonly Color DarkOrange = new(r: 255, g: 140, b: 0);
    /// <summary>The color dark orchid, ARGB <c>#FF9932CC</c>.</summary>
    public static readonly Color DarkOrchid = new(r: 153, g: 50, b: 204);
    /// <summary>The color dark red, ARGB <c>#FF8B0000</c>.</summary>
    public static readonly Color DarkRed = new(r: 139, g: 0, b: 0);
    /// <summary>The color dark salmon, ARGB <c>#FFE9967A</c>.</summary>
    public static readonly Color DarkSalmon = new(r: 233, g: 150, b: 122);
    /// <summary>The color dark sea green, ARGB <c>#FF8FBC8F</c>.</summary>
    public static readonly Color DarkSeaGreen = new(r: 143, g: 188, b: 143);
    /// <summary>The color dark slate blue, ARGB <c>#FF483D8B</c>.</summary>
    public static readonly Color DarkSlateBlue = new(r: 72, g: 61, b: 139);
    /// <summary>The color dark slate gray, ARGB <c>#FF2F4F4F</c>.</summary>
    public static readonly Color DarkSlateGray = new(r: 47, g: 79, b: 79);
    /// <summary>The color dark turquoise, ARGB <c>#FF00CED1</c>.</summary>
    public static readonly Color DarkTurquoise = new(r: 0, g: 206, b: 209);
    /// <summary>The color dark violet, ARGB <c>#FF9400D3</c>.</summary>
    public static readonly Color DarkViolet = new(r: 148, g: 0, b: 211);
    /// <summary>The color deep pink, ARGB <c>#FFFF1493</c>.</summary>
    public static readonly Color DeepPink = new(r: 255, g: 20, b: 147);
    /// <summary>The color deep sky blue, ARGB <c>#FF00BFFF</c>.</summary>
    public static readonly Color DeepSkyBlue = new(r: 0, g: 191, b: 255);
    /// <summary>The color dim gray, ARGB <c>#FF696969</c>.</summary>
    public static readonly Color DimGray = new(r: 105, g: 105, b: 105);
    /// <summary>The color dodger blue, ARGB <c>#FF1E90FF</c>.</summary>
    public static readonly Color DodgerBlue = new(r: 30, g: 144, b: 255);
    /// <summary>The color firebrick, ARGB <c>#FFB22222</c>.</summary>
    public static readonly Color Firebrick = new(r: 178, g: 34, b: 34);
    /// <summary>The color floral white, ARGB <c>#FFFFFAF0</c>.</summary>
    public static readonly Color FloralWhite = new(r: 255, g: 250, b: 240);
    /// <summary>The color forest green, ARGB <c>#FF228B22</c>.</summary>
    public static readonly Color ForestGreen = new(r: 34, g: 139, b: 34);
    /// <summary>The color fuchsia, ARGB <c>#FFFF00FF</c>.</summary>
    public static readonly Color Fuchsia = new(r: 255, g: 0, b: 255);
    /// <summary>The color gainsboro, ARGB <c>#FFDCDCDC</c>.</summary>
    public static readonly Color Gainsboro = new(r: 220, g: 220, b: 220);
    /// <summary>The color ghost white, ARGB <c>#FFF8F8FF</c>.</summary>
    public static readonly Color GhostWhite = new(r: 248, g: 248, b: 255);
    /// <summary>The color gold, ARGB <c>#FFFFD700</c>.</summary>
    public static readonly Color Gold = new(r: 255, g: 215, b: 0);
    /// <summary>The color goldenrod, ARGB <c>#FFDAA520</c>.</summary>
    public static readonly Color Goldenrod = new(r: 218, g: 165, b: 32);
    /// <summary>The color gray, ARGB <c>#FF808080</c>.</summary>
    public static readonly Color Gray = new(r: 128, g: 128, b: 128);
    /// <summary>The color green, ARGB <c>#FF008000</c>.</summary>
    public static readonly Color Green = new(r: 0, g: 128, b: 0);
    /// <summary>The color green yellow, ARGB <c>#FFADFF2F</c>.</summary>
    public static readonly Color GreenYellow = new(r: 173, g: 255, b: 47);
    /// <summary>The color honeydew, ARGB <c>#FFF0FFF0</c>.</summary>
    public static readonly Color Honeydew = new(r: 240, g: 255, b: 240);
    /// <summary>The color hot pink, ARGB <c>#FFFF69B4</c>.</summary>
    public static readonly Color HotPink = new(r: 255, g: 105, b: 180);
    /// <summary>The color indian red, ARGB <c>#FFCD5C5C</c>.</summary>
    public static readonly Color IndianRed = new(r: 205, g: 92, b: 92);
    /// <summary>The color indigo, ARGB <c>#FF4B0082</c>.</summary>
    public static readonly Color Indigo = new(r: 75, g: 0, b: 130);
    /// <summary>The color ivory, ARGB <c>#FFFFFFF0</c>.</summary>
    public static readonly Color Ivory = new(r: 255, g: 255, b: 240);
    /// <summary>The color khaki, ARGB <c>#FFF0E68C</c>.</summary>
    public static readonly Color Khaki = new(r: 240, g: 230, b: 140);
    /// <summary>The color lavender, ARGB <c>#FFE6E6FA</c>.</summary>
    public static readonly Color Lavender = new(r: 230, g: 230, b: 250);
    /// <summary>The color lavender blush, ARGB <c>#FFFFF0F5</c>.</summary>
    public static readonly Color LavenderBlush = new(r: 255, g: 240, b: 245);
    /// <summary>The color lawn green, ARGB <c>#FF7CFC00</c>.</summary>
    public static readonly Color LawnGreen = new(r: 124, g: 252, b: 0);
    /// <summary>The color lemon chiffon, ARGB <c>#FFFFFACD</c>.</summary>
    public static readonly Color LemonChiffon = new(r: 255, g: 250, b: 205);
    /// <summary>The color light blue, ARGB <c>#FFADD8E6</c>.</summary>
    public static readonly Color LightBlue = new(r: 173, g: 216, b: 230);
    /// <summary>The color light coral, ARGB <c>#FFF08080</c>.</summary>
    public static readonly Color LightCoral = new(r: 240, g: 128, b: 128);
    /// <summary>The color light cyan, ARGB <c>#FFE0FFFF</c>.</summary>
    public static readonly Color LightCyan = new(r: 224, g: 255, b: 255);
    /// <summary>The color light goldenrod yellow, ARGB <c>#FFFAFAD2</c>.</summary>
    public static readonly Color LightGoldenrodYellow = new(r: 250, g: 250, b: 210);
    /// <summary>The color light gray, ARGB <c>#FFD3D3D3</c>.</summary>
    public static readonly Color LightGray = new(r: 211, g: 211, b: 211);
    /// <summary>The color light green, ARGB <c>#FF90EE90</c>.</summary>
    public static readonly Color LightGreen = new(r: 144, g: 238, b: 144);
    /// <summary>The color light pink, ARGB <c>#FFFFB6C1</c>.</summary>
    public static readonly Color LightPink = new(r: 255, g: 182, b: 193);
    /// <summary>The color light salmon, ARGB <c>#FFFFA07A</c>.</summary>
    public static readonly Color LightSalmon = new(r: 255, g: 160, b: 122);
    /// <summary>The color light sea green, ARGB <c>#FF20B2AA</c>.</summary>
    public static readonly Color LightSeaGreen = new(r: 32, g: 178, b: 170);
    /// <summary>The color light sky blue, ARGB <c>#FF87CEFA</c>.</summary>
    public static readonly Color LightSkyBlue = new(r: 135, g: 206, b: 250);
    /// <summary>The color light slate gray, ARGB <c>#FF778899</c>.</summary>
    public static readonly Color LightSlateGray = new(r: 119, g: 136, b: 153);
    /// <summary>The color light steel blue, ARGB <c>#FFB0C4DE</c>.</summary>
    public static readonly Color LightSteelBlue = new(r: 176, g: 196, b: 222);
    /// <summary>The color light yellow, ARGB <c>#FFFFFFE0</c>.</summary>
    public static readonly Color LightYellow = new(r: 255, g: 255, b: 224);
    /// <summary>The color lime, ARGB <c>#FF00FF00</c>.</summary>
    public static readonly Color Lime = new(r: 0, g: 255, b: 0);
    /// <summary>The color lime green, ARGB <c>#FF32CD32</c>.</summary>
    public static readonly Color LimeGreen = new(r: 50, g: 205, b: 50);
    /// <summary>The color linen, ARGB <c>#FFFAF0E6</c>.</summary>
    public static readonly Color Linen = new(r: 250, g: 240, b: 230);
    /// <summary>The color magenta, ARGB <c>#FFFF00FF</c>.</summary>
    public static readonly Color Magenta = new(r: 255, g: 0, b: 255);
    /// <summary>The color maroon, ARGB <c>#FF800000</c>.</summary>
    public static readonly Color Maroon = new(r: 128, g: 0, b: 0);
    /// <summary>The color medium aquamarine, ARGB <c>#FF66CDAA</c>.</summary>
    public static readonly Color MediumAquamarine = new(r: 102, g: 205, b: 170);
    /// <summary>The color medium blue, ARGB <c>#FF0000CD</c>.</summary>
    public static readonly Color MediumBlue = new(r: 0, g: 0, b: 205);
    /// <summary>The color medium orchid, ARGB <c>#FFBA55D3</c>.</summary>
    public static readonly Color MediumOrchid = new(r: 186, g: 85, b: 211);
    /// <summary>The color medium purple, ARGB <c>#FF9370DB</c>.</summary>
    public static readonly Color MediumPurple = new(r: 147, g: 112, b: 219);
    /// <summary>The color medium sea green, ARGB <c>#FF3CB371</c>.</summary>
    public static readonly Color MediumSeaGreen = new(r: 60, g: 179, b: 113);
    /// <summary>The color medium slate blue, ARGB <c>#FF7B68EE</c>.</summary>
    public static readonly Color MediumSlateBlue = new(r: 123, g: 104, b: 238);
    /// <summary>The color medium spring green, ARGB <c>#FF00FA9A</c>.</summary>
    public static readonly Color MediumSpringGreen = new(r: 0, g: 250, b: 154);
    /// <summary>The color medium turquoise, ARGB <c>#FF48D1CC</c>.</summary>
    public static readonly Color MediumTurquoise = new(r: 72, g: 209, b: 204);
    /// <summary>The color medium violet red, ARGB <c>#FFC71585</c>.</summary>
    public static readonly Color MediumVioletRed = new(r: 199, g: 21, b: 133);
    /// <summary>The color midnight blue, ARGB <c>#FF191970</c>.</summary>
    public static readonly Color MidnightBlue = new(r: 25, g: 25, b: 112);
    /// <summary>The color mint cream, ARGB <c>#FFF5FFFA</c>.</summary>
    public static readonly Color MintCream = new(r: 245, g: 255, b: 250);
    /// <summary>The color misty rose, ARGB <c>#FFFFE4E1</c>.</summary>
    public static readonly Color MistyRose = new(r: 255, g: 228, b: 225);
    /// <summary>The color moccasin, ARGB <c>#FFFFE4B5</c>.</summary>
    public static readonly Color Moccasin = new(r: 255, g: 228, b: 181);
    /// <summary>The color navajo white, ARGB <c>#FFFFDEAD</c>.</summary>
    public static readonly Color NavajoWhite = new(r: 255, g: 222, b: 173);
    /// <summary>The color navy, ARGB <c>#FF000080</c>.</summary>
    public static readonly Color Navy = new(r: 0, g: 0, b: 128);
    /// <summary>The color old lace, ARGB <c>#FFFDF5E6</c>.</summary>
    public static readonly Color OldLace = new(r: 253, g: 245, b: 230);
    /// <summary>The color olive, ARGB <c>#FF808000</c>.</summary>
    public static readonly Color Olive = new(r: 128, g: 128, b: 0);
    /// <summary>The color olive drab, ARGB <c>#FF6B8E23</c>.</summary>
    public static readonly Color OliveDrab = new(r: 107, g: 142, b: 35);
    /// <summary>The color orange, ARGB <c>#FFFFA500</c>.</summary>
    public static readonly Color Orange = new(r: 255, g: 165, b: 0);
    /// <summary>The color orange red, ARGB <c>#FFFF4500</c>.</summary>
    public static readonly Color OrangeRed = new(r: 255, g: 69, b: 0);
    /// <summary>The color orchid, ARGB <c>#FFDA70D6</c>.</summary>
    public static readonly Color Orchid = new(r: 218, g: 112, b: 214);
    /// <summary>The color pale goldenrod, ARGB <c>#FFEEE8AA</c>.</summary>
    public static readonly Color PaleGoldenrod = new(r: 238, g: 232, b: 170);
    /// <summary>The color pale green, ARGB <c>#FF98FB98</c>.</summary>
    public static readonly Color PaleGreen = new(r: 152, g: 251, b: 152);
    /// <summary>The color pale turquoise, ARGB <c>#FFAFEEEE</c>.</summary>
    public static readonly Color PaleTurquoise = new(r: 175, g: 238, b: 238);
    /// <summary>The color pale violet red, ARGB <c>#FFDB7093</c>.</summary>
    public static readonly Color PaleVioletRed = new(r: 219, g: 112, b: 147);
    /// <summary>The color papaya whip, ARGB <c>#FFFFEFD5</c>.</summary>
    public static readonly Color PapayaWhip = new(r: 255, g: 239, b: 213);
    /// <summary>The color peach puff, ARGB <c>#FFFFDAB9</c>.</summary>
    public static readonly Color PeachPuff = new(r: 255, g: 218, b: 185);
    /// <summary>The color peru, ARGB <c>#FFCD853F</c>.</summary>
    public static readonly Color Peru = new(r: 205, g: 133, b: 63);
    /// <summary>The color pink, ARGB <c>#FFFFC0CB</c>.</summary>
    public static readonly Color Pink = new(r: 255, g: 192, b: 203);
    /// <summary>The color plum, ARGB <c>#FFDDA0DD</c>.</summary>
    public static readonly Color Plum = new(r: 221, g: 160, b: 221);
    /// <summary>The color powder blue, ARGB <c>#FFB0E0E6</c>.</summary>
    public static readonly Color PowderBlue = new(r: 176, g: 224, b: 230);
    /// <summary>The color purple, ARGB <c>#FF800080</c>.</summary>
    public static readonly Color Purple = new(r: 128, g: 0, b: 128);
    /// <summary>The color red, ARGB <c>#FFFF0000</c>.</summary>
    public static readonly Color Red = new(r: 255, g: 0, b: 0);
    /// <summary>The color rosy brown, ARGB <c>#FFBC8F8F</c>.</summary>
    public static readonly Color RosyBrown = new(r: 188, g: 143, b: 143);
    /// <summary>The color royal blue, ARGB <c>#FF4169E1</c>.</summary>
    public static readonly Color RoyalBlue = new(r: 65, g: 105, b: 225);
    /// <summary>The color saddle brown, ARGB <c>#FF8B4513</c>.</summary>
    public static readonly Color SaddleBrown = new(r: 139, g: 69, b: 19);
    /// <summary>The color salmon, ARGB <c>#FFFA8072</c>.</summary>
    public static readonly Color Salmon = new(r: 250, g: 128, b: 114);
    /// <summary>The color sandy brown, ARGB <c>#FFF4A460</c>.</summary>
    public static readonly Color SandyBrown = new(r: 244, g: 164, b: 96);
    /// <summary>The color sea green, ARGB <c>#FF2E8B57</c>.</summary>
    public static readonly Color SeaGreen = new(r: 46, g: 139, b: 87);
    /// <summary>The color sea shell, ARGB <c>#FFFFF5EE</c>.</summary>
    public static readonly Color SeaShell = new(r: 255, g: 245, b: 238);
    /// <summary>The color sienna, ARGB <c>#FFA0522D</c>.</summary>
    public static readonly Color Sienna = new(r: 160, g: 82, b: 45);
    /// <summary>The color silver, ARGB <c>#FFC0C0C0</c>.</summary>
    public static readonly Color Silver = new(r: 192, g: 192, b: 192);
    /// <summary>The color sky blue, ARGB <c>#FF87CEEB</c>.</summary>
    public static readonly Color SkyBlue = new(r: 135, g: 206, b: 235);
    /// <summary>The color slate blue, ARGB <c>#FF6A5ACD</c>.</summary>
    public static readonly Color SlateBlue = new(r: 106, g: 90, b: 205);
    /// <summary>The color slate gray, ARGB <c>#FF708090</c>.</summary>
    public static readonly Color SlateGray = new(r: 112, g: 128, b: 144);
    /// <summary>The color snow, ARGB <c>#FFFFFAFA</c>.</summary>
    public static readonly Color Snow = new(r: 255, g: 250, b: 250);
    /// <summary>The color spring green, ARGB <c>#FF00FF7F</c>.</summary>
    public static readonly Color SpringGreen = new(r: 0, g: 255, b: 127);
    /// <summary>The color steel blue, ARGB <c>#FF4682B4</c>.</summary>
    public static readonly Color SteelBlue = new(r: 70, g: 130, b: 180);
    /// <summary>The color tan, ARGB <c>#FFD2B48C</c>.</summary>
    public static readonly Color Tan = new(r: 210, g: 180, b: 140);
    /// <summary>The color teal, ARGB <c>#FF008080</c>.</summary>
    public static readonly Color Teal = new(r: 0, g: 128, b: 128);
    /// <summary>The color thistle, ARGB <c>#FFD8BFD8</c>.</summary>
    public static readonly Color Thistle = new(r: 216, g: 191, b: 216);
    /// <summary>The color tomato, ARGB <c>#FFFF6347</c>.</summary>
    public static readonly Color Tomato = new(r: 255, g: 99, b: 71);
    /// <summary>The color turquoise, ARGB <c>#FF40E0D0</c>.</summary>
    public static readonly Color Turquoise = new(r: 64, g: 224, b: 208);
    /// <summary>The color violet, ARGB <c>#FFEE82EE</c>.</summary>
    public static readonly Color Violet = new(r: 238, g: 130, b: 238);
    /// <summary>The color wheat, ARGB <c>#FFF5DEB3</c>.</summary>
    public static readonly Color Wheat = new(r: 245, g: 222, b: 179);
    /// <summary>The color white, ARGB <c>#FFFFFFFF</c>.</summary>
    public static readonly Color White = new(r: 255, g: 255, b: 255);
    /// <summary>The color white smoke, ARGB <c>#FFF5F5F5</c>.</summary>
    public static readonly Color WhiteSmoke = new(r: 245, g: 245, b: 245);
    /// <summary>The color yellow, ARGB <c>#FFFFFF00</c>.</summary>
    public static readonly Color Yellow = new(r: 255, g: 255, b: 0);
    /// <summary>The color yellow green, ARGB <c>#FF9ACD32</c>.</summary>
    public static readonly Color YellowGreen = new(r: 154, g: 205, b: 50);
    /// <summary>A color with every component zero, ARGB <c>#00000000</c>.</summary>
    public static readonly Color Empty = new(a: 0, r: 0, g: 0, b: 0);
}
