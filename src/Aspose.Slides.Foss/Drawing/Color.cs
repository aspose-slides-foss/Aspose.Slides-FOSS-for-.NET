namespace Aspose.Slides.Foss.Drawing;

/// <summary>
/// Immutable value type representing an ARGB color.
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

    /// <inheritdoc/>
    public override string ToString() =>
        _a == 255
            ? $"Color(r={_r}, g={_g}, b={_b})"
            : $"Color(a={_a}, r={_r}, g={_g}, b={_b})";

    // Named color constants
    public static readonly Color Transparent = new(a: 0, r: 255, g: 255, b: 255);
    public static readonly Color AliceBlue = new(r: 240, g: 248, b: 255);
    public static readonly Color AntiqueWhite = new(r: 250, g: 235, b: 215);
    public static readonly Color Aqua = new(r: 0, g: 255, b: 255);
    public static readonly Color Aquamarine = new(r: 127, g: 255, b: 212);
    public static readonly Color Azure = new(r: 240, g: 255, b: 255);
    public static readonly Color Beige = new(r: 245, g: 245, b: 220);
    public static readonly Color Bisque = new(r: 255, g: 228, b: 196);
    public static readonly Color Black = new(r: 0, g: 0, b: 0);
    public static readonly Color BlanchedAlmond = new(r: 255, g: 235, b: 205);
    public static readonly Color Blue = new(r: 0, g: 0, b: 255);
    public static readonly Color BlueViolet = new(r: 138, g: 43, b: 226);
    public static readonly Color Brown = new(r: 165, g: 42, b: 42);
    public static readonly Color BurlyWood = new(r: 222, g: 184, b: 135);
    public static readonly Color CadetBlue = new(r: 95, g: 158, b: 160);
    public static readonly Color Chartreuse = new(r: 127, g: 255, b: 0);
    public static readonly Color Chocolate = new(r: 210, g: 105, b: 30);
    public static readonly Color Coral = new(r: 255, g: 127, b: 80);
    public static readonly Color CornflowerBlue = new(r: 100, g: 149, b: 237);
    public static readonly Color Cornsilk = new(r: 255, g: 248, b: 220);
    public static readonly Color Crimson = new(r: 220, g: 20, b: 60);
    public static readonly Color Cyan = new(r: 0, g: 255, b: 255);
    public static readonly Color DarkBlue = new(r: 0, g: 0, b: 139);
    public static readonly Color DarkCyan = new(r: 0, g: 139, b: 139);
    public static readonly Color DarkGoldenrod = new(r: 184, g: 134, b: 11);
    public static readonly Color DarkGray = new(r: 169, g: 169, b: 169);
    public static readonly Color DarkGreen = new(r: 0, g: 100, b: 0);
    public static readonly Color DarkKhaki = new(r: 189, g: 183, b: 107);
    public static readonly Color DarkMagenta = new(r: 139, g: 0, b: 139);
    public static readonly Color DarkOliveGreen = new(r: 85, g: 107, b: 47);
    public static readonly Color DarkOrange = new(r: 255, g: 140, b: 0);
    public static readonly Color DarkOrchid = new(r: 153, g: 50, b: 204);
    public static readonly Color DarkRed = new(r: 139, g: 0, b: 0);
    public static readonly Color DarkSalmon = new(r: 233, g: 150, b: 122);
    public static readonly Color DarkSeaGreen = new(r: 143, g: 188, b: 143);
    public static readonly Color DarkSlateBlue = new(r: 72, g: 61, b: 139);
    public static readonly Color DarkSlateGray = new(r: 47, g: 79, b: 79);
    public static readonly Color DarkTurquoise = new(r: 0, g: 206, b: 209);
    public static readonly Color DarkViolet = new(r: 148, g: 0, b: 211);
    public static readonly Color DeepPink = new(r: 255, g: 20, b: 147);
    public static readonly Color DeepSkyBlue = new(r: 0, g: 191, b: 255);
    public static readonly Color DimGray = new(r: 105, g: 105, b: 105);
    public static readonly Color DodgerBlue = new(r: 30, g: 144, b: 255);
    public static readonly Color Firebrick = new(r: 178, g: 34, b: 34);
    public static readonly Color FloralWhite = new(r: 255, g: 250, b: 240);
    public static readonly Color ForestGreen = new(r: 34, g: 139, b: 34);
    public static readonly Color Fuchsia = new(r: 255, g: 0, b: 255);
    public static readonly Color Gainsboro = new(r: 220, g: 220, b: 220);
    public static readonly Color GhostWhite = new(r: 248, g: 248, b: 255);
    public static readonly Color Gold = new(r: 255, g: 215, b: 0);
    public static readonly Color Goldenrod = new(r: 218, g: 165, b: 32);
    public static readonly Color Gray = new(r: 128, g: 128, b: 128);
    public static readonly Color Green = new(r: 0, g: 128, b: 0);
    public static readonly Color GreenYellow = new(r: 173, g: 255, b: 47);
    public static readonly Color Honeydew = new(r: 240, g: 255, b: 240);
    public static readonly Color HotPink = new(r: 255, g: 105, b: 180);
    public static readonly Color IndianRed = new(r: 205, g: 92, b: 92);
    public static readonly Color Indigo = new(r: 75, g: 0, b: 130);
    public static readonly Color Ivory = new(r: 255, g: 255, b: 240);
    public static readonly Color Khaki = new(r: 240, g: 230, b: 140);
    public static readonly Color Lavender = new(r: 230, g: 230, b: 250);
    public static readonly Color LavenderBlush = new(r: 255, g: 240, b: 245);
    public static readonly Color LawnGreen = new(r: 124, g: 252, b: 0);
    public static readonly Color LemonChiffon = new(r: 255, g: 250, b: 205);
    public static readonly Color LightBlue = new(r: 173, g: 216, b: 230);
    public static readonly Color LightCoral = new(r: 240, g: 128, b: 128);
    public static readonly Color LightCyan = new(r: 224, g: 255, b: 255);
    public static readonly Color LightGoldenrodYellow = new(r: 250, g: 250, b: 210);
    public static readonly Color LightGray = new(r: 211, g: 211, b: 211);
    public static readonly Color LightGreen = new(r: 144, g: 238, b: 144);
    public static readonly Color LightPink = new(r: 255, g: 182, b: 193);
    public static readonly Color LightSalmon = new(r: 255, g: 160, b: 122);
    public static readonly Color LightSeaGreen = new(r: 32, g: 178, b: 170);
    public static readonly Color LightSkyBlue = new(r: 135, g: 206, b: 250);
    public static readonly Color LightSlateGray = new(r: 119, g: 136, b: 153);
    public static readonly Color LightSteelBlue = new(r: 176, g: 196, b: 222);
    public static readonly Color LightYellow = new(r: 255, g: 255, b: 224);
    public static readonly Color Lime = new(r: 0, g: 255, b: 0);
    public static readonly Color LimeGreen = new(r: 50, g: 205, b: 50);
    public static readonly Color Linen = new(r: 250, g: 240, b: 230);
    public static readonly Color Magenta = new(r: 255, g: 0, b: 255);
    public static readonly Color Maroon = new(r: 128, g: 0, b: 0);
    public static readonly Color MediumAquamarine = new(r: 102, g: 205, b: 170);
    public static readonly Color MediumBlue = new(r: 0, g: 0, b: 205);
    public static readonly Color MediumOrchid = new(r: 186, g: 85, b: 211);
    public static readonly Color MediumPurple = new(r: 147, g: 112, b: 219);
    public static readonly Color MediumSeaGreen = new(r: 60, g: 179, b: 113);
    public static readonly Color MediumSlateBlue = new(r: 123, g: 104, b: 238);
    public static readonly Color MediumSpringGreen = new(r: 0, g: 250, b: 154);
    public static readonly Color MediumTurquoise = new(r: 72, g: 209, b: 204);
    public static readonly Color MediumVioletRed = new(r: 199, g: 21, b: 133);
    public static readonly Color MidnightBlue = new(r: 25, g: 25, b: 112);
    public static readonly Color MintCream = new(r: 245, g: 255, b: 250);
    public static readonly Color MistyRose = new(r: 255, g: 228, b: 225);
    public static readonly Color Moccasin = new(r: 255, g: 228, b: 181);
    public static readonly Color NavajoWhite = new(r: 255, g: 222, b: 173);
    public static readonly Color Navy = new(r: 0, g: 0, b: 128);
    public static readonly Color OldLace = new(r: 253, g: 245, b: 230);
    public static readonly Color Olive = new(r: 128, g: 128, b: 0);
    public static readonly Color OliveDrab = new(r: 107, g: 142, b: 35);
    public static readonly Color Orange = new(r: 255, g: 165, b: 0);
    public static readonly Color OrangeRed = new(r: 255, g: 69, b: 0);
    public static readonly Color Orchid = new(r: 218, g: 112, b: 214);
    public static readonly Color PaleGoldenrod = new(r: 238, g: 232, b: 170);
    public static readonly Color PaleGreen = new(r: 152, g: 251, b: 152);
    public static readonly Color PaleTurquoise = new(r: 175, g: 238, b: 238);
    public static readonly Color PaleVioletRed = new(r: 219, g: 112, b: 147);
    public static readonly Color PapayaWhip = new(r: 255, g: 239, b: 213);
    public static readonly Color PeachPuff = new(r: 255, g: 218, b: 185);
    public static readonly Color Peru = new(r: 205, g: 133, b: 63);
    public static readonly Color Pink = new(r: 255, g: 192, b: 203);
    public static readonly Color Plum = new(r: 221, g: 160, b: 221);
    public static readonly Color PowderBlue = new(r: 176, g: 224, b: 230);
    public static readonly Color Purple = new(r: 128, g: 0, b: 128);
    public static readonly Color Red = new(r: 255, g: 0, b: 0);
    public static readonly Color RosyBrown = new(r: 188, g: 143, b: 143);
    public static readonly Color RoyalBlue = new(r: 65, g: 105, b: 225);
    public static readonly Color SaddleBrown = new(r: 139, g: 69, b: 19);
    public static readonly Color Salmon = new(r: 250, g: 128, b: 114);
    public static readonly Color SandyBrown = new(r: 244, g: 164, b: 96);
    public static readonly Color SeaGreen = new(r: 46, g: 139, b: 87);
    public static readonly Color SeaShell = new(r: 255, g: 245, b: 238);
    public static readonly Color Sienna = new(r: 160, g: 82, b: 45);
    public static readonly Color Silver = new(r: 192, g: 192, b: 192);
    public static readonly Color SkyBlue = new(r: 135, g: 206, b: 235);
    public static readonly Color SlateBlue = new(r: 106, g: 90, b: 205);
    public static readonly Color SlateGray = new(r: 112, g: 128, b: 144);
    public static readonly Color Snow = new(r: 255, g: 250, b: 250);
    public static readonly Color SpringGreen = new(r: 0, g: 255, b: 127);
    public static readonly Color SteelBlue = new(r: 70, g: 130, b: 180);
    public static readonly Color Tan = new(r: 210, g: 180, b: 140);
    public static readonly Color Teal = new(r: 0, g: 128, b: 128);
    public static readonly Color Thistle = new(r: 216, g: 191, b: 216);
    public static readonly Color Tomato = new(r: 255, g: 99, b: 71);
    public static readonly Color Turquoise = new(r: 64, g: 224, b: 208);
    public static readonly Color Violet = new(r: 238, g: 130, b: 238);
    public static readonly Color Wheat = new(r: 245, g: 222, b: 179);
    public static readonly Color White = new(r: 255, g: 255, b: 255);
    public static readonly Color WhiteSmoke = new(r: 245, g: 245, b: 245);
    public static readonly Color Yellow = new(r: 255, g: 255, b: 0);
    public static readonly Color YellowGreen = new(r: 154, g: 205, b: 50);
    public static readonly Color Empty = new(a: 0, r: 0, g: 0, b: 0);
}
