using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;

namespace TinyExcel;

public struct XColor : IEquatable<XColor>
{
    //private static readonly Dictionary<string, XColor> _paletteColors = new Dictionary<string, XColor>(StringComparer.OrdinalIgnoreCase)
    //{
    //    { "ButtonFace", FromRgb(0xF0F0F0) },
    //    { "WindowText", FromRgb(0x000000) },
    //    { "Menu", FromRgb(0xF0F0F0) },
    //    { "Highlight", FromRgb(0x0078D7) },
    //    { "HighlightText", FromRgb(0xFFFFFF) },
    //    { "CaptionText", FromRgb(0x000000) },
    //    { "ActiveCaption", FromRgb(0x99B4D1) },
    //    { "ButtonHighlight", FromRgb(0xFFFFFF) },
    //    { "ButtonShadow", FromRgb(0xA0A0A0) },
    //    { "ButtonText", FromRgb(0x000000) },
    //    { "GrayText", FromRgb(0x6D6D6D) },
    //    { "InactiveCaption", FromRgb(0xBFCDDB) },
    //    { "InactiveCaptionText", FromRgb(0x000000) },
    //    { "InfoBackground", FromRgb(0xFFFFE1) },
    //    { "InfoText", FromRgb(0x000000) },
    //    { "MenuText", FromRgb(0x000000) },
    //    { "Scrollbar", FromRgb(0xC8C8C8) },
    //    { "Window", FromRgb(0xFFFFFF) },
    //    { "WindowFrame", FromRgb(0x646464) },
    //    { "ThreeDLightShadow", FromRgb(0x000000) },
    //    { "ThreeDDarkShadow", FromRgb(0x696969) },
    //    { "ActiveBorder", FromRgb(0xB4B4B4) },
    //    { "InactiveBorder", FromRgb(0xF4F7FC) },
    //    { "Background", FromRgb(0x000000) },
    //    { "AppWorkspace", FromRgb(0xABABAB) },
    //    { "ThreeDFace", FromRgb(0xF0F0F0) },
    //    { "ThreeDShadow", FromRgb(0xA0A0A0) },
    //    { "ThreeDHighlight", FromRgb(0xFFFFFF) }
    //};
    private static readonly Dictionary<int, XColor> _indexedColors = new Dictionary<int, XColor>
    {
        { 0, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF000000"), Value = 0 }},
        { 1, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FFFFFFFF"), Value = 1 }},
        { 2, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FFFF0000"), Value = 2 }},
        { 3, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF00FF00"), Value = 3 }},
        { 4, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF0000FF"), Value = 4 }},
        { 5, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FFFFFF00"), Value = 5 }},
        { 6, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FFFF00FF"), Value = 6 }},
        { 7, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF00FFFF"), Value = 7 }},
        { 8, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF000000"), Value = 8 }},
        { 9, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FFFFFFFF"), Value = 9 }},
        { 10, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FFFF0000"), Value = 10 }},
        { 11, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF00FF00"), Value = 11 }},
        { 12, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF0000FF"), Value = 12 }},
        { 13, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FFFFFF00"), Value = 13 }},
        { 14, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FFFF00FF"), Value = 14 }},
        { 15, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF00FFFF"), Value = 15 }},
        { 16, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF800000"), Value = 16 }},
        { 17, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF008000"), Value = 17 }},
        { 18, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF000080"), Value = 18 }},
        { 19, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF808000"), Value = 19 }},
        { 20, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF800080"), Value = 20 }},
        { 21, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF008080"), Value = 21 }},
        { 22, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FFC0C0C0"), Value = 22 }},
        { 23, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF808080"), Value = 23 }},
        { 24, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF9999FF"), Value = 24 }},
        { 25, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF993366"), Value = 25 }},
        { 26, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FFFFFFCC"), Value = 26 }},
        { 27, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FFCCFFFF"), Value = 27 }},
        { 28, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF660066"), Value = 28 }},
        { 29, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FFFF8080"), Value = 29 }},
        { 30, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF0066CC"), Value = 30 }},
        { 31, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FFCCCCFF"), Value = 31 }},
        { 32, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF000080"), Value = 32 }},
        { 33, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FFFF00FF"), Value = 33 }},
        { 34, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FFFFFF00"), Value = 34 }},
        { 35, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF00FFFF"), Value = 35 }},
        { 36, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF800080"), Value = 36 }},
        { 37, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF800000"), Value = 37 }},
        { 38, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF008080"), Value = 38 }},
        { 39, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF0000FF"), Value = 39 }},
        { 40, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF00CCFF"), Value = 40 }},
        { 41, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FFCCFFFF"), Value = 41 }},
        { 42, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FFCCFFCC"), Value = 42 }},
        { 43, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FFFFFF99"), Value = 43 }},
        { 44, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF99CCFF"), Value = 44 }},
        { 45, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FFFF99CC"), Value = 45 }},
        { 46, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FFCC99FF"), Value = 46 }},
        { 47, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FFFFCC99"), Value = 47 }},
        { 48, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF3366FF"), Value = 48 }},
        { 49, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF33CCCC"), Value = 49 }},
        { 50, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF99CC00"), Value = 50 }},
        { 51, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FFFFCC00"), Value = 51 }},
        { 52, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FFFF9900"), Value = 52 }},
        { 53, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FFFF6600"), Value = 53 }},
        { 54, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF666699"), Value = 54 }},
        { 55, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF969696"), Value = 55 }},
        { 56, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF003366"), Value = 56 }},
        { 57, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF339966"), Value = 57 }},
        { 58, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF003300"), Value = 58 }},
        { 59, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF333300"), Value = 59 }},
        { 60, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF993300"), Value = 60 }},
        { 61, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF993366"), Value = 61 }},
        { 62, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF333399"), Value = 62 }},
        { 63, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = ParseFromHtml("#FF333333"), Value = 63 }},
        { 64, new XColor{ IsEmpty = false, ColorType = XColorType.Indexed, Color = Color.Transparent, Value = 64 }}
    };
    private static readonly Dictionary<Color, int> _colorIndices = new Dictionary<Color, int>
    {
        {ParseFromHtml("#FF000000"),0},
        {ParseFromHtml("#FFFFFFFF"),1},
        {ParseFromHtml("#FFFF0000"),2},
        {ParseFromHtml("#FF00FF00"),3},
        {ParseFromHtml("#FF0000FF"),4},
        {ParseFromHtml("#FFFFFF00"),5},
        {ParseFromHtml("#FFFF00FF"),6},
        {ParseFromHtml("#FF00FFFF"),7},
        {ParseFromHtml("#FF000000"),8},
        {ParseFromHtml("#FFFFFFFF"),9},
        {ParseFromHtml("#FFFF0000"),10},
        {ParseFromHtml("#FF00FF00"),11},
        {ParseFromHtml("#FF0000FF"),12},
        {ParseFromHtml("#FFFFFF00"),13},
        {ParseFromHtml("#FFFF00FF"),14},
        {ParseFromHtml("#FF00FFFF"),15},
        {ParseFromHtml("#FF800000"),16},
        {ParseFromHtml("#FF008000"),17},
        {ParseFromHtml("#FF000080"),18},
        {ParseFromHtml("#FF808000"),19},
        {ParseFromHtml("#FF800080"),20},
        {ParseFromHtml("#FF008080"),21},
        {ParseFromHtml("#FFC0C0C0"),22},
        {ParseFromHtml("#FF808080"),23},
        {ParseFromHtml("#FF9999FF"),24},
        {ParseFromHtml("#FF993366"),25},
        {ParseFromHtml("#FFFFFFCC"),26},
        {ParseFromHtml("#FFCCFFFF"),27},
        {ParseFromHtml("#FF660066"),28},
        {ParseFromHtml("#FFFF8080"),29},
        {ParseFromHtml("#FF0066CC"),30},
        {ParseFromHtml("#FFCCCCFF"),31},
        {ParseFromHtml("#FF000080"),32},
        {ParseFromHtml("#FFFF00FF"),33},
        {ParseFromHtml("#FFFFFF00"),34},
        {ParseFromHtml("#FF00FFFF"),35},
        {ParseFromHtml("#FF800080"),36},
        {ParseFromHtml("#FF800000"),37},
        {ParseFromHtml("#FF008080"),38},
        {ParseFromHtml("#FF0000FF"),39},
        {ParseFromHtml("#FF00CCFF"),40},
        {ParseFromHtml("#FFCCFFFF"),41},
        {ParseFromHtml("#FFCCFFCC"),42},
        {ParseFromHtml("#FFFFFF99"),43},
        {ParseFromHtml("#FF99CCFF"),44},
        {ParseFromHtml("#FFFF99CC"),45},
        {ParseFromHtml("#FFCC99FF"),46},
        {ParseFromHtml("#FFFFCC99"),47},
        {ParseFromHtml("#FF3366FF"),48},
        {ParseFromHtml("#FF33CCCC"),49},
        {ParseFromHtml("#FF99CC00"),50},
        {ParseFromHtml("#FFFFCC00"),51},
        {ParseFromHtml("#FFFF9900"),52},
        {ParseFromHtml("#FFFF6600"),53},
        {ParseFromHtml("#FF666699"),54},
        {ParseFromHtml("#FF969696"),55},
        {ParseFromHtml("#FF003366"),56},
        {ParseFromHtml("#FF339966"),57},
        {ParseFromHtml("#FF003300"),58},
        {ParseFromHtml("#FF333300"),59},
        {ParseFromHtml("#FF993300"),60},
        {ParseFromHtml("#FF993366"),61},
        {ParseFromHtml("#FF333399"),62},
        {ParseFromHtml("#FF333333"),63},
        {Color.Transparent,64}
    };

    public bool IsEmpty { get; set; } = true;
    public XColorType ColorType { get; set; }
    /// <summary>
    /// 索引或是ThemeColor枚举值，或是null
    /// </summary>
    public object Value { get; set; }
    public Color Color { get; set; }
    public double? Tint { get; set; }

    public static readonly XColor Empty = new XColor();

    public XColor() { }

    public async Task Write(StreamWriter writer)
       => await this.Write(writer, "color");
    public async Task<bool> Write(StreamWriter writer, string localName)
    {
        if (this.IsEmpty) return false;

        await writer.WriteAsync($"<{localName}");
        switch (this.ColorType)
        {
            //<x:color indexed="1" />
            case XColorType.Indexed:
                await writer.WriteAsync($" indexed={this.Value}");
                break;
            case XColorType.Color:
                //<x:color rgb="FF000000" />
                await writer.WriteAsync($" rgb={this.Color.ToArgbString()}");
                break;
            //<x:color theme="1" tint="0.3" />
            case XColorType.Theme:
                await writer.WriteAsync($" theme={this.Value}");
                if (this.Tint.HasValue)
                    await writer.WriteAsync($" tint={this.Tint}");
                break;
        };
        await writer.WriteAsync("/>");
        return true;
    }

    public bool Equals(XColor other)
    {
        if (this.IsEmpty != other.IsEmpty) return false;
        if (this.IsEmpty && other.IsEmpty) return true;
        if (!this.IsEmpty && !other.IsEmpty && this.ColorType == other.ColorType)
        {
            switch (this.ColorType)
            {
                case XColorType.Color:
                    return this.Color.ToArgb() == other.Color.ToArgb();
                case XColorType.Theme:
                    if (this.Value.Equals(other.Value) && this.Tint == other.Tint)
                        return true;
                    break;
                case XColorType.Indexed:
                    return this.Value.Equals(other.Value);
            }
        }
        return false;
    }
    public override bool Equals(object other)
        => other is XColor && Equals((XColor)other);
    public override int GetHashCode() => HashCode.Combine(this.IsEmpty, this.ColorType, this.Value, this.Tint);
    public static bool operator ==(XColor left, XColor right) => left.Equals(right);
    public static bool operator !=(XColor left, XColor right) => !(left == right);

    public static XColor From(Color color)
    {
        if (_colorIndices.TryGetValue(color, out var index))
            return _indexedColors[index];
        return new XColor { IsEmpty = false, ColorType = XColorType.Color, Value = color, Color = color };
    }
    public static XColor FromTheme(XThemeColor themeColor)
        => new XColor { IsEmpty = false, ColorType = XColorType.Theme, Value = themeColor };
    public static XColor FromTheme(XThemeColor themeColor, double tint)
       => new XColor { IsEmpty = false, ColorType = XColorType.Theme, Value = (int)themeColor, Tint = tint };
    public static XColor FromIndex(int colorIndex)
    {
        if (!_indexedColors.TryGetValue(colorIndex, out var result))
            throw new IndexOutOfRangeException($"不存在索引{colorIndex}的颜色，索引范围：0 - 64");
        return result;
    }
    public static XColor FromArgb(int argb) => From(Color.FromArgb(argb));
    public static XColor FromArgb(uint argb) => From(Color.FromArgb(unchecked((int)argb)));
    public static XColor FromArgb(int r, int g, int b) => From(Color.FromArgb(r, g, b));
    public static XColor FromArgb(int a, int r, int g, int b) => From(Color.FromArgb(a, r, g, b));
    public static XColor FromRgb(int rgb)
    {
        unchecked
        {
            return From(Color.FromArgb(rgb | (int)0xFF000000));
        }
    }
    public static XColor FromName(string name) => From(Color.FromName(name));
    public static XColor FromHtml(string htmlColor)
    {
        var color = ParseFromHtml(htmlColor);
        return From(color);
    }
    public static XColor FromHexRgb(String hexColorRgb) => From(ParseFromRgb(hexColorRgb));
    public static XColor NoColor { get; } = new XColor();
    public static XColor AliceBlue => From(Color.AliceBlue);
    public static XColor AntiqueWhite => From(Color.AntiqueWhite);
    public static XColor Aqua => From(Color.Aqua);
    public static XColor Aquamarine => From(Color.Aquamarine);
    public static XColor Azure => From(Color.Azure);
    public static XColor Beige => From(Color.Beige);
    public static XColor Bisque => From(Color.Bisque);
    public static XColor Black => From(Color.Black);
    public static XColor BlanchedAlmond => From(Color.BlanchedAlmond);
    public static XColor Blue => From(Color.Blue);
    public static XColor BlueViolet => From(Color.BlueViolet);
    public static XColor Brown => From(Color.Brown);
    public static XColor BurlyWood => From(Color.BurlyWood);
    public static XColor CadetBlue => From(Color.CadetBlue);
    public static XColor Chartreuse => From(Color.Chartreuse);
    public static XColor Chocolate => From(Color.Chocolate);
    public static XColor Coral => From(Color.Coral);
    public static XColor CornflowerBlue => From(Color.CornflowerBlue);
    public static XColor Cornsilk => From(Color.Cornsilk);
    public static XColor Crimson => From(Color.Crimson);
    public static XColor Cyan => From(Color.Cyan);
    public static XColor DarkBlue => From(Color.DarkBlue);
    public static XColor DarkCyan => From(Color.DarkCyan);
    public static XColor DarkGoldenrod => From(Color.DarkGoldenrod);
    public static XColor DarkGray => From(Color.DarkGray);
    public static XColor DarkGreen => From(Color.DarkGreen);
    public static XColor DarkKhaki => From(Color.DarkKhaki);
    public static XColor DarkMagenta => From(Color.DarkMagenta);
    public static XColor DarkOliveGreen => From(Color.DarkOliveGreen);
    public static XColor DarkOrange => From(Color.DarkOrange);
    public static XColor DarkOrchid => From(Color.DarkOrchid);
    public static XColor DarkRed => From(Color.DarkRed);
    public static XColor DarkSalmon => From(Color.DarkSalmon);
    public static XColor DarkSeaGreen => From(Color.DarkSeaGreen);
    public static XColor DarkSlateBlue => From(Color.DarkSlateBlue);
    public static XColor DarkSlateGray => From(Color.DarkSlateGray);
    public static XColor DarkTurquoise => From(Color.DarkTurquoise);
    public static XColor DarkViolet => From(Color.DarkViolet);
    public static XColor DeepPink => From(Color.DeepPink);
    public static XColor DeepSkyBlue => From(Color.DeepSkyBlue);
    public static XColor DimGray => From(Color.DimGray);
    public static XColor DodgerBlue => From(Color.DodgerBlue);
    public static XColor Firebrick => From(Color.Firebrick);
    public static XColor FloralWhite => From(Color.FloralWhite);
    public static XColor ForestGreen => From(Color.ForestGreen);
    public static XColor Fuchsia => From(Color.Fuchsia);
    public static XColor Gainsboro => From(Color.Gainsboro);
    public static XColor GhostWhite => From(Color.GhostWhite);
    public static XColor Gold => From(Color.Gold);
    public static XColor Goldenrod => From(Color.Goldenrod);
    public static XColor Gray => From(Color.Gray);
    public static XColor Green => From(Color.Green);
    public static XColor GreenYellow => From(Color.GreenYellow);
    public static XColor Honeydew => From(Color.Honeydew);
    public static XColor HotPink => From(Color.HotPink);
    public static XColor IndianRed => From(Color.IndianRed);
    public static XColor Indigo => From(Color.Indigo);
    public static XColor Ivory => From(Color.Ivory);
    public static XColor Khaki => From(Color.Khaki);
    public static XColor Lavender => From(Color.Lavender);
    public static XColor LavenderBlush => From(Color.LavenderBlush);
    public static XColor LawnGreen => From(Color.LawnGreen);
    public static XColor LemonChiffon => From(Color.LemonChiffon);
    public static XColor LightBlue => From(Color.LightBlue);
    public static XColor LightCoral => From(Color.LightCoral);
    public static XColor LightCyan => From(Color.LightCyan);
    public static XColor LightGoldenrodYellow => From(Color.LightGoldenrodYellow);
    public static XColor LightGray => From(Color.LightGray);
    public static XColor LightGreen => From(Color.LightGreen);
    public static XColor LightPink => From(Color.LightPink);
    public static XColor LightSalmon => From(Color.LightSalmon);
    public static XColor LightSeaGreen => From(Color.LightSeaGreen);
    public static XColor LightSkyBlue => From(Color.LightSkyBlue);
    public static XColor LightSlateGray => From(Color.LightSlateGray);
    public static XColor LightSteelBlue => From(Color.LightSteelBlue);
    public static XColor LightYellow => From(Color.LightYellow);
    public static XColor Lime => From(Color.Lime);
    public static XColor LimeGreen => From(Color.LimeGreen);
    public static XColor Linen => From(Color.Linen);
    public static XColor Magenta => From(Color.Magenta);
    public static XColor Maroon => From(Color.Maroon);
    public static XColor MediumAquamarine => From(Color.MediumAquamarine);
    public static XColor MediumBlue => From(Color.MediumBlue);
    public static XColor MediumOrchid => From(Color.MediumOrchid);
    public static XColor MediumPurple => From(Color.MediumPurple);
    public static XColor MediumSeaGreen => From(Color.MediumSeaGreen);
    public static XColor MediumSlateBlue => From(Color.MediumSlateBlue);
    public static XColor MediumSpringGreen => From(Color.MediumSpringGreen);
    public static XColor MediumTurquoise => From(Color.MediumTurquoise);
    public static XColor MediumVioletRed => From(Color.MediumVioletRed);
    public static XColor MidnightBlue => From(Color.MidnightBlue);
    public static XColor MintCream => From(Color.MintCream);
    public static XColor MistyRose => From(Color.MistyRose);
    public static XColor Moccasin => From(Color.Moccasin);
    public static XColor NavajoWhite => From(Color.NavajoWhite);
    public static XColor Navy => From(Color.Navy);
    public static XColor OldLace => From(Color.OldLace);
    public static XColor Olive => From(Color.Olive);
    public static XColor OliveDrab => From(Color.OliveDrab);
    public static XColor Orange => From(Color.Orange);
    public static XColor OrangeRed => From(Color.OrangeRed);
    public static XColor Orchid => From(Color.Orchid);
    public static XColor PaleGoldenrod => From(Color.PaleGoldenrod);
    public static XColor PaleGreen => From(Color.PaleGreen);
    public static XColor PaleTurquoise => From(Color.PaleTurquoise);
    public static XColor PaleVioletRed => From(Color.PaleVioletRed);
    public static XColor PapayaWhip => From(Color.PapayaWhip);
    public static XColor PeachPuff => From(Color.PeachPuff);
    public static XColor Peru => From(Color.Peru);
    public static XColor Pink => From(Color.Pink);
    public static XColor Plum => From(Color.Plum);
    public static XColor PowderBlue => From(Color.PowderBlue);
    public static XColor Purple => From(Color.Purple);
    public static XColor Red => From(Color.Red);
    public static XColor RosyBrown => From(Color.RosyBrown);
    public static XColor RoyalBlue => From(Color.RoyalBlue);
    public static XColor SaddleBrown => From(Color.SaddleBrown);
    public static XColor Salmon => From(Color.Salmon);
    public static XColor SandyBrown => From(Color.SandyBrown);
    public static XColor SeaGreen => From(Color.SeaGreen);
    public static XColor SeaShell => From(Color.SeaShell);
    public static XColor Sienna => From(Color.Sienna);
    public static XColor Silver => From(Color.Silver);
    public static XColor SkyBlue => From(Color.SkyBlue);
    public static XColor SlateBlue => From(Color.SlateBlue);
    public static XColor SlateGray => From(Color.SlateGray);
    public static XColor Snow => From(Color.Snow);
    public static XColor SpringGreen => From(Color.SpringGreen);
    public static XColor SteelBlue => From(Color.SteelBlue);
    public static XColor Tan => From(Color.Tan);
    public static XColor Teal => From(Color.Teal);
    public static XColor Thistle => From(Color.Thistle);
    public static XColor Tomato => From(Color.Tomato);
    public static XColor Turquoise => From(Color.Turquoise);
    public static XColor Violet => From(Color.Violet);
    public static XColor Wheat => From(Color.Wheat);
    public static XColor White => From(Color.White);
    public static XColor WhiteSmoke => From(Color.WhiteSmoke);
    public static XColor Yellow => From(Color.Yellow);
    public static XColor YellowGreen => From(Color.YellowGreen);
    public static XColor AirForceBlue => FromHtml("#FF5D8AA8");
    public static XColor Alizarin => FromHtml("#FFE32636");
    public static XColor Almond => FromHtml("#FFEFDECD");
    public static XColor Amaranth => FromHtml("#FFE52B50");
    public static XColor Amber => FromHtml("#FFFFBF00");
    public static XColor AmberSaeEce => FromHtml("#FFFF7E00");
    public static XColor AmericanRose => FromHtml("#FFFF033E");
    public static XColor Amethyst => FromHtml("#FF9966CC");
    public static XColor AntiFlashWhite => FromHtml("#FFF2F3F4");
    public static XColor AntiqueBrass => FromHtml("#FFCD9575");
    public static XColor AntiqueFuchsia => FromHtml("#FF915C83");
    public static XColor AppleGreen => FromHtml("#FF8DB600");
    public static XColor Apricot => FromHtml("#FFFBCEB1");
    public static XColor Aquamarine1 => FromHtml("#FF7FFFD0");
    public static XColor ArmyGreen => FromHtml("#FF4B5320");
    public static XColor Arsenic => FromHtml("#FF3B444B");
    public static XColor ArylideYellow => FromHtml("#FFE9D66B");
    public static XColor AshGrey => FromHtml("#FFB2BEB5");
    public static XColor Asparagus => FromHtml("#FF87A96B");
    public static XColor AtomicTangerine => FromHtml("#FFFF9966");
    public static XColor Auburn => FromHtml("#FF6D351A");
    public static XColor Aureolin => FromHtml("#FFFDEE00");
    public static XColor Aurometalsaurus => FromHtml("#FF6E7F80");
    public static XColor Awesome => FromHtml("#FFFF2052");
    public static XColor AzureColorWheel => FromHtml("#FF007FFF");
    public static XColor BabyBlue => FromHtml("#FF89CFF0");
    public static XColor BabyBlueEyes => FromHtml("#FFA1CAF1");
    public static XColor BabyPink => FromHtml("#FFF4C2C2");
    public static XColor BallBlue => FromHtml("#FF21ABCD");
    public static XColor BananaMania => FromHtml("#FFFAE7B5");
    public static XColor BattleshipGrey => FromHtml("#FF848482");
    public static XColor Bazaar => FromHtml("#FF98777B");
    public static XColor BeauBlue => FromHtml("#FFBCD4E6");
    public static XColor Beaver => FromHtml("#FF9F8170");
    public static XColor Bistre => FromHtml("#FF3D2B1F");
    public static XColor Bittersweet => FromHtml("#FFFE6F5E");
    public static XColor BleuDeFrance => FromHtml("#FF318CE7");
    public static XColor BlizzardBlue => FromHtml("#FFACE5EE");
    public static XColor Blond => FromHtml("#FFFAF0BE");
    public static XColor BlueBell => FromHtml("#FFA2A2D0");
    public static XColor BlueGray => FromHtml("#FF6699CC");
    public static XColor BlueGreen => FromHtml("#FF00DDDD");
    public static XColor BluePigment => FromHtml("#FF333399");
    public static XColor BlueRyb => FromHtml("#FF0247FE");
    public static XColor Blush => FromHtml("#FFDE5D83");
    public static XColor Bole => FromHtml("#FF79443B");
    public static XColor BondiBlue => FromHtml("#FF0095B6");
    public static XColor BostonUniversityRed => FromHtml("#FFCC0000");
    public static XColor BrandeisBlue => FromHtml("#FF0070FF");
    public static XColor Brass => FromHtml("#FFB5A642");
    public static XColor BrickRed => FromHtml("#FFCB4154");
    public static XColor BrightCerulean => FromHtml("#FF1DACD6");
    public static XColor BrightGreen => FromHtml("#FF66FF00");
    public static XColor BrightLavender => FromHtml("#FFBF94E4");
    public static XColor BrightMaroon => FromHtml("#FFC32148");
    public static XColor BrightPink => FromHtml("#FFFF007F");
    public static XColor BrightTurquoise => FromHtml("#FF08E8DE");
    public static XColor BrightUbe => FromHtml("#FFD19FE8");
    public static XColor BrilliantLavender => FromHtml("#FFF4BBFF");
    public static XColor BrilliantRose => FromHtml("#FFFF55A3");
    public static XColor BrinkPink => FromHtml("#FFFB607F");
    public static XColor BritishRacingGreen => FromHtml("#FF004225");
    public static XColor Bronze => FromHtml("#FFCD7F32");
    public static XColor BrownTraditional => FromHtml("#FF964B00");
    public static XColor BubbleGum => FromHtml("#FFFFC1CC");
    public static XColor Bubbles => FromHtml("#FFE7FEFF");
    public static XColor Buff => FromHtml("#FFF0DC82");
    public static XColor BulgarianRose => FromHtml("#FF480607");
    public static XColor Burgundy => FromHtml("#FF800020");
    public static XColor BurntOrange => FromHtml("#FFCC5500");
    public static XColor BurntSienna => FromHtml("#FFE97451");
    public static XColor BurntUmber => FromHtml("#FF8A3324");
    public static XColor Byzantine => FromHtml("#FFBD33A4");
    public static XColor Byzantium => FromHtml("#FF702963");
    public static XColor Cadet => FromHtml("#FF536872");
    public static XColor CadetGrey => FromHtml("#FF91A3B0");
    public static XColor CadmiumGreen => FromHtml("#FF006B3C");
    public static XColor CadmiumOrange => FromHtml("#FFED872D");
    public static XColor CadmiumRed => FromHtml("#FFE30022");
    public static XColor CadmiumYellow => FromHtml("#FFFFF600");
    public static XColor CalPolyPomonaGreen => FromHtml("#FF1E4D2B");
    public static XColor CambridgeBlue => FromHtml("#FFA3C1AD");
    public static XColor Camel => FromHtml("#FFC19A6B");
    public static XColor CamouflageGreen => FromHtml("#FF78866B");
    public static XColor CanaryYellow => FromHtml("#FFFFEF00");
    public static XColor CandyAppleRed => FromHtml("#FFFF0800");
    public static XColor CandyPink => FromHtml("#FFE4717A");
    public static XColor CaputMortuum => FromHtml("#FF592720");
    public static XColor Cardinal => FromHtml("#FFC41E3A");
    public static XColor CaribbeanGreen => FromHtml("#FF00CC99");
    public static XColor Carmine => FromHtml("#FF960018");
    public static XColor CarminePink => FromHtml("#FFEB4C42");
    public static XColor CarmineRed => FromHtml("#FFFF0038");
    public static XColor CarnationPink => FromHtml("#FFFFA6C9");
    public static XColor Carnelian => FromHtml("#FFB31B1B");
    public static XColor CarolinaBlue => FromHtml("#FF99BADD");
    public static XColor CarrotOrange => FromHtml("#FFED9121");
    public static XColor Ceil => FromHtml("#FF92A1CF");
    public static XColor Celadon => FromHtml("#FFACE1AF");
    public static XColor CelestialBlue => FromHtml("#FF4997D0");
    public static XColor Cerise => FromHtml("#FFDE3163");
    public static XColor CerisePink => FromHtml("#FFEC3B83");
    public static XColor Cerulean => FromHtml("#FF007BA7");
    public static XColor CeruleanBlue => FromHtml("#FF2A52BE");
    public static XColor Chamoisee => FromHtml("#FFA0785A");
    public static XColor Champagne => FromHtml("#FFF7E7CE");
    public static XColor Charcoal => FromHtml("#FF36454F");
    public static XColor ChartreuseTraditional => FromHtml("#FFDFFF00");
    public static XColor CherryBlossomPink => FromHtml("#FFFFB7C5");
    public static XColor Chocolate1 => FromHtml("#FF7B3F00");
    public static XColor ChromeYellow => FromHtml("#FFFFA700");
    public static XColor Cinereous => FromHtml("#FF98817B");
    public static XColor Cinnabar => FromHtml("#FFE34234");
    public static XColor Citrine => FromHtml("#FFE4D00A");
    public static XColor ClassicRose => FromHtml("#FFFBCCE7");
    public static XColor Cobalt => FromHtml("#FF0047AB");
    public static XColor ColumbiaBlue => FromHtml("#FF9BDDFF");
    public static XColor CoolBlack => FromHtml("#FF002E63");
    public static XColor CoolGrey => FromHtml("#FF8C92AC");
    public static XColor Copper => FromHtml("#FFB87333");
    public static XColor CopperRose => FromHtml("#FF996666");
    public static XColor Coquelicot => FromHtml("#FFFF3800");
    public static XColor CoralPink => FromHtml("#FFF88379");
    public static XColor CoralRed => FromHtml("#FFFF4040");
    public static XColor Cordovan => FromHtml("#FF893F45");
    public static XColor Corn => FromHtml("#FFFBEC5D");
    public static XColor CornellRed => FromHtml("#FFB31B1B");
    public static XColor CosmicLatte => FromHtml("#FFFFF8E7");
    public static XColor CottonCandy => FromHtml("#FFFFBCD9");
    public static XColor Cream => FromHtml("#FFFFFDD0");
    public static XColor CrimsonGlory => FromHtml("#FFBE0032");
    public static XColor CyanProcess => FromHtml("#FF00B7EB");
    public static XColor Daffodil => FromHtml("#FFFFFF31");
    public static XColor Dandelion => FromHtml("#FFF0E130");
    public static XColor DarkBrown => FromHtml("#FF654321");
    public static XColor DarkByzantium => FromHtml("#FF5D3954");
    public static XColor DarkCandyAppleRed => FromHtml("#FFA40000");
    public static XColor DarkCerulean => FromHtml("#FF08457E");
    public static XColor DarkChampagne => FromHtml("#FFC2B280");
    public static XColor DarkChestnut => FromHtml("#FF986960");
    public static XColor DarkCoral => FromHtml("#FFCD5B45");
    public static XColor DarkElectricBlue => FromHtml("#FF536878");
    public static XColor DarkGreen1 => FromHtml("#FF013220");
    public static XColor DarkJungleGreen => FromHtml("#FF1A2421");
    public static XColor DarkLava => FromHtml("#FF483C32");
    public static XColor DarkLavender => FromHtml("#FF734F96");
    public static XColor DarkMidnightBlue => FromHtml("#FF003366");
    public static XColor DarkPastelBlue => FromHtml("#FF779ECB");
    public static XColor DarkPastelGreen => FromHtml("#FF03C03C");
    public static XColor DarkPastelPurple => FromHtml("#FF966FD6");
    public static XColor DarkPastelRed => FromHtml("#FFC23B22");
    public static XColor DarkPink => FromHtml("#FFE75480");
    public static XColor DarkPowderBlue => FromHtml("#FF003399");
    public static XColor DarkRaspberry => FromHtml("#FF872657");
    public static XColor DarkScarlet => FromHtml("#FF560319");
    public static XColor DarkSienna => FromHtml("#FF3C1414");
    public static XColor DarkSpringGreen => FromHtml("#FF177245");
    public static XColor DarkTan => FromHtml("#FF918151");
    public static XColor DarkTangerine => FromHtml("#FFFFA812");
    public static XColor DarkTaupe => FromHtml("#FF483C32");
    public static XColor DarkTerraCotta => FromHtml("#FFCC4E5C");
    public static XColor DartmouthGreen => FromHtml("#FF00693E");
    public static XColor DavysGrey => FromHtml("#FF555555");
    public static XColor DebianRed => FromHtml("#FFD70A53");
    public static XColor DeepCarmine => FromHtml("#FFA9203E");
    public static XColor DeepCarminePink => FromHtml("#FFEF3038");
    public static XColor DeepCarrotOrange => FromHtml("#FFE9692C");
    public static XColor DeepCerise => FromHtml("#FFDA3287");
    public static XColor DeepChampagne => FromHtml("#FFFAD6A5");
    public static XColor DeepChestnut => FromHtml("#FFB94E48");
    public static XColor DeepFuchsia => FromHtml("#FFC154C1");
    public static XColor DeepJungleGreen => FromHtml("#FF004B49");
    public static XColor DeepLilac => FromHtml("#FF9955BB");
    public static XColor DeepMagenta => FromHtml("#FFCC00CC");
    public static XColor DeepPeach => FromHtml("#FFFFCBA4");
    public static XColor DeepSaffron => FromHtml("#FFFF9933");
    public static XColor Denim => FromHtml("#FF1560BD");
    public static XColor Desert => FromHtml("#FFC19A6B");
    public static XColor DesertSand => FromHtml("#FFEDC9AF");
    public static XColor DogwoodRose => FromHtml("#FFD71868");
    public static XColor DollarBill => FromHtml("#FF85BB65");
    public static XColor Drab => FromHtml("#FF967117");
    public static XColor DukeBlue => FromHtml("#FF00009C");
    public static XColor EarthYellow => FromHtml("#FFE1A95F");
    public static XColor Ecru => FromHtml("#FFC2B280");
    public static XColor Eggplant => FromHtml("#FF614051");
    public static XColor Eggshell => FromHtml("#FFF0EAD6");
    public static XColor EgyptianBlue => FromHtml("#FF1034A6");
    public static XColor ElectricBlue => FromHtml("#FF7DF9FF");
    public static XColor ElectricCrimson => FromHtml("#FFFF003F");
    public static XColor ElectricIndigo => FromHtml("#FF6F00FF");
    public static XColor ElectricLavender => FromHtml("#FFF4BBFF");
    public static XColor ElectricLime => FromHtml("#FFCCFF00");
    public static XColor ElectricPurple => FromHtml("#FFBF00FF");
    public static XColor ElectricUltramarine => FromHtml("#FF3F00FF");
    public static XColor ElectricViolet => FromHtml("#FF8F00FF");
    public static XColor Emerald => FromHtml("#FF50C878");
    public static XColor EtonBlue => FromHtml("#FF96C8A2");
    public static XColor Fallow => FromHtml("#FFC19A6B");
    public static XColor FaluRed => FromHtml("#FF801818");
    public static XColor Fandango => FromHtml("#FFB53389");
    public static XColor FashionFuchsia => FromHtml("#FFF400A1");
    public static XColor Fawn => FromHtml("#FFE5AA70");
    public static XColor Feldgrau => FromHtml("#FF4D5D53");
    public static XColor FernGreen => FromHtml("#FF4F7942");
    public static XColor FerrariRed => FromHtml("#FFFF2800");
    public static XColor FieldDrab => FromHtml("#FF6C541E");
    public static XColor FireEngineRed => FromHtml("#FFCE2029");
    public static XColor Flame => FromHtml("#FFE25822");
    public static XColor FlamingoPink => FromHtml("#FFFC8EAC");
    public static XColor Flavescent => FromHtml("#FFF7E98E");
    public static XColor Flax => FromHtml("#FFEEDC82");
    public static XColor FluorescentOrange => FromHtml("#FFFFBF00");
    public static XColor FluorescentYellow => FromHtml("#FFCCFF00");
    public static XColor Folly => FromHtml("#FFFF004F");
    public static XColor ForestGreenTraditional => FromHtml("#FF014421");
    public static XColor FrenchBeige => FromHtml("#FFA67B5B");
    public static XColor FrenchBlue => FromHtml("#FF0072BB");
    public static XColor FrenchLilac => FromHtml("#FF86608E");
    public static XColor FrenchRose => FromHtml("#FFF64A8A");
    public static XColor FuchsiaPink => FromHtml("#FFFF77FF");
    public static XColor Fulvous => FromHtml("#FFE48400");
    public static XColor FuzzyWuzzy => FromHtml("#FFCC6666");
    public static XColor Gamboge => FromHtml("#FFE49B0F");
    public static XColor Ginger => FromHtml("#FFF9F9FF");
    public static XColor Glaucous => FromHtml("#FF6082B6");
    public static XColor GoldenBrown => FromHtml("#FF996515");
    public static XColor GoldenPoppy => FromHtml("#FFFCC200");
    public static XColor GoldenYellow => FromHtml("#FFFFDF00");
    public static XColor GoldMetallic => FromHtml("#FFD4AF37");
    public static XColor GrannySmithApple => FromHtml("#FFA8E4A0");
    public static XColor GrayAsparagus => FromHtml("#FF465945");
    public static XColor GreenPigment => FromHtml("#FF00A550");
    public static XColor GreenRyb => FromHtml("#FF66B032");
    public static XColor Grullo => FromHtml("#FFA99A86");
    public static XColor HalayaUbe => FromHtml("#FF663854");
    public static XColor HanBlue => FromHtml("#FF446CCF");
    public static XColor HanPurple => FromHtml("#FF5218FA");
    public static XColor HansaYellow => FromHtml("#FFE9D66B");
    public static XColor Harlequin => FromHtml("#FF3FFF00");
    public static XColor HarvardCrimson => FromHtml("#FFC90016");
    public static XColor HarvestGold => FromHtml("#FFDA9100");
    public static XColor Heliotrope => FromHtml("#FFDF73FF");
    public static XColor HollywoodCerise => FromHtml("#FFF400A1");
    public static XColor HookersGreen => FromHtml("#FF007000");
    public static XColor HotMagenta => FromHtml("#FFFF1DCE");
    public static XColor HunterGreen => FromHtml("#FF355E3B");
    public static XColor Iceberg => FromHtml("#FF71A6D2");
    public static XColor Icterine => FromHtml("#FFFCF75E");
    public static XColor Inchworm => FromHtml("#FFB2EC5D");
    public static XColor IndiaGreen => FromHtml("#FF138808");
    public static XColor IndianYellow => FromHtml("#FFE3A857");
    public static XColor IndigoDye => FromHtml("#FF00416A");
    public static XColor InternationalKleinBlue => FromHtml("#FF002FA7");
    public static XColor InternationalOrange => FromHtml("#FFFF4F00");
    public static XColor Iris => FromHtml("#FF5A4FCF");
    public static XColor Isabelline => FromHtml("#FFF4F0EC");
    public static XColor IslamicGreen => FromHtml("#FF009000");
    public static XColor Jade => FromHtml("#FF00A86B");
    public static XColor Jasper => FromHtml("#FFD73B3E");
    public static XColor JazzberryJam => FromHtml("#FFA50B5E");
    public static XColor Jonquil => FromHtml("#FFFADA5E");
    public static XColor JuneBud => FromHtml("#FFBDDA57");
    public static XColor JungleGreen => FromHtml("#FF29AB87");
    public static XColor KellyGreen => FromHtml("#FF4CBB17");
    public static XColor KhakiHtmlCssKhaki => FromHtml("#FFC3B091");
    public static XColor LanguidLavender => FromHtml("#FFD6CADD");
    public static XColor LapisLazuli => FromHtml("#FF26619C");
    public static XColor LaSalleGreen => FromHtml("#FF087830");
    public static XColor LaserLemon => FromHtml("#FFFEFE22");
    public static XColor Lava => FromHtml("#FFCF1020");
    public static XColor LavenderBlue => FromHtml("#FFCCCCFF");
    public static XColor LavenderFloral => FromHtml("#FFB57EDC");
    public static XColor LavenderGray => FromHtml("#FFC4C3D0");
    public static XColor LavenderIndigo => FromHtml("#FF9457EB");
    public static XColor LavenderPink => FromHtml("#FFFBAED2");
    public static XColor LavenderPurple => FromHtml("#FF967BB6");
    public static XColor LavenderRose => FromHtml("#FFFBA0E3");
    public static XColor Lemon => FromHtml("#FFFFF700");
    public static XColor LightApricot => FromHtml("#FFFDD5B1");
    public static XColor LightBrown => FromHtml("#FFB5651D");
    public static XColor LightCarminePink => FromHtml("#FFE66771");
    public static XColor LightCornflowerBlue => FromHtml("#FF93CCEA");
    public static XColor LightFuchsiaPink => FromHtml("#FFF984EF");
    public static XColor LightMauve => FromHtml("#FFDCD0FF");
    public static XColor LightPastelPurple => FromHtml("#FFB19CD9");
    public static XColor LightSalmonPink => FromHtml("#FFFF9999");
    public static XColor LightTaupe => FromHtml("#FFB38B6D");
    public static XColor LightThulianPink => FromHtml("#FFE68FAC");
    public static XColor LightYellow1 => FromHtml("#FFFFFFED");
    public static XColor Lilac => FromHtml("#FFC8A2C8");
    public static XColor LimeColorWheel => FromHtml("#FFBFFF00");
    public static XColor LincolnGreen => FromHtml("#FF195905");
    public static XColor Liver => FromHtml("#FF534B4F");
    public static XColor Lust => FromHtml("#FFE62020");
    public static XColor MacaroniAndCheese => FromHtml("#FFFFBD88");
    public static XColor MagentaDye => FromHtml("#FFCA1F7B");
    public static XColor MagentaProcess => FromHtml("#FFFF0090");
    public static XColor MagicMint => FromHtml("#FFAAF0D1");
    public static XColor Magnolia => FromHtml("#FFF8F4FF");
    public static XColor Mahogany => FromHtml("#FFC04000");
    public static XColor Maize => FromHtml("#FFFBEC5D");
    public static XColor MajorelleBlue => FromHtml("#FF6050DC");
    public static XColor Malachite => FromHtml("#FF0BDA51");
    public static XColor Manatee => FromHtml("#FF979AAA");
    public static XColor MangoTango => FromHtml("#FFFF8243");
    public static XColor MaroonX11 => FromHtml("#FFB03060");
    public static XColor Mauve => FromHtml("#FFE0B0FF");
    public static XColor Mauvelous => FromHtml("#FFEF98AA");
    public static XColor MauveTaupe => FromHtml("#FF915F6D");
    public static XColor MayaBlue => FromHtml("#FF73C2FB");
    public static XColor MeatBrown => FromHtml("#FFE5B73B");
    public static XColor MediumAquamarine1 => FromHtml("#FF66DDAA");
    public static XColor MediumCandyAppleRed => FromHtml("#FFE2062C");
    public static XColor MediumCarmine => FromHtml("#FFAF4035");
    public static XColor MediumChampagne => FromHtml("#FFF3E5AB");
    public static XColor MediumElectricBlue => FromHtml("#FF035096");
    public static XColor MediumJungleGreen => FromHtml("#FF1C352D");
    public static XColor MediumPersianBlue => FromHtml("#FF0067A5");
    public static XColor MediumRedViolet => FromHtml("#FFBB3385");
    public static XColor MediumSpringBud => FromHtml("#FFC9DC87");
    public static XColor MediumTaupe => FromHtml("#FF674C47");
    public static XColor Melon => FromHtml("#FFFDBCB4");
    public static XColor MidnightGreenEagleGreen => FromHtml("#FF004953");
    public static XColor MikadoYellow => FromHtml("#FFFFC40C");
    public static XColor Mint => FromHtml("#FF3EB489");
    public static XColor MintGreen => FromHtml("#FF98FF98");
    public static XColor ModeBeige => FromHtml("#FF967117");
    public static XColor MoonstoneBlue => FromHtml("#FF73A9C2");
    public static XColor MordantRed19 => FromHtml("#FFAE0C00");
    public static XColor MossGreen => FromHtml("#FFADDFAD");
    public static XColor MountainMeadow => FromHtml("#FF30BA8F");
    public static XColor MountbattenPink => FromHtml("#FF997A8D");
    public static XColor MsuGreen => FromHtml("#FF18453B");
    public static XColor Mulberry => FromHtml("#FFC54B8C");
    public static XColor Mustard => FromHtml("#FFFFDB58");
    public static XColor Myrtle => FromHtml("#FF21421E");
    public static XColor NadeshikoPink => FromHtml("#FFF6ADC6");
    public static XColor NapierGreen => FromHtml("#FF2A8000");
    public static XColor NaplesYellow => FromHtml("#FFFADA5E");
    public static XColor NeonCarrot => FromHtml("#FFFFA343");
    public static XColor NeonFuchsia => FromHtml("#FFFE59C2");
    public static XColor NeonGreen => FromHtml("#FF39FF14");
    public static XColor NonPhotoBlue => FromHtml("#FFA4DDED");
    public static XColor OceanBoatBlue => FromHtml("#FFCC7422");
    public static XColor Ochre => FromHtml("#FFCC7722");
    public static XColor OldGold => FromHtml("#FFCFB53B");
    public static XColor OldLavender => FromHtml("#FF796878");
    public static XColor OldMauve => FromHtml("#FF673147");
    public static XColor OldRose => FromHtml("#FFC08081");
    public static XColor OliveDrab7 => FromHtml("#FF3C341F");
    public static XColor Olivine => FromHtml("#FF9AB973");
    public static XColor Onyx => FromHtml("#FF0F0F0F");
    public static XColor OperaMauve => FromHtml("#FFB784A7");
    public static XColor OrangeColorWheel => FromHtml("#FFFF7F00");
    public static XColor OrangePeel => FromHtml("#FFFF9F00");
    public static XColor OrangeRyb => FromHtml("#FFFB9902");
    public static XColor OtterBrown => FromHtml("#FF654321");
    public static XColor OuCrimsonRed => FromHtml("#FF990000");
    public static XColor OuterSpace => FromHtml("#FF414A4C");
    public static XColor OutrageousOrange => FromHtml("#FFFF6E4A");
    public static XColor OxfordBlue => FromHtml("#FF002147");
    public static XColor PakistanGreen => FromHtml("#FF00421B");
    public static XColor PalatinateBlue => FromHtml("#FF273BE2");
    public static XColor PalatinatePurple => FromHtml("#FF682860");
    public static XColor PaleAqua => FromHtml("#FFBCD4E6");
    public static XColor PaleBrown => FromHtml("#FF987654");
    public static XColor PaleCarmine => FromHtml("#FFAF4035");
    public static XColor PaleCerulean => FromHtml("#FF9BC4E2");
    public static XColor PaleChestnut => FromHtml("#FFDDADAF");
    public static XColor PaleCopper => FromHtml("#FFDA8A67");
    public static XColor PaleCornflowerBlue => FromHtml("#FFABCDEF");
    public static XColor PaleGold => FromHtml("#FFE6BE8A");
    public static XColor PaleMagenta => FromHtml("#FFF984E5");
    public static XColor PalePink => FromHtml("#FFFADADD");
    public static XColor PaleRobinEggBlue => FromHtml("#FF96DED1");
    public static XColor PaleSilver => FromHtml("#FFC9C0BB");
    public static XColor PaleSpringBud => FromHtml("#FFECEBBD");
    public static XColor PaleTaupe => FromHtml("#FFBC987E");
    public static XColor PansyPurple => FromHtml("#FF78184A");
    public static XColor ParisGreen => FromHtml("#FF50C878");
    public static XColor PastelBlue => FromHtml("#FFAEC6CF");
    public static XColor PastelBrown => FromHtml("#FF836953");
    public static XColor PastelGray => FromHtml("#FFCFCFC4");
    public static XColor PastelGreen => FromHtml("#FF77DD77");
    public static XColor PastelMagenta => FromHtml("#FFF49AC2");
    public static XColor PastelOrange => FromHtml("#FFFFB347");
    public static XColor PastelPink => FromHtml("#FFFFD1DC");
    public static XColor PastelPurple => FromHtml("#FFB39EB5");
    public static XColor PastelRed => FromHtml("#FFFF6961");
    public static XColor PastelViolet => FromHtml("#FFCB99C9");
    public static XColor PastelYellow => FromHtml("#FFFDFD96");
    public static XColor PaynesGrey => FromHtml("#FF40404F");
    public static XColor Peach => FromHtml("#FFFFE5B4");
    public static XColor PeachOrange => FromHtml("#FFFFCC99");
    public static XColor PeachYellow => FromHtml("#FFFADFAD");
    public static XColor Pear => FromHtml("#FFD1E231");
    public static XColor Pearl => FromHtml("#FFF0EAD6");
    public static XColor Peridot => FromHtml("#FFE6E200");
    public static XColor Periwinkle => FromHtml("#FFCCCCFF");
    public static XColor PersianBlue => FromHtml("#FF1C39BB");
    public static XColor PersianGreen => FromHtml("#FF00A693");
    public static XColor PersianIndigo => FromHtml("#FF32127A");
    public static XColor PersianOrange => FromHtml("#FFD99058");
    public static XColor PersianPink => FromHtml("#FFF77FBE");
    public static XColor PersianPlum => FromHtml("#FF701C1C");
    public static XColor PersianRed => FromHtml("#FFCC3333");
    public static XColor PersianRose => FromHtml("#FFFE28A2");
    public static XColor Persimmon => FromHtml("#FFEC5800");
    public static XColor Phlox => FromHtml("#FFDF00FF");
    public static XColor PhthaloBlue => FromHtml("#FF000F89");
    public static XColor PhthaloGreen => FromHtml("#FF123524");
    public static XColor PiggyPink => FromHtml("#FFFDDDE6");
    public static XColor PineGreen => FromHtml("#FF01796F");
    public static XColor PinkOrange => FromHtml("#FFFF9966");
    public static XColor PinkPearl => FromHtml("#FFE7ACCF");
    public static XColor PinkSherbet => FromHtml("#FFF78FA7");
    public static XColor Pistachio => FromHtml("#FF93C572");
    public static XColor Platinum => FromHtml("#FFE5E4E2");
    public static XColor PlumTraditional => FromHtml("#FF8E4585");
    public static XColor PortlandOrange => FromHtml("#FFFF5A36");
    public static XColor PrincetonOrange => FromHtml("#FFFF8F00");
    public static XColor Prune => FromHtml("#FF701C1C");
    public static XColor PrussianBlue => FromHtml("#FF003153");
    public static XColor PsychedelicPurple => FromHtml("#FFDF00FF");
    public static XColor Puce => FromHtml("#FFCC8899");
    public static XColor Pumpkin => FromHtml("#FFFF7518");
    public static XColor PurpleHeart => FromHtml("#FF69359C");
    public static XColor PurpleMountainMajesty => FromHtml("#FF9678B6");
    public static XColor PurpleMunsell => FromHtml("#FF9F00C5");
    public static XColor PurplePizzazz => FromHtml("#FFFE4EDA");
    public static XColor PurpleTaupe => FromHtml("#FF50404D");
    public static XColor PurpleX11 => FromHtml("#FFA020F0");
    public static XColor RadicalRed => FromHtml("#FFFF355E");
    public static XColor Raspberry => FromHtml("#FFE30B5D");
    public static XColor RaspberryGlace => FromHtml("#FF915F6D");
    public static XColor RaspberryPink => FromHtml("#FFE25098");
    public static XColor RaspberryRose => FromHtml("#FFB3446C");
    public static XColor RawUmber => FromHtml("#FF826644");
    public static XColor RazzleDazzleRose => FromHtml("#FFFF33CC");
    public static XColor Razzmatazz => FromHtml("#FFE3256B");
    public static XColor RedMunsell => FromHtml("#FFF2003C");
    public static XColor RedNcs => FromHtml("#FFC40233");
    public static XColor RedPigment => FromHtml("#FFED1C24");
    public static XColor RedRyb => FromHtml("#FFFE2712");
    public static XColor Redwood => FromHtml("#FFAB4E52");
    public static XColor Regalia => FromHtml("#FF522D80");
    public static XColor RichBlack => FromHtml("#FF004040");
    public static XColor RichBrilliantLavender => FromHtml("#FFF1A7FE");
    public static XColor RichCarmine => FromHtml("#FFD70040");
    public static XColor RichElectricBlue => FromHtml("#FF0892D0");
    public static XColor RichLavender => FromHtml("#FFA76BCF");
    public static XColor RichLilac => FromHtml("#FFB666D2");
    public static XColor RichMaroon => FromHtml("#FFB03060");
    public static XColor RifleGreen => FromHtml("#FF414833");
    public static XColor RobinEggBlue => FromHtml("#FF00CCCC");
    public static XColor Rose => FromHtml("#FFFF007F");
    public static XColor RoseBonbon => FromHtml("#FFF9429E");
    public static XColor RoseEbony => FromHtml("#FF674846");
    public static XColor RoseGold => FromHtml("#FFB76E79");
    public static XColor RoseMadder => FromHtml("#FFE32636");
    public static XColor RosePink => FromHtml("#FFFF66CC");
    public static XColor RoseQuartz => FromHtml("#FFAA98A9");
    public static XColor RoseTaupe => FromHtml("#FF905D5D");
    public static XColor RoseVale => FromHtml("#FFAB4E52");
    public static XColor Rosewood => FromHtml("#FF65000B");
    public static XColor RossoCorsa => FromHtml("#FFD40000");
    public static XColor RoyalAzure => FromHtml("#FF0038A8");
    public static XColor RoyalBlueTraditional => FromHtml("#FF002366");
    public static XColor RoyalFuchsia => FromHtml("#FFCA2C92");
    public static XColor RoyalPurple => FromHtml("#FF7851A9");
    public static XColor Ruby => FromHtml("#FFE0115F");
    public static XColor Ruddy => FromHtml("#FFFF0028");
    public static XColor RuddyBrown => FromHtml("#FFBB6528");
    public static XColor RuddyPink => FromHtml("#FFE18E96");
    public static XColor Rufous => FromHtml("#FFA81C07");
    public static XColor Russet => FromHtml("#FF80461B");
    public static XColor Rust => FromHtml("#FFB7410E");
    public static XColor SacramentoStateGreen => FromHtml("#FF00563F");
    public static XColor SafetyOrangeBlazeOrange => FromHtml("#FFFF6700");
    public static XColor Saffron => FromHtml("#FFF4C430");
    public static XColor Salmon1 => FromHtml("#FFFF8C69");
    public static XColor SalmonPink => FromHtml("#FFFF91A4");
    public static XColor Sand => FromHtml("#FFC2B280");
    public static XColor SandDune => FromHtml("#FF967117");
    public static XColor Sandstorm => FromHtml("#FFECD540");
    public static XColor SandyTaupe => FromHtml("#FF967117");
    public static XColor Sangria => FromHtml("#FF92000A");
    public static XColor SapGreen => FromHtml("#FF507D2A");
    public static XColor Sapphire => FromHtml("#FF082567");
    public static XColor SatinSheenGold => FromHtml("#FFCBA135");
    public static XColor Scarlet => FromHtml("#FFFF2000");
    public static XColor SchoolBusYellow => FromHtml("#FFFFD800");
    public static XColor ScreaminGreen => FromHtml("#FF76FF7A");
    public static XColor SealBrown => FromHtml("#FF321414");
    public static XColor SelectiveYellow => FromHtml("#FFFFBA00");
    public static XColor Sepia => FromHtml("#FF704214");
    public static XColor Shadow => FromHtml("#FF8A795D");
    public static XColor ShamrockGreen => FromHtml("#FF009E60");
    public static XColor ShockingPink => FromHtml("#FFFC0FC0");
    public static XColor Sienna1 => FromHtml("#FF882D17");
    public static XColor Sinopia => FromHtml("#FFCB410B");
    public static XColor Skobeloff => FromHtml("#FF007474");
    public static XColor SkyMagenta => FromHtml("#FFCF71AF");
    public static XColor SmaltDarkPowderBlue => FromHtml("#FF003399");
    public static XColor SmokeyTopaz => FromHtml("#FF933D41");
    public static XColor SmokyBlack => FromHtml("#FF100C08");
    public static XColor SpiroDiscoBall => FromHtml("#FF0FC0FC");
    public static XColor SplashedWhite => FromHtml("#FFFEFDFF");
    public static XColor SpringBud => FromHtml("#FFA7FC00");
    public static XColor StPatricksBlue => FromHtml("#FF23297A");
    public static XColor StilDeGrainYellow => FromHtml("#FFFADA5E");
    public static XColor Straw => FromHtml("#FFE4D96F");
    public static XColor Sunglow => FromHtml("#FFFFCC33");
    public static XColor Sunset => FromHtml("#FFFAD6A5");
    public static XColor Tangelo => FromHtml("#FFF94D00");
    public static XColor Tangerine => FromHtml("#FFF28500");
    public static XColor TangerineYellow => FromHtml("#FFFFCC00");
    public static XColor Taupe => FromHtml("#FF483C32");
    public static XColor TaupeGray => FromHtml("#FF8B8589");
    public static XColor TeaGreen => FromHtml("#FFD0F0C0");
    public static XColor TealBlue => FromHtml("#FF367588");
    public static XColor TealGreen => FromHtml("#FF006D5B");
    public static XColor TeaRoseOrange => FromHtml("#FFF88379");
    public static XColor TeaRoseRose => FromHtml("#FFF4C2C2");
    public static XColor TennéTawny => FromHtml("#FFCD5700");
    public static XColor TerraCotta => FromHtml("#FFE2725B");
    public static XColor ThulianPink => FromHtml("#FFDE6FA1");
    public static XColor TickleMePink => FromHtml("#FFFC89AC");
    public static XColor TiffanyBlue => FromHtml("#FF0ABAB5");
    public static XColor TigersEye => FromHtml("#FFE08D3C");
    public static XColor Timberwolf => FromHtml("#FFDBD7D2");
    public static XColor TitaniumYellow => FromHtml("#FFEEE600");
    public static XColor Toolbox => FromHtml("#FF746CC0");
    public static XColor TractorRed => FromHtml("#FFFD0E35");
    public static XColor TropicalRainForest => FromHtml("#FF00755E");
    public static XColor TuftsBlue => FromHtml("#FF417DC1");
    public static XColor Tumbleweed => FromHtml("#FFDEAA88");
    public static XColor TurkishRose => FromHtml("#FFB57281");
    public static XColor Turquoise1 => FromHtml("#FF30D5C8");
    public static XColor TurquoiseBlue => FromHtml("#FF00FFEF");
    public static XColor TurquoiseGreen => FromHtml("#FFA0D6B4");
    public static XColor TuscanRed => FromHtml("#FF823535");
    public static XColor TwilightLavender => FromHtml("#FF8A496B");
    public static XColor TyrianPurple => FromHtml("#FF66023C");
    public static XColor UaBlue => FromHtml("#FF0033AA");
    public static XColor UaRed => FromHtml("#FFD9004C");
    public static XColor Ube => FromHtml("#FF8878C3");
    public static XColor UclaBlue => FromHtml("#FF536895");
    public static XColor UclaGold => FromHtml("#FFFFB300");
    public static XColor UfoGreen => FromHtml("#FF3CD070");
    public static XColor Ultramarine => FromHtml("#FF120A8F");
    public static XColor UltramarineBlue => FromHtml("#FF4166F5");
    public static XColor UltraPink => FromHtml("#FFFF6FFF");
    public static XColor Umber => FromHtml("#FF635147");
    public static XColor UnitedNationsBlue => FromHtml("#FF5B92E5");
    public static XColor UnmellowYellow => FromHtml("#FFFFFF66");
    public static XColor UpForestGreen => FromHtml("#FF014421");
    public static XColor UpMaroon => FromHtml("#FF7B1113");
    public static XColor UpsdellRed => FromHtml("#FFAE2029");
    public static XColor Urobilin => FromHtml("#FFE1AD21");
    public static XColor UscCardinal => FromHtml("#FF990000");
    public static XColor UscGold => FromHtml("#FFFFCC00");
    public static XColor UtahCrimson => FromHtml("#FFD3003F");
    public static XColor Vanilla => FromHtml("#FFF3E5AB");
    public static XColor VegasGold => FromHtml("#FFC5B358");
    public static XColor VenetianRed => FromHtml("#FFC80815");
    public static XColor Verdigris => FromHtml("#FF43B3AE");
    public static XColor Vermilion => FromHtml("#FFE34234");
    public static XColor Veronica => FromHtml("#FFA020F0");
    public static XColor Violet1 => FromHtml("#FF8F00FF");
    public static XColor VioletColorWheel => FromHtml("#FF7F00FF");
    public static XColor VioletRyb => FromHtml("#FF8601AF");
    public static XColor Viridian => FromHtml("#FF40826D");
    public static XColor VividAuburn => FromHtml("#FF922724");
    public static XColor VividBurgundy => FromHtml("#FF9F1D35");
    public static XColor VividCerise => FromHtml("#FFDA1D81");
    public static XColor VividTangerine => FromHtml("#FFFFA089");
    public static XColor VividViolet => FromHtml("#FF9F00FF");
    public static XColor WarmBlack => FromHtml("#FF004242");
    public static XColor Wenge => FromHtml("#FF645452");
    public static XColor WildBlueYonder => FromHtml("#FFA2ADD0");
    public static XColor WildStrawberry => FromHtml("#FFFF43A4");
    public static XColor WildWatermelon => FromHtml("#FFFC6C85");
    public static XColor Wisteria => FromHtml("#FFC9A0DC");
    public static XColor Xanadu => FromHtml("#FF738678");
    public static XColor YaleBlue => FromHtml("#FF0F4D92");
    public static XColor YellowMunsell => FromHtml("#FFEFCC00");
    public static XColor YellowNcs => FromHtml("#FFFFD300");
    public static XColor YellowProcess => FromHtml("#FFFFEF00");
    public static XColor YellowRyb => FromHtml("#FFFEFE33");
    public static XColor Zaffre => FromHtml("#FF0014A8");
    public static XColor ZinnwalditeBrown => FromHtml("#FF2C1608");
    public static XColor Transparent => From(Color.Transparent);

    private static Color ParseFromHtml(string htmlColor)
    {
        // Half working incorrect parser:
        // * accepts #aarrggbb, but HTML would expect #rrggbbaa
        // * doesn't accept color names
        var argb = htmlColor.AsSpan();
        if (argb[0] == '#')
            argb = argb.Slice(1);
        if (argb.Length == 8)
            return Color.FromArgb(ReadHex(argb, 0, 2), ReadHex(argb, 2, 2), ReadHex(argb, 4, 2), ReadHex(argb, 6, 2));

        if (argb.Length == 6)
            return Color.FromArgb(ReadHex(argb, 0, 2), ReadHex(argb, 2, 2), ReadHex(argb, 4, 2));

        if (argb.Length == 3)
        {
            var r = ReadHex(argb, 0, 1);
            var g = ReadHex(argb, 1, 1);
            var b = ReadHex(argb, 2, 1);
            return Color.FromArgb((r << 4) | r, (g << 4) | g, (b << 4) | b);
        }
        throw new NotSupportedException("不支持的htmlColor的颜色值");
    }
    private static Color ParseFromRgb(string rgbColor)
    {
        if (rgbColor.Length != 6)
            throw new FormatException("Color should have 6 chars.");

        var rgbSpan = rgbColor.AsSpan();
        return Color.FromArgb(ReadHex(rgbSpan, 0, 2), ReadHex(rgbSpan, 2, 2), ReadHex(rgbSpan, 4, 2));
    }
    private static int ReadHex(ReadOnlySpan<char> text, int start, int length)
    {
        var value = 0;
        for (var i = start; i < start + length; ++i)
        {
            if (!TryGetHex(text[i], out var hexDigit))
                throw new FormatException($"Unable to parse {text.ToString()}.");

            value = value * 16 + (int)hexDigit;
        }
        return value;
    }
    private static bool TryGetHex(char c, out uint hexDigit)
    {
        switch (c)
        {
            case >= '0' and <= '9':
                hexDigit = c - (uint)'0';
                return true;
            case >= 'A' and <= 'F':
                hexDigit = c - (uint)'A' + 10;
                return true;
            case >= 'a' and <= 'f':
                hexDigit = c - (uint)'a' + 10;
                return true;
            default:
                hexDigit = 0;
                return false;
        }
    }
}