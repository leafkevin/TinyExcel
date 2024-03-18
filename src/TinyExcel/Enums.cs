using System;

namespace TinyExcel;

[Flags]
public enum ApplicationType
{
    None = 0,
    Word = 1,
    Excel = 2,
    PowerPoint = 4,
    All = 7
}


public enum CellDataType { Text, Number, Boolean, DateTime, TimeSpan }
#region Color
public enum XColorType
{
    Color,
    Theme,
    Indexed
}
public enum XThemeColor
{
    Background1,
    Text1,
    Background2,
    Text2,
    Accent1,
    Accent2,
    Accent3,
    Accent4,
    Accent5,
    Accent6,
    Hyperlink,
    FollowedHyperlink
}
#endregion

#region Font
public enum XFontFamily { NotApplicable, Roman, Swiss, Modern, Script, Decorative }
public enum XFontVerticalAlignment { Baseline, Subscript, Superscript }
public enum XFontScheme
{
    /// <summary>
    /// Not a part of theme scheme.
    /// </summary>
    None = 0,

    /// <summary>
    /// A major font of a theme, generally used for headings.
    /// </summary>
    Major = 1,

    /// <summary>
    /// A minor font of a theme, generally used to body and paragraphs.
    /// </summary>
    Minor = 2
}
public enum XFontUnderline
{
    Double,
    DoubleAccounting,
    None,
    Single,
    SingleAccounting
}
public enum XFontCharSet
{
    /// <summary>
    /// ASCII character set.
    /// </summary>
    Ansi = 0,
    /// <summary>
    /// System default character set.
    /// </summary>
    Default = 1,
    /// <summary>
    /// Symbol character set.
    /// </summary>
    Symbol = 2,
    /// <summary>
    /// Characters used by Macintosh.
    /// </summary>
    Mac = 77,
    /// <summary>
    /// Japanese character set.
    /// </summary>
    ShiftJIS = 128,
    /// <summary>
    /// Korean character set.
    /// </summary>
    Hangul = 129,
    /// <summary>
    /// Another common spelling of the Korean character set.
    /// </summary>
    Hangeul = 129,
    /// <summary>
    /// Korean character set.
    /// </summary>
    Johab = 130,
    /// <summary>
    /// Chinese character set used in mainland China.
    /// </summary>
    GB2312 = 134,
    /// <summary>
    /// Chinese character set used mostly in Hong Kong SAR and Taiwan.
    /// </summary>
    ChineseBig5 = 136,
    /// <summary>
    /// Greek character set.
    /// </summary>
    Greek = 161,
    /// <summary>
    /// Turkish character set.
    /// </summary>
    Turkish = 162,
    /// <summary>
    /// Vietnamese character set.
    /// </summary>
    Vietnamese = 163,
    /// <summary>
    /// Hebrew character set.
    /// </summary>
    Hebrew = 177,
    /// <summary>
    /// Arabic character set.
    /// </summary>
    Arabic = 178,
    /// <summary>
    /// Baltic character set.
    /// </summary>
    Baltic = 186,
    /// <summary>
    /// Russian character set.
    /// </summary>
    Russian = 204,
    /// <summary>
    /// Thai character set.
    /// </summary>
    Thai = 222,
    /// <summary>
    /// Eastern European character set.
    /// </summary>
    EastEurope = 238,
    /// <summary>
    /// Extended ASCII character set used with disk operating system (DOS) and some Microsoft Windows fonts.
    /// </summary>
    Oem = 255
}
#endregion

#region Border
public enum XBorderStyle
{
    DashDot,
    DashDotDot,
    Dashed,
    Dotted,
    Double,
    Hair,
    Medium,
    MediumDashDot,
    MediumDashDotDot,
    MediumDashed,
    None,
    SlantDashDot,
    Thick,
    Thin
}
#endregion
public enum XFillPattern
{
    DarkDown,
    DarkGray,
    DarkGrid,
    DarkHorizontal,
    DarkTrellis,
    DarkUp,
    DarkVertical,
    Gray0625,
    Gray125,
    LightDown,
    LightGray,
    LightGrid,
    LightHorizontal,
    LightTrellis,
    LightUp,
    LightVertical,
    MediumGray,
    None,
    Solid
}
#region Alignment
public enum XHorizontalAlignment { Center, CenterContinuous, Distributed, Fill, General, Justify, Left, Right }
public enum XVerticalAlignment { Bottom, Center, Distributed, Justify, Top }
public enum XReadingOrder { ContextDependent, LeftToRight, RightToLeft }
#endregion