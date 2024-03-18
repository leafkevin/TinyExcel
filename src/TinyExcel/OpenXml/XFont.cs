using System;

namespace TinyExcel;

public struct XFont : IEquatable<XFont>
{
    public XFontFamily Family { get; set; } = XFontFamily.Swiss;
    public string Name { get; set; } = "Calibri";
    public double Size { get; set; } = 11;
    public XColor Color { get; set; } = XColor.FromArgb(0, 0, 0);
    public XFontVerticalAlignment VerticalAlignment { get; set; } = XFontVerticalAlignment.Baseline;
    public XFontCharSet Charset { get; set; } = XFontCharSet.GB2312;
    public XFontScheme Scheme { get; set; } = XFontScheme.None;
    public bool Bold { get; set; } = false;
    public bool Italic { get; set; } = false;
    public XFontUnderline Underline { get; set; } = XFontUnderline.None;
    public bool Strikethrough { get; set; } = false;
    public bool Shadow { get; set; } = false;
    public XFontCharSet CharSet { get; set; } = XFontCharSet.Default;

    public static readonly XFont Default = new XFont
    {
        Bold = false,
        Italic = false,
        Underline = XFontUnderline.None,
        Strikethrough = false,
        VerticalAlignment = XFontVerticalAlignment.Baseline,
        Size = 11,
        Color = XColor.FromArgb(0, 0, 0),
        Name = "Calibri",
        Family = XFontFamily.Swiss,
        CharSet = XFontCharSet.Default,
        Scheme = XFontScheme.None
    };

    public XFont() { }

    public bool Equals(XFont other)
    {
        return Bold == other.Bold && Italic == other.Italic && Underline == other.Underline && Strikethrough == other.Strikethrough
              && VerticalAlignment == other.VerticalAlignment && Shadow == other.Shadow && Size.Equals(other.Size) && Color == other.Color
              && Family == other.Family && CharSet == other.CharSet && Scheme == other.Scheme && string.Equals(Name, other.Name, StringComparison.OrdinalIgnoreCase);
    }
    public override bool Equals(object other) => other is XFont && Equals((XFont)other);
    public override int GetHashCode()
    {
        var hashCode = new HashCode();
        hashCode.Add(this.Bold);
        hashCode.Add(this.Italic);
        hashCode.Add(this.Underline);
        hashCode.Add(this.Strikethrough);
        hashCode.Add(this.VerticalAlignment);
        hashCode.Add(this.Shadow);
        hashCode.Add(this.Size);
        hashCode.Add(this.Color);
        hashCode.Add(this.Family);
        hashCode.Add(this.CharSet);
        hashCode.Add(this.Scheme);
        hashCode.Add(this.Name);
        return hashCode.ToHashCode();
    }
    public static bool operator ==(XFont left, XFont right) => left.Equals(right);
    public static bool operator !=(XFont left, XFont right) => !(left == right);
}
