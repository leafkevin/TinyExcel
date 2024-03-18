using System;
using System.Diagnostics.CodeAnalysis;

namespace TinyExcel;

public struct XBorder : IEquatable<XBorder>
{
    public XBorderStyle LeftStyle { get; set; } = XBorderStyle.None;
    public XColor LeftColor { get; set; } = XColor.Black;
    public XBorderStyle RightStyle { get; set; } = XBorderStyle.None;
    public XColor RightColor { get; set; } = XColor.Black;
    public XBorderStyle TopStyle { get; set; } = XBorderStyle.None;
    public XColor TopColor { get; set; } = XColor.Black;
    public XBorderStyle BottomStyle { get; set; } = XBorderStyle.None;
    public XColor BottomColor { get; set; } = XColor.Black;
    public XBorderStyle DiagonalStyle { get; set; } = XBorderStyle.None;
    public XColor DiagonalColor { get; set; } = XColor.Black;
    public bool DiagonalUp { get; set; }
    public bool DiagonalDown { get; set; }

    public static readonly XBorder Default = new XBorder
    {
        BottomStyle = XBorderStyle.None,
        DiagonalStyle = XBorderStyle.None,
        LeftStyle = XBorderStyle.None,
        RightStyle = XBorderStyle.None,
        TopStyle = XBorderStyle.None,
        BottomColor = XColor.Black,
        DiagonalColor = XColor.Black,
        LeftColor = XColor.Black,
        RightColor = XColor.Black,
        TopColor = XColor.Black,
        DiagonalDown = false,
        DiagonalUp = false,
    };

    public XBorder() { }

    public bool Equals(XBorder other)
    {
        return Equals(LeftStyle, LeftColor, other.LeftStyle, other.LeftColor)
            && Equals(RightStyle, RightColor, other.RightStyle, other.RightColor)
            && Equals(TopStyle, TopColor, other.TopStyle, other.TopColor)
            && Equals(BottomStyle, BottomColor, other.BottomStyle, other.BottomColor)
            && Equals(DiagonalStyle, DiagonalColor, other.DiagonalStyle, other.DiagonalColor)
            && DiagonalUp == other.DiagonalUp
            && DiagonalDown == other.DiagonalDown;
    }
    public override bool Equals([NotNullWhen(true)] object other) => other is XBorder && Equals((XBorder)other);
    private bool Equals(XBorderStyle style1, XColor color1, XBorderStyle style2, XColor color2)
    {
        return (style1 == XBorderStyle.None && style2 == XBorderStyle.None)
            || style1 == style2 && color1 == color2;
    }
    public override int GetHashCode()
    {
        var hashCode = new HashCode();
        hashCode.Add(this.LeftStyle);
        hashCode.Add(this.LeftColor);
        hashCode.Add(this.RightStyle);
        hashCode.Add(this.RightColor);
        hashCode.Add(this.TopStyle);
        hashCode.Add(this.TopColor);
        hashCode.Add(this.BottomStyle);
        hashCode.Add(this.BottomColor);
        hashCode.Add(this.DiagonalStyle);
        hashCode.Add(this.DiagonalColor);
        hashCode.Add(this.DiagonalUp);
        hashCode.Add(this.DiagonalDown);
        return hashCode.ToHashCode();
    }
    public static bool operator ==(XBorder left, XBorder right) => left.Equals(right);
    public static bool operator !=(XBorder left, XBorder right) => !(left == right);
}
