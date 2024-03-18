using System;
using System.Diagnostics.CodeAnalysis;

namespace TinyExcel;

public struct XStyle : IEquatable<XStyle>
{
    public XFont Font { get; set; }
    public XAlignment Alignment { get; set; }
    public XBorder Border { get; set; }
    public XFill Fill { get; set; }
    /// <summary>
    /// Should the text values of a cell saved to the file be prefixed by a quote (<c>'</c>) character?
    /// Has no effect if cell values is not a <see cref="XLDataType.Text"/>. Doesn't affect values during runtime,
    /// text values are returned without quote.
    /// </summary>
    public bool IncludeQuotePrefix { get; set; }
    public XNumberFormat NumberFormat { get; set; }
    public XProtection Protection { get; set; }

    public bool Equals(XStyle other)
    {
        return this.Alignment == other.Alignment
            && this.Border == other.Border
            && this.Fill == other.Fill
            && this.Font == other.Font
            && this.IncludeQuotePrefix == other.IncludeQuotePrefix
            && this.NumberFormat == other.NumberFormat
            && this.Protection == other.Protection;
    }
    public override bool Equals([NotNullWhen(true)] object other) => other is XStyle && Equals((XStyle)other);
    public override int GetHashCode()
    {
        var hashCode = new HashCode();
        hashCode.Add(this.Alignment);
        hashCode.Add(this.Border);
        hashCode.Add(this.Fill);
        hashCode.Add(this.Font);
        hashCode.Add(this.IncludeQuotePrefix);
        hashCode.Add(this.NumberFormat);
        hashCode.Add(this.Protection);
        return hashCode.ToHashCode();
    }
    public static bool operator ==(XStyle left, XStyle right) => left.Equals(right);
    public static bool operator !=(XStyle left, XStyle right) => !(left.Equals(right));
}
