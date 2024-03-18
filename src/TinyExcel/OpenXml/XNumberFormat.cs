using System;
using System.Diagnostics.CodeAnalysis;

namespace TinyExcel;

public struct XNumberFormat : IEquatable<XNumberFormat>
{
    public int NumberFormatId { get; set; }
    public string Format { get; set; }

    public static readonly XNumberFormat Default = new XNumberFormat() { Format = string.Empty };

    public bool Equals(XNumberFormat other)
        => this.NumberFormatId == other.NumberFormatId && this.Format == other.Format;
    public override bool Equals([NotNullWhen(true)] object other) => other is XNumberFormat && Equals((XNumberFormat)other);
    public override int GetHashCode() => HashCode.Combine(this.NumberFormatId, this.Format);
    public static bool operator ==(XNumberFormat left, XNumberFormat right) => left.Equals(right);
    public static bool operator !=(XNumberFormat left, XNumberFormat right) => !(left.Equals(right));
}
