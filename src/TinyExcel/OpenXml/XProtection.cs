using System;
using System.Diagnostics.CodeAnalysis;

namespace TinyExcel;

public struct XProtection : IEquatable<XProtection>
{
    public bool Locked { get; set; } = true;
    public bool Hidden { get; set; }

    public static readonly XProtection Default = new XProtection();

    public XProtection() { }
    public bool Equals(XProtection other)
        => Locked == other.Locked && Hidden == other.Hidden;
    public override bool Equals([NotNullWhen(true)] object other) => other is XProtection && Equals((XProtection)other);
    public override int GetHashCode() => HashCode.Combine(this.Locked, this.Hidden);
    public static bool operator ==(XProtection left, XProtection right) => left.Equals(right);
    public static bool operator !=(XProtection left, XProtection right) => !(left.Equals(right));
}
