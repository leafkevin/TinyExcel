using System;
using System.IO;
using System.Threading.Tasks;

namespace TinyExcel;

public struct XProtection : IEquatable<XProtection>
{
    public bool Locked { get; set; }
    public bool Hidden { get; set; }

    public static readonly XProtection Default = new XProtection();

    public XProtection() { }

    public async Task Write(StreamWriter writer)
    {
        //<x:protection locked="1" hidden="0" />
        if (!(this.Locked && this.Hidden))
            return;
        await writer.WriteAsync("<protection");
        if (this.Locked)
            await writer.WriteAsync($" locked=\"{this.Locked.ToValue()}\"");
        if (this.Locked)
            await writer.WriteAsync($" hidden=\"{this.Hidden.ToValue()}\"");
        await writer.WriteAsync("/>");
    }

    public bool Equals(XProtection other)
        => Locked == other.Locked && Hidden == other.Hidden;
    public override bool Equals(object other)
        => other is XProtection && Equals((XProtection)other);
    public override int GetHashCode() => HashCode.Combine(this.Locked, this.Hidden);
    public static bool operator ==(XProtection left, XProtection right) => left.Equals(right);
    public static bool operator !=(XProtection left, XProtection right) => !(left.Equals(right));
}
