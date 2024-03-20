using System;
using System.IO;
using System.Threading.Tasks;

namespace TinyExcel;

public struct XFill : IEquatable<XFill>
{
    public XColor BackgroundColor { get; set; } = XColor.Empty;
    public XColor PatternColor { get; set; } = XColor.Empty;
    public XFillPattern PatternType { get; set; } = XFillPattern.None;

    public static readonly XFill Default = new XFill { PatternType = XFillPattern.None };
    public static readonly XFill Default1 = new XFill { PatternType = XFillPattern.Gray125 };

    public XFill() { }

    public async Task Write(StreamWriter writer, XFill xFill)
    {
        //<x:fill count="2">
        //    <x:patternFill patternType = "solid">
        //        <x:fgColor rgb = "284472C4" />
        //    </x:patternFill>
        //</x:fill>
        await writer.WriteAsync("<fill>");
        await writer.WriteAsync($"<patternFill patternType=\"{Enum.GetName(this.PatternType).ToCamelCase()}\">");
        if (await this.BackgroundColor.Write(writer, "fgColor"))
            await writer.WriteAsync("</patternFill>");
        else await writer.WriteAsync("/>");
        await writer.WriteAsync("</fill>");
    }

    public bool Equals(XFill other) => this.BackgroundColor == other.BackgroundColor
        && this.PatternColor == other.PatternColor && this.PatternType == other.PatternType;
    public override bool Equals(object other)
        => other is XFill && Equals((XFill)other);
    public override int GetHashCode() => HashCode.Combine(this.BackgroundColor, this.PatternColor, this.PatternType);
    public static bool operator ==(XFill left, XFill right) => left.Equals(right);
    public static bool operator !=(XFill left, XFill right) => !(left == right);
}
