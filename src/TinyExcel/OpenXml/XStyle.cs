using System;
using System.IO;
using System.Threading.Tasks;

namespace TinyExcel;

public struct XStyle : IEquatable<XStyle>
{
    public XFont Font { get; set; }
    public XAlignment Alignment { get; set; }
    public XBorder Border { get; set; }
    public XFill Fill { get; set; }
    public bool IncludeQuotePrefix { get; set; }
    public XNumberFormat NumberFormat { get; set; }
    public XProtection Protection { get; set; }

    public async Task Write(StreamWriter writer)
    {
        //<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyBorder="0"/>
        await writer.WriteAsync("<xf count=>");
        if (this.Bold) await writer.WriteAsync("<b/>");
        if (this.Italic) await writer.WriteAsync("<i/>");
        if (this.Underline != XFontUnderline.None)
            await writer.WriteAsync($"<u val=\"{Enum.GetName(this.Underline).ToCamelCase()}\"/>");
        await writer.WriteAsync($"<vertAlign val=\"{Enum.GetName(this.VerticalAlignment).ToCamelCase()}\"/>");
        await writer.WriteAsync($"<sz val=\"{this.Size}\"/>");
        await this.Color.Write(writer);
        await writer.WriteAsync($"<name val=\"{this.Name}\"/>");
        await writer.WriteAsync($"<family val=\"{(int)this.Family}\"/>");
        await writer.WriteAsync($"<charset val=\"{(int)this.Charset}\"/>");
        await writer.WriteAsync("</font>");
    }

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
    public override bool Equals(object other) => other is XStyle && Equals((XStyle)other);
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
