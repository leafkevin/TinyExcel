using System;
using System.IO;
using System.Threading.Tasks;

namespace TinyExcel;

public struct XStyle : IEquatable<XStyle>, IXRefElement
{
    public int RefId { get; set; }
    public XFont Font { get; set; } = XFont.Default;
    public XAlignment? Alignment { get; set; } = XAlignment.Default;
    public XBorder Border { get; set; } = XBorder.Default;
    public XFill Fill { get; set; } = XFill.Default;
    public bool IncludeQuotePrefix { get; set; }
    public XNumberFormat NumberFormat { get; set; } = XNumberFormat.Default;
    public XProtection Protection { get; set; } = XProtection.Default;

    public static readonly XStyle Default = new XStyle
    {
        Font = XFont.Default,
        Alignment = XAlignment.Default,
        Border = XBorder.Default,
        Fill = XFill.Default,
        IncludeQuotePrefix = false,
        NumberFormat = XNumberFormat.Default,
        Protection = XProtection.Default
    };

    public XStyle() { }

    public async Task Write(StreamWriter writer)
    {
        //<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        //<xf numFmtId="0" fontId="3" fillId="0" borderId="1" xfId="0">
        //	<alignment vertical="top" wrapText="1"/>
        //</xf>
        //<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyNumberFormat="1" applyFill="1" applyBorder="0" applyAlignment="1" applyProtection="1">
        //	<protection locked="1" hidden="0" />
        //</xf>
        //除了第一条记录，其余记录是表示应用的格式记录，所以，默认applyXXX=1，对应元素值也不等于0，如：fontId=1,fillId=2等，第一条以后的记录，applyXXX=1可省略
        await writer.WriteAsync("<xf");
        await writer.WriteAsync($" numFmtId=\"{this.NumberFormat.RefId}\"");
        await writer.WriteAsync($" fillId=\"{this.Fill.RefId}\"");
        await writer.WriteAsync($" borderId=\"{this.Border.RefId}\"");
        if (this.Alignment.HasValue)
            await this.Alignment.Value.Write(writer);

        //第一条记录
        if (this.RefId == 0)
        {
            await writer.WriteAsync("/>");
            return;
        }
        //第一条以后的记录，需要输出applyXXX
        await writer.WriteAsync($" applyNumberFormat=\"{(this.NumberFormat.RefId != 0).ToValue()}\"");
        await writer.WriteAsync($" applyFill=\"{(this.Fill.RefId != 0).ToValue()}\"");
        await writer.WriteAsync($" applyBorder=\"{(this.Border.RefId != 0).ToValue()}\"");
        await writer.WriteAsync("</xf>");
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