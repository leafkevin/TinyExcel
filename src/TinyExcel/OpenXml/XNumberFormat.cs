using System;
using System.IO;
using System.Threading.Tasks;

namespace TinyExcel;

public struct XNumberFormat : IEquatable<XNumberFormat>, IXRefElement
{
    public int RefId { get; set; }
    public string Format { get; set; }

    public static readonly XNumberFormat Default = new XNumberFormat() { Format = string.Empty };

    public XNumberFormat() { }

    public async Task Write(StreamWriter writer)
    {
        //<numFmt numFmtId="0" formatCode="" /> 
        //<numFmt numFmtId="176" formatCode="yyyy\-mm\-dd;@"/>
        await writer.WriteAsync($"<numFmt numFmtId=\"{this.RefId}\"");
        //Format是否需要考虑转义，比如："¥"#,##0.00;"¥"\-#,##0.00  实际：&quot;¥&quot;#,##0.00;&quot;¥&quot;\-#,##0.00
        //有的$符号会被替换，有的本地化币种中就含有$符号，不应该替换，有[]包装
        //      <numFmt numFmtId="7" formatCode="&quot;¥&quot;#,##0.00;&quot;¥&quot;\-#,##0.00"/>
        //<numFmt numFmtId="176" formatCode="yyyy\-mm\-dd\ hh:mm:ss"/>
        //<numFmt numFmtId="177" formatCode="\$#,##0.00;\-\$#,##0.00"/>
        //<numFmt numFmtId="179" formatCode="[$€-83C]#,##0.00;[Red]\-[$€-83C]#,##0.00"/>
        if (!string.IsNullOrEmpty(this.Format))
        {
            var format = this.Format.Replace("\"", "&quot;").Replace("-", "\\-");
            await writer.WriteAsync($" formatCode=\"{format}\"");
        }
        await writer.WriteAsync("/>");
    }

    public bool Equals(XNumberFormat other)
        => this.RefId == other.RefId && this.Format == other.Format;
    public override bool Equals(object other) => other is XNumberFormat && Equals((XNumberFormat)other);
    public override int GetHashCode() => HashCode.Combine(this.RefId, this.Format);
    public static bool operator ==(XNumberFormat left, XNumberFormat right) => left.Equals(right);
    public static bool operator !=(XNumberFormat left, XNumberFormat right) => !(left.Equals(right));
}