using System;
using System.Threading.Tasks;
using System.Xml;

namespace TinyExcel;

public class OpenXmlWriter
{
    public async Task WriteFont(XmlWriter writer, XFont xFont)
    {
        await writer.WriteStartElementAsync("x", "font", null);

        if (xFont.Bold)
        {
            await writer.WriteStartElementAsync("x", "b", null);
            await writer.WriteEndElementAsync();
        }
        if (xFont.Italic)
        {
            await writer.WriteStartElementAsync("x", "i", null);
            await writer.WriteEndElementAsync();
        }
        if (xFont.Underline != XFontUnderline.None)
        {
            await writer.WriteStartElementAsync("x", "u", null);
            await writer.WriteAttributeStringAsync(null, "val", null, Enum.GetName(xFont.Underline).ToLower());
            await writer.WriteEndElementAsync();
        }
        await writer.WriteStartElementAsync("x", "vertAlign", null);
        await writer.WriteAttributeStringAsync(null, "val", null, Enum.GetName(xFont.Underline).ToLower());
        await writer.WriteEndElementAsync();

        await writer.WriteStartElementAsync("x", "sz", null);
        await writer.WriteAttributeStringAsync(null, "val", null, $"{xFont.Size}");
        await writer.WriteEndElementAsync();

        await WriteColor(writer, xFont.Color);

        await writer.WriteStartElementAsync("x", "name", null);
        await writer.WriteAttributeStringAsync(null, "val", null, $"{xFont.Name}");
        await writer.WriteEndElementAsync();

        await writer.WriteStartElementAsync("x", "family", null);
        await writer.WriteAttributeStringAsync(null, "val", null, $"{(int)xFont.Family}");
        await writer.WriteEndElementAsync();

        await writer.WriteStartElementAsync("x", "charset", null);
        await writer.WriteAttributeStringAsync(null, "val", null, $"{(int)xFont.Charset}");
        await writer.WriteEndElementAsync();

        await writer.WriteEndElementAsync();
    }
    public async Task WriteColor(XmlWriter writer, XColor xColor)
        => await this.WriteColor(writer, xColor, "color");
    public async Task WriteColor(XmlWriter writer, XColor xColor, string localName)
    {
        if (xColor.IsEmpty) return;

        await writer.WriteStartElementAsync("x", localName, null);
        switch (xColor.ColorType)
        {
            //<x:color indexed="1" />
            case XColorType.Indexed:
                await writer.WriteAttributeStringAsync(null, "indexed", null, $"{xColor.Value}");
                break;
            case XColorType.Color:
                //<x:color rgb="FF000000" />
                await writer.WriteAttributeStringAsync(null, "rgb", null, $"{xColor.Color.ToArgb().ToString("X")}");
                break;
            //<x:color theme="1" tint="0.3" />
            case XColorType.Theme:
                await writer.WriteAttributeStringAsync(null, "theme", null, $"{xColor.Value}");
                if (xColor.Tint.HasValue)
                    await writer.WriteAttributeStringAsync(null, "tint", null, $"{xColor.Tint}");
                break;
        };
        await writer.WriteEndElementAsync();
    }
    public async Task WriteBorder(XmlWriter writer, XBorder xBorder)
    {
        //<x:border diagonalUp="0" diagonalDown="0">
        await writer.WriteStartElementAsync("x", "border", null);
        await writer.WriteAttributeStringAsync(null, "diagonalUp", null, xBorder.DiagonalUp ? "1" : "0");
        await writer.WriteAttributeStringAsync(null, "diagonalDown", null, xBorder.DiagonalDown ? "1" : "0");

        //  <x:left style="dashDot">
        await writer.WriteStartElementAsync("x", "left", null);
        await writer.WriteAttributeStringAsync(null, "style", null, Enum.GetName(xBorder.LeftStyle).ToCamelCase());
        await this.WriteColor(writer, xBorder.LeftColor);
        await writer.WriteEndElementAsync();

        ///  <x:top style="dashDot">
        await writer.WriteStartElementAsync("x", "top", null);
        await writer.WriteAttributeStringAsync(null, "style", null, Enum.GetName(xBorder.TopStyle).ToCamelCase());
        await this.WriteColor(writer, xBorder.TopColor);
        await writer.WriteEndElementAsync();

        //  <x:right style="dashDot">
        await writer.WriteStartElementAsync("x", "right", null);
        await writer.WriteAttributeStringAsync(null, "style", null, Enum.GetName(xBorder.RightStyle).ToCamelCase());
        await this.WriteColor(writer, xBorder.RightColor);
        await writer.WriteEndElementAsync();

        //  <x:bottom style="dashDot">        
        await writer.WriteStartElementAsync("x", "bottom", null);
        await writer.WriteAttributeStringAsync(null, "style", null, Enum.GetName(xBorder.BottomStyle).ToCamelCase());
        await this.WriteColor(writer, xBorder.BottomColor);
        await writer.WriteEndElementAsync();

        await writer.WriteEndElementAsync();
    }
    public async Task WriteAlignment(XmlWriter writer, XAlignment xAlignment)
    {
        //<x:alignment horizontal="general" vertical="bottom" textRotation="150" wrapText="0" indent="0" relativeIndent="0" justifyLastLine="0" shrinkToFit="0" readingOrder="0" />
        await writer.WriteStartElementAsync("x", "alignment", null);
        await writer.WriteAttributeStringAsync(null, "horizontal", null, Enum.GetName(xAlignment.Horizontal).ToCamelCase());
        await writer.WriteAttributeStringAsync(null, "vertical", null, Enum.GetName(xAlignment.Vertical).ToCamelCase());
        await writer.WriteAttributeStringAsync(null, "textRotation", null, xAlignment.TextRotation.ToString());
        await writer.WriteAttributeStringAsync(null, "wrapText", null, xAlignment.WrapText ? "1" : "0");
        await writer.WriteAttributeStringAsync(null, "indent", null, xAlignment.Indent.ToString());
        await writer.WriteAttributeStringAsync(null, "relativeIndent", null, xAlignment.RelativeIndent.ToString());
        await writer.WriteAttributeStringAsync(null, "justifyLastLine", null, xAlignment.JustifyLastLine.ToString());
        await writer.WriteAttributeStringAsync(null, "shrinkToFit", null, xAlignment.ShrinkToFit.ToString());
        await writer.WriteAttributeStringAsync(null, "readingOrder", null, $"{(int)xAlignment.ReadingOrder}");
        await writer.WriteEndElementAsync();
    }
    public async Task WriteProtection(XmlWriter writer, XProtection xProtection)
    {
        //<x:protection locked="1" hidden="0" />
        await writer.WriteStartElementAsync("x", "protection", null);
        await writer.WriteAttributeStringAsync(null, "locked", null, xProtection.Locked ? "1" : "0");
        await writer.WriteAttributeStringAsync(null, "hidden", null, xProtection.Hidden ? "1" : "0");
        await writer.WriteEndElementAsync();
    }
    public async Task WriteFill(XmlWriter writer, XFill xFill)
    {
        //<x:fill>
        //    <x:patternFill patternType = "solid">
        //        <x:fgColor rgb = "284472C4" />
        //    </x:patternFill>
        //</x:fill>
        await writer.WriteStartElementAsync("x", "fill", null);
        await writer.WriteStartElementAsync("x", "patternFill", null);
        await writer.WriteAttributeStringAsync(null, "patternType", null, Enum.GetName(xFill.PatternType).ToCamelCase());
        await this.WriteColor(writer, xFill.BackgroundColor, "fgColor");
        await writer.WriteEndElementAsync();
        await writer.WriteEndElementAsync();
    }
    public async Task WriteNumberFormat(XmlWriter writer, XNumberFormat xNumberFormat)
    {
        //<x:numFmt numFmtId="0" formatCode="" /> 
        await writer.WriteStartElementAsync("x", "numFmt", null);
        await writer.WriteAttributeStringAsync(null, "numFmtId", null, xNumberFormat.NumberFormatId.ToString());
        //Format是否需要考虑转义，比如："¥"#,##0.00;"¥"\-#,##0.00  实际：&quot;¥&quot;#,##0.00;&quot;¥&quot;\-#,##0.00
        //有的$符号会被替换，有的本地化币种中就含有$符号，不应该替换，有[]包装
  //      <numFmt numFmtId="7" formatCode="&quot;¥&quot;#,##0.00;&quot;¥&quot;\-#,##0.00"/>
		//<numFmt numFmtId="176" formatCode="yyyy\-mm\-dd\ hh:mm:ss"/>
		//<numFmt numFmtId="177" formatCode="\$#,##0.00;\-\$#,##0.00"/>
		//<numFmt numFmtId="179" formatCode="[$€-83C]#,##0.00;[Red]\-[$€-83C]#,##0.00"/>
        var format = xNumberFormat.Format.Replace("\"", "&quot;").Replace("-", "\\-");
        await writer.WriteAttributeStringAsync(null, "formatCode", null, format);
        await writer.WriteEndElementAsync();
    }
}
