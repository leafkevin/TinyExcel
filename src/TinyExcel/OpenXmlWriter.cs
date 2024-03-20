using System;
using System.IO;
using System.Threading.Tasks;

namespace TinyExcel;

public class OpenXmlWriter
{
    public async Task WriteFont(StreamWriter writer, XFont xFont)
    {
        await writer.WriteAsync("<font>");
        if (xFont.Bold) await writer.WriteAsync("<b/>");
        if (xFont.Italic) await writer.WriteAsync("<i/>");
        if (xFont.Underline != XFontUnderline.None)
            await writer.WriteAsync($"<u val=\"{Enum.GetName(xFont.Underline).ToCamelCase()}\"/>");
        await writer.WriteAsync($"<vertAlign val=\"{Enum.GetName(xFont.VerticalAlignment).ToCamelCase()}\"/>");
        await writer.WriteAsync($"<sz val=\"{xFont.Size}\"/>");
        await WriteColor(writer, xFont.Color);
        await writer.WriteAsync($"<name val=\"{xFont.Name}\"/>");
        await writer.WriteAsync($"<family val=\"{(int)xFont.Family}\"/>");
        await writer.WriteAsync($"<charset val=\"{(int)xFont.Charset}\"/>");
        await writer.WriteAsync("</font>");
    }
    public async Task WriteColor(StreamWriter writer, XColor xColor)
        => await this.WriteColor(writer, xColor, "color");
    public async Task<bool> WriteColor(StreamWriter writer, XColor xColor, string localName)
    {
        if (xColor.IsEmpty) return false;

        await writer.WriteAsync($"<{localName}");
        switch (xColor.ColorType)
        {
            //<x:color indexed="1" />
            case XColorType.Indexed:
                await writer.WriteAsync($" indexed={xColor.Value}");
                break;
            case XColorType.Color:
                //<x:color rgb="FF000000" />
                await writer.WriteAsync($" rgb={xColor.Color.ToArgbString()}");
                break;
            //<x:color theme="1" tint="0.3" />
            case XColorType.Theme:
                await writer.WriteAsync($" theme={xColor.Value}");
                if (xColor.Tint.HasValue)
                    await writer.WriteAsync($" tint={xColor.Tint}");
                break;
        };
        await writer.WriteAsync("/>");
        return true;
    }
    public async Task WriteBorder(StreamWriter writer, XBorder xBorder)
    {
        //<x:border diagonalUp="0" diagonalDown="0">
        await writer.WriteAsync($"<border");
        if (xBorder.DiagonalUp)
            await writer.WriteAsync($" diagonalUp=\"{xBorder.DiagonalUp.ToValue()}\"");
        if (xBorder.DiagonalDown)
            await writer.WriteAsync($" diagonalDown=\"{xBorder.DiagonalDown.ToValue()}\"");
        await writer.WriteAsync(">");

        //  <x:left style="dashDot">
        await writer.WriteAsync($"<left style=\"{Enum.GetName(xBorder.LeftStyle).ToCamelCase()}\">");
        await this.WriteColor(writer, xBorder.LeftColor);
        await writer.WriteAsync("</left>");

        ///  <x:top style="dashDot">
        await writer.WriteAsync($"<top style=\"{Enum.GetName(xBorder.TopStyle).ToCamelCase()}\">");
        await this.WriteColor(writer, xBorder.TopColor);
        await writer.WriteAsync("</top>");

        //  <x:right style="dashDot">
        await writer.WriteAsync($"<right style=\"{Enum.GetName(xBorder.RightStyle).ToCamelCase()}\">");
        await this.WriteColor(writer, xBorder.RightColor);
        await writer.WriteAsync("</right>");

        //  <x:bottom style="dashDot">
        await writer.WriteAsync($"<bottom style=\"{Enum.GetName(xBorder.BottomStyle).ToCamelCase()}\">");
        await this.WriteColor(writer, xBorder.BottomColor);
        await writer.WriteAsync("</bottom>");

        await writer.WriteAsync("</border>");
    }
    public async Task WriteAlignment(StreamWriter writer, XAlignment xAlignment)
    {
        //<x:alignment horizontal="general" vertical="bottom" textRotation="150" wrapText="0" indent="0" relativeIndent="0" justifyLastLine="0" shrinkToFit="0" readingOrder="0" />
        await writer.WriteAsync("<alignment");
        await writer.WriteAsync($" horizontal=\"{Enum.GetName(xAlignment.Horizontal).ToCamelCase()}\"");
        await writer.WriteAsync($" vertical=\"{Enum.GetName(xAlignment.Vertical).ToCamelCase()}\"");

        if (xAlignment.TextRotation > 0)
            await writer.WriteAsync($" textRotation=\"{xAlignment.TextRotation}\"");
        if (xAlignment.WrapText)
            await writer.WriteAsync($" wrapText=\"{xAlignment.WrapText.ToValue()}\"");
        if (xAlignment.Indent > 0)
            await writer.WriteAsync($" indent=\"{xAlignment.Indent}\"");
        if (xAlignment.RelativeIndent > 0)
            await writer.WriteAsync($" relativeIndent=\"{xAlignment.RelativeIndent}\"");
        if (xAlignment.JustifyLastLine)
            await writer.WriteAsync($" justifyLastLine=\"{xAlignment.JustifyLastLine.ToValue()}\"");
        if (xAlignment.ShrinkToFit)
            await writer.WriteAsync($" shrinkToFit=\"{xAlignment.ShrinkToFit}\"");
        if (xAlignment.ReadingOrder > 0)
            await writer.WriteAsync($" readingOrder=\"{(int)xAlignment.ReadingOrder}");
        await writer.WriteAsync("/>");
    }
    public async Task WriteProtection(StreamWriter writer, XProtection xProtection)
    {
        //<x:protection locked="1" hidden="0" />
        if (!(xProtection.Locked && xProtection.Hidden))
            return;
        await writer.WriteAsync("<protection");
        if (xProtection.Locked)
            await writer.WriteAsync($" locked=\"{xProtection.Locked.ToValue()}\"");
        if (xProtection.Locked)
            await writer.WriteAsync($" hidden=\"{xProtection.Hidden.ToValue()}\"");
        await writer.WriteAsync("/>");
    }
    public async Task WriteFill(StreamWriter writer, XFill xFill)
    {
        //<x:fill count="2">
        //    <x:patternFill patternType = "solid">
        //        <x:fgColor rgb = "284472C4" />
        //    </x:patternFill>
        //</x:fill>
        await writer.WriteAsync("<fill>");
        await writer.WriteAsync($"<patternFill patternType=\"{Enum.GetName(xFill.PatternType).ToCamelCase()}\">");
        if (await this.WriteColor(writer, xFill.BackgroundColor, "fgColor"))
            await writer.WriteAsync("</patternFill>");
        else await writer.WriteAsync("/>");
        await writer.WriteAsync("</fill>");
    }
    public async Task WriteNumberFormat(StreamWriter writer, XNumberFormat xNumberFormat)
    {
        //<x:numFmt numFmtId="0" formatCode="" /> 
        await writer.WriteAsync($"<numFmt numFmtId=\"{xNumberFormat.NumberFormatId}\"");
        //Format是否需要考虑转义，比如："¥"#,##0.00;"¥"\-#,##0.00  实际：&quot;¥&quot;#,##0.00;&quot;¥&quot;\-#,##0.00
        //有的$符号会被替换，有的本地化币种中就含有$符号，不应该替换，有[]包装
        //      <numFmt numFmtId="7" formatCode="&quot;¥&quot;#,##0.00;&quot;¥&quot;\-#,##0.00"/>
        //<numFmt numFmtId="176" formatCode="yyyy\-mm\-dd\ hh:mm:ss"/>
        //<numFmt numFmtId="177" formatCode="\$#,##0.00;\-\$#,##0.00"/>
        //<numFmt numFmtId="179" formatCode="[$€-83C]#,##0.00;[Red]\-[$€-83C]#,##0.00"/>
        if (!string.IsNullOrEmpty(xNumberFormat.Format))
        {
            var format = xNumberFormat.Format.Replace("\"", "&quot;").Replace("-", "\\-");
            await writer.WriteAsync($" formatCode=\"{format}\"");
        }
        await writer.WriteAsync("/>");
    }
}
