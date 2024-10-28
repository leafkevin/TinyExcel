using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Xml;

namespace TinyExcel;

public class StyleSheet
{
    private XFontsPart fontsPart = new();
    private XFillsPart fillsPart = new();
    private XBordersPart bordersPart = new();
    private XNumberFormatsPart formatsPart = new();

    public Task Parse(XmlNode node)
    {
        return Task.CompletedTask;
    }
    public async Task Write(StreamWriter writer)
    {
        if (this.IsEmpty) return;

        await writer.WriteAsync($"<fonts count=\"{this.sharedElements.Count}\"");
        if (this.knownFonts)
            await writer.WriteAsync($"x14ac:knownFonts=\"{this.knownFonts.ToValue()}\"");
        await writer.WriteAsync(">");
        foreach (var font in this.sharedElements)
        {
            await font.Write(writer);
        }
        await writer.WriteAsync("</fonts>");
        await writer.FlushAsync();
    }
}