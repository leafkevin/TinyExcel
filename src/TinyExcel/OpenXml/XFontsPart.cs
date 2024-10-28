using System.IO;
using System.Threading.Tasks;
using System.Xml;

namespace TinyExcel;

public class XFontsPart : XElementPart<XFont>
{
    private bool knownFonts;

    public XFont AddFont(XFont font)
    {
        //TODO: 需要处理knownFonts字段值
        var refFont = base.AddElement(font);
        return refFont;
    }
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