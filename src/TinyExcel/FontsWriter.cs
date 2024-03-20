using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Xml;

namespace TinyExcel;

public class FontsWriter : OpenXmlWriter
{
    private bool knownFonts { get; set; }
    private List<XFont> sharedFonts { get; set; } = new();
    private Dictionary<XFont, int> sharedFontIndices { get; set; } = new();

    public int GetOrAddFontId(XFont font)
    {
        if (!this.sharedFontIndices.TryGetValue(font, out var refIndex))
        {
            refIndex = this.sharedFontIndices.Count;
            this.sharedFonts.Add(font);
            this.sharedFontIndices.Add(font, refIndex);
        }
        return refIndex;
    }
    public async Task Parse(XmlNode node)
    {
        await writer.WriteAsync($"<fonts count=\"{this.sharedFonts.Count}\"");
        if (this.knownFonts.HasValue)
            await writer.WriteAsync($"x14ac:knownFonts=\"{this.knownFonts.Value.ToValue()}\"");
        await writer.WriteAsync(">");
        foreach (var font in this.sharedFonts)
        {
            await this.WriteFont(writer, font);
        }
        await writer.WriteAsync("</fonts>");
        await writer.FlushAsync();
    }
    public async Task Write(StreamWriter writer)
    {
        await writer.WriteAsync($"<fonts count=\"{this.sharedFonts.Count}\"");
        if (this.knownFonts)
            await writer.WriteAsync($"x14ac:knownFonts=\"{this.knownFonts.ToValue()}\"");
        await writer.WriteAsync(">");
        foreach (var font in this.sharedFonts)
        {
            await this.WriteFont(writer, font);
        }
        await writer.WriteAsync("</fonts>");
        await writer.FlushAsync();
    }
}
