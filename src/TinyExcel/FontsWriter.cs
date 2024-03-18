using System.Collections.Generic;
using System.Threading.Tasks;
using System.Xml;

namespace TinyExcel;

public class FontsWriter : OpenXmlWriter
{
    private bool? KnownFonts { get; set; }
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

    public async Task Write(XmlWriter writer)
    {
        await writer.WriteStartElementAsync("x", "fonts", null);
        await writer.WriteAttributeStringAsync(null, "count", null, $"{this.sharedFonts.Count}");

        foreach (var font in this.sharedFonts)
        {
            await this.WriteFont(writer, font);
        }

        await writer.WriteEndElementAsync();
        await writer.FlushAsync();
    }
}
