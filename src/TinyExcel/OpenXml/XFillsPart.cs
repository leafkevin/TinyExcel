using System.IO;
using System.Threading.Tasks;
using System.Xml;

namespace TinyExcel;

public class XFillsPart : XElementPart<XFill>
{
    public XFill AddFont(XFill fill) => base.AddElement(fill);
    public Task Parse(XmlNode node)
    {
        return Task.CompletedTask;
    }
    public async Task Write(StreamWriter writer)
    {
        if (this.IsEmpty) return;

        await writer.WriteAsync($"<fills count=\"{this.sharedElements.Count}\">");
        foreach (var fill in this.sharedElements)
        {
            await fill.Write(writer);
        }
        await writer.WriteAsync("</fills>");
        await writer.FlushAsync();
    }
}