using System.IO;
using System.Threading.Tasks;
using System.Xml;

namespace TinyExcel;

public class XNumberFormatsPart : XElementPart<XNumberFormat>
{
    public XNumberFormat AddNumberFormat(XNumberFormat numberFormat) => base.AddElement(numberFormat);
    public Task Parse(XmlNode node)
    {
        return Task.CompletedTask;
    }
    public async Task Write(StreamWriter writer)
    {
        if (this.IsEmpty) return;

        //<numFmts count="1">
        await writer.WriteAsync($"<numFmts count=\"{this.sharedElements.Count}\">");
        foreach (var format in this.sharedElements)
        {
            await format.Write(writer);
        }
        await writer.WriteAsync("</numFmts>");
        await writer.FlushAsync();
    }
}