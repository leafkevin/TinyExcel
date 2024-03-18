using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TinyExcel;

public class ShardingStringPackage
{
    private int RefCount { get; set; }
    private List<string> sharedStrings { get; set; } = new();
    private Dictionary<string, int> SharedStringIndices { get; set; } = new();

    public int GetOrAddStringId(string strValue)
    {
        if (!this.SharedStringIndices.TryGetValue(strValue, out var refIndex))
        {
            refIndex = this.sharedStrings.Count;
            this.sharedStrings.Add(strValue);
            this.SharedStringIndices.Add(strValue, refIndex);
        }
        this.RefCount++;
        return refIndex;
    }
    public async Task Write(XmlWriter writer)
    {
        writer.WriteStartDocument();

        // Due to streaming and XLWorkbook structure, we don't know count before strings are written.
        // Attributes count and uniqueCount are optional thus are omitted.
        await writer.WriteStartElementAsync("x", "sst", OpenXmlConstants.Main2006SsNs);
        //这两个值count，uniqueCount是可选的，可以不设置
        await writer.WriteAttributeStringAsync(null, "count", null, $"{this.RefCount}");
        await writer.WriteAttributeStringAsync(null, "uniqueCount", null, $"{this.sharedStrings.Count}");

        foreach (var sharedString in this.sharedStrings)
        {
            await writer.WriteStartElementAsync("x", "si", null);

            //TODO:暂时没有处理PhoneticRun
            if (sharedString != null && sharedString.Trim().Length < sharedString.Length)
                writer.WriteAttributeString("xml", "space", null, "preserve");
            await writer.WriteElementStringAsync("x", "t", null, sharedString);
            await writer.WriteEndElementAsync();
        }
        await writer.WriteEndElementAsync();
        await writer.FlushAsync();
    }
    public async Task WriteTo(Stream stream)
    {
        var partUri = new Uri("/xl/sharedStrings.xml");
        var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
        using var package = Package.Open(stream, FileMode.Create);
        var packagePart = package.CreatePart(partUri, contentType, CompressionOption.Normal);

        var encoding = new UTF8Encoding(false);
        using var writer = XmlWriter.Create(packagePart.GetStream(), new XmlWriterSettings { CloseOutput = true, Encoding = encoding });
        await this.Write(writer);
        package.Flush();
        writer.Close();
        package.Close();
    }
}
