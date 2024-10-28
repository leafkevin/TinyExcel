using System.Collections.Generic;
using System.Xml.Linq;
namespace TinyExcel;

public class XCellStylesPart
{
    protected List<TElement> sharedElements = new();
    protected Dictionary<TElement, int> sharedElementIndices = new();

    public bool IsEmpty => sharedElements.Count == 0;
}