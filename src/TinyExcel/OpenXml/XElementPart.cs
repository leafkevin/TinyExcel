using System.Collections.Generic;

namespace TinyExcel;

public class XElementPart<TElement> where TElement : struct, IXRefElement
{
    protected List<TElement> sharedElements = new();
    protected Dictionary<TElement, int> sharedElementIndices = new();

    public bool IsEmpty => sharedElements.Count == 0;

    public TElement AddElement(TElement element)
    {
        if (this.sharedElementIndices.TryGetValue(element, out var refIndex))
            return this.sharedElements[refIndex];
        element.RefId = this.sharedElementIndices.Count;
        this.sharedElements.Add(element);
        this.sharedElementIndices.Add(element, element.RefId);
        return element;
    }
}
public interface IXRefElement
{
    public int RefId { get; set; }
}