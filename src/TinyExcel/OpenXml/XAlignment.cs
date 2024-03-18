using System;
using System.Diagnostics.CodeAnalysis;

namespace TinyExcel;

public struct XAlignment : IEquatable<XAlignment>
{
    public XHorizontalAlignment Horizontal { get; set; } = XHorizontalAlignment.General;
    public XVerticalAlignment Vertical { get; set; } = XVerticalAlignment.Bottom;
    public int Indent { get; set; }
    public bool JustifyLastLine { get; set; }
    public XReadingOrder ReadingOrder { get; set; } = XReadingOrder.ContextDependent;
    public int RelativeIndent { get; set; }
    public bool ShrinkToFit { get; set; }
    public int TextRotation { get; set; }
    public bool WrapText { get; set; }
    public bool TopToBottom { get; set; }

    public static readonly XAlignment Default = new XAlignment
    {
        Horizontal = XHorizontalAlignment.General,
        Vertical = XVerticalAlignment.Bottom,
        ReadingOrder = XReadingOrder.ContextDependent,
        Indent = 0,
        JustifyLastLine = false,
        RelativeIndent = 0,
        ShrinkToFit = false,
        TextRotation = 0,
        WrapText = false,
        TopToBottom = false
    };

    public XAlignment() { }

    public bool Equals(XAlignment other)
    {
        return Horizontal == other.Horizontal
            && Vertical == other.Vertical
            && Indent == other.Indent
            && JustifyLastLine == other.JustifyLastLine
            && ReadingOrder == other.ReadingOrder
            && RelativeIndent == other.RelativeIndent
            && ShrinkToFit == other.ShrinkToFit
            && TextRotation == other.TextRotation
            && WrapText == other.WrapText
            && TopToBottom == other.TopToBottom;
    }
    public override bool Equals([NotNullWhen(true)] object other) => other is XAlignment && Equals((XAlignment)other);
    public override int GetHashCode()
    {
        var hashCode = new HashCode();
        hashCode.Add(this.Horizontal);
        hashCode.Add(this.Vertical);
        hashCode.Add(this.Indent);
        hashCode.Add(this.JustifyLastLine);
        hashCode.Add(this.ReadingOrder);
        hashCode.Add(this.RelativeIndent);
        hashCode.Add(this.ShrinkToFit);
        hashCode.Add(this.TextRotation);
        hashCode.Add(this.WrapText);
        hashCode.Add(this.TopToBottom);
        return hashCode.ToHashCode();
    }
    public static bool operator ==(XAlignment left, XAlignment right) => left.Equals(right);
    public static bool operator !=(XAlignment left, XAlignment right) => !(left.Equals(right));
}
