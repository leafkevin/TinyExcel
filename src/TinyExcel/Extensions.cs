using System.Drawing;

namespace TinyExcel;

static class Extensions
{
    public static string ToCamelCase(this string strValue)
    {
        if (string.IsNullOrEmpty(strValue)) return strValue;
        return strValue.Substring(0, 1).ToLower() + strValue.Substring(1);
    }
    public static string ToValue(this bool bValue) => bValue ? "1" : "0";
    public static string ToArgbString(this Color color) => color.ToArgb().ToString("X");
}
