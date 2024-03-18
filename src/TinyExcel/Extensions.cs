namespace TinyExcel;

static class Extensions
{
    public static string ToCamelCase(this string strValue)
    {
        if (string.IsNullOrEmpty(strValue)) return strValue;
        return strValue.Substring(0, 1).ToLower() + strValue.Substring(1);
    }
}
