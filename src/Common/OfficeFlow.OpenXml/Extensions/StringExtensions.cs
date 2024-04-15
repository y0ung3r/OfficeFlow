namespace OfficeFlow.OpenXml.Extensions;

internal static class StringExtensions
{
    private static readonly char[] RestrictedXmlCharacters =
    {
        '\a', '\b', '\v', '\f',
        '\x0001', '\x0002', '\x0003', '\x0004', '\x0005',
        '\x0006', '\x000E', '\x000F', '\x0010', '\x0011',
        '\x0012', '\x0013', '\x0014', '\x0015', '\x0016',
        '\x0017', '\x0018', '\x0019', '\x001A', '\x001B',
        '\x001C', '\x001E', '\x001F', '\x007F', '\x0080',
        '\x0081', '\x0082', '\x0083', '\x0084', '\x0086',
        '\x0087', '\x0088', '\x0089', '\x008A', '\x008B',
        '\x008C', '\x008D', '\x008E', '\x008F', '\x0090',
        '\x0091', '\x0092', '\x0093', '\x0094', '\x0095',
        '\x0096', '\x0097', '\x0098', '\x0099', '\x009A',
        '\x009B', '\x009C', '\x009D', '\x009E', '\x009F'
    };

    internal static string? RemoveRestrictedXmlCharacters(this string? text)
    {
        if (text is null)
            return text;

        return string.Join(
            string.Empty,
            text.Split(RestrictedXmlCharacters));
    }
}