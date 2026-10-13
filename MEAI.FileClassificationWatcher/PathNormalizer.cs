using System.Runtime.InteropServices;
using System.Text;

public static class PathNormalizer
{
    [DllImport("mpr.dll", CharSet = CharSet.Unicode)]
    private static extern int WNetGetConnection(string localName, StringBuilder remoteName, ref int length);

    private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, string> _cache =
        new(StringComparer.OrdinalIgnoreCase);

    // Maps "I:\folder\file.xlsx" -> "\\server\share\folder\file.xlsx" when I: is a mapped drive.
    // Local drives and unmapped letters are returned unchanged.
    public static string Canonical(string path)
    {
        var full = Path.GetFullPath(path);
        if (full.Length < 3 || full[1] != ':') return full; // already UNC or not drive-based

        var drive = full.Substring(0, 2); // "I:"
        var unc = _cache.GetOrAdd(drive, d =>
        {
            int len = 512;
            var sb = new StringBuilder(len);
            return WNetGetConnection(d, sb, ref len) == 0 ? sb.ToString().TrimEnd('\\') : d;
        });

        return unc == drive ? full : unc + full.Substring(2);
    }
}