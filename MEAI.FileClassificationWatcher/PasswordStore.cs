using System.Security.Cryptography;
using System.Text;

public static class PasswordStore
{
    private static readonly string Folder = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "MEAI", "PasswordStore");

    public static void Save(string documentGuid, string password)
    {
        Directory.CreateDirectory(Folder);
        var encrypted = ProtectedData.Protect(
            Encoding.UTF8.GetBytes(password), null, DataProtectionScope.CurrentUser);
        File.WriteAllBytes(Path.Combine(Folder, documentGuid + ".dat"), encrypted);
    }

    public static string? TryGet(string documentGuid)
    {
        var file = Path.Combine(Folder, documentGuid + ".dat");
        if (!File.Exists(file)) return null;
        try
        {
            var decrypted = ProtectedData.Unprotect(
                File.ReadAllBytes(file), null, DataProtectionScope.CurrentUser);
            return Encoding.UTF8.GetString(decrypted);
        }
        catch { return null; } // corrupted/unreadable — treat as "no password known"
    }

    public static void Delete(string documentGuid) =>
        File.Delete(Path.Combine(Folder, documentGuid + ".dat")); // no-op if missing
}