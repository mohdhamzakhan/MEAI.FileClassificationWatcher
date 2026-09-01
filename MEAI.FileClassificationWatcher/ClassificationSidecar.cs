using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace MEAI.FileClassificationWatcher
{
    // Transient bookkeeping file used ONLY while a file is being actively edited — written
    // on create/change, and deleted once the file is confirmed closed (see
    // FileClassificationService.HandlePossibleCloseAsync). It exists purely to avoid
    // re-prompting on every no-op save via LastContentHash; it is NOT the source of truth
    // for "what is this file currently classified as" — that's the API/DB, looked up by
    // DocumentGuid.
    //
    // DocumentGuid is DETERMINISTIC (derived from the file's full path), not random. This
    // matters specifically because the sidecar gets deleted on close: if the GUID were
    // random and stored only here, deleting the sidecar would sever the link to this file's
    // classification history in the DB, and the next edit session would mint a new GUID and
    // start a disconnected history. A path-derived GUID means any future session can
    // recompute the same ID and look up the file's real current classification from the API
    // even with no sidecar present at all.
    public class ClassificationSidecar
    {
        public string DocumentGuid { get; set; } = string.Empty;
        public string Classification { get; set; } = string.Empty;
        public string LastContentHash { get; set; } = string.Empty;
        public DateTime LastConfirmed { get; set; } = DateTime.Now;

        public static string SidecarPathFor(string originalPath) => originalPath + ".meaiclass.json";

        public static bool IsSidecarFile(string path) => path.EndsWith(".meaiclass.json", StringComparison.OrdinalIgnoreCase);

        // Stable identity for a file based on its path alone — no stored state required.
        // NOTE: this means a rename/move is treated as a different document (a fresh
        // classification prompt), same limitation as before, just now explicit.
        public static string DeterministicDocumentGuid(string originalPath)
        {
            var normalized = Path.GetFullPath(originalPath).ToLowerInvariant();
            using var sha = SHA256.Create();
            var hash = sha.ComputeHash(Encoding.UTF8.GetBytes(normalized));
            // Format as a GUID-shaped string purely for readability/consistency with the
            // Office add-ins' DocumentGuid values — it's just an opaque key either way.
            return new Guid(hash[..16]).ToString();
        }

        public static ClassificationSidecar? Load(string originalPath)
        {
            var path = SidecarPathFor(originalPath);
            if (!File.Exists(path)) return null;
            try
            {
                return JsonSerializer.Deserialize<ClassificationSidecar>(File.ReadAllText(path));
            }
            catch
            {
                return null;
            }
        }

        public void Save(string originalPath)
        {
            var path = SidecarPathFor(originalPath);
            File.WriteAllText(path, JsonSerializer.Serialize(this));
            try { File.SetAttributes(path, FileAttributes.Hidden); } catch { /* best effort */ }
        }

        // Called once the file is confirmed closed — the sidecar's job (avoiding
        // re-prompts during an active edit session) is done, so nothing should linger
        // on disk after the user is finished with the file.
        public static void Delete(string originalPath)
        {
            try
            {
                var path = SidecarPathFor(originalPath);
                if (File.Exists(path)) File.Delete(path);
            }
            catch
            {
                // best-effort cleanup — a leftover sidecar isn't harmful, just untidy
            }
        }
    }
}