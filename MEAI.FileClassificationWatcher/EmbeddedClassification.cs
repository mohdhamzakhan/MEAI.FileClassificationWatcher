using System.IO.Compression;
using System.Xml.Linq;

namespace MEAI.FileClassificationWatcher
{
    public static class EmbeddedClassification
    {
        // The level already stamped inside the file, or null if absent/unreadable.
        public static ClassificationLevel? TryRead(string path)
        {
            try
            {
                string? value = Path.GetExtension(path).ToLowerInvariant() switch
                {
                    ".docx" or ".xlsx" or ".pptx" => ReadOoxml(path),
                    ".pdf" => ReadPdf(path),
                    _ => null // legacy .doc/.xls/.ppt are binary — can't read without Office
                };
                return ClassificationLevelExtensions.Parse(value);
            }
            catch { return null; }
        }

        // Password-protected Office files are wrapped in an OLE container, not a zip.
        // We can't read them (or open them headlessly) without the password.
        public static bool IsEncryptedOoxml(string path)
        {
            var ext = Path.GetExtension(path).ToLowerInvariant();
            if (ext is not (".docx" or ".xlsx" or ".pptx")) return false;
            try
            {
                using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                var sig = new byte[8];
                return fs.Read(sig, 0, 8) == 8 &&
                       sig.SequenceEqual(new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 });
            }
            catch { return false; }
        }

        private static string? ReadOoxml(string path)
        {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var zip = new ZipArchive(fs, ZipArchiveMode.Read);
            var entry = zip.GetEntry("docProps/custom.xml");
            if (entry == null) return null;

            using var s = entry.Open();
            var doc = XDocument.Load(s);
            var prop = doc.Descendants().FirstOrDefault(e =>
                e.Name.LocalName == "property" && (string?)e.Attribute("name") == "MEAI_Classification");
            return prop?.Value.Trim();
        }

        private static string? ReadPdf(string path)
        {
            using var doc = PdfSharpCore.Pdf.IO.PdfReader.Open(
                path, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.InformationOnly);
            var v = doc.Info.Elements.GetString("/MEAI_Classification");
            return string.IsNullOrWhiteSpace(v) ? null : v;
        }
    }
}