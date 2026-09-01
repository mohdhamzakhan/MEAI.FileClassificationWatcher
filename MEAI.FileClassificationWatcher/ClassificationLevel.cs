using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MEAI.FileClassificationWatcher
{
    public enum ClassificationLevel
    {
        Public,
        Confidential,
        Secret,
        TopSecret
    }

    public static class ClassificationLevelExtensions
    {
        public static string ToDisplayName(this ClassificationLevel level) => level switch
        {
            ClassificationLevel.TopSecret => "Top Secret",
            ClassificationLevel.Secret => "Secret",
            ClassificationLevel.Confidential => "Confidential",
            ClassificationLevel.Public => "Public",
            _ => level.ToString()
        };

        public static ClassificationLevel? Parse(string? value)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;
            return Enum.TryParse<ClassificationLevel>(value.Replace(" ", ""), true, out var result)
                ? result
                : null;
        }
    }
}
