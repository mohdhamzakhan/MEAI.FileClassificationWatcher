using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace MEAI.FileClassificationWatcher
{
    public class WatcherConfig
    {
        // Hardcoded, ALWAYS applied regardless of what central settings say — a floor,
        // not a ceiling. Watching entire drives (including C:\ itself) means an empty or
        // misconfigured ExcludedFolders from the DB would otherwise have this watcher
        // recursively crawling C:\Windows, C:\Program Files, and every temp/cache folder
        // on the machine — that's not just noisy, it's a real risk of FileSystemWatcher's
        // internal event buffer overflowing (silently dropping events) under that volume.
        // Central settings can ADD to this list but can't remove from it.
        private static readonly string[] MinimumSafetyExclusions =
        {
            "Windows", "Program Files", "Program Files (x86)", "ProgramData",
            "$Recycle.Bin", "System Volume Information",
            "node_modules", ".git", "bin", "obj"
        };

        // Automatically populated with all local and mapped network drives.
        // If your JSON includes a "WatchedFolders" array, it will OVERWRITE this list.
        public List<string> WatchedFolders { get; set; } = new();

        public WatcherConfig()
        {
            try
            {
                foreach (var drive in DriveInfo.GetDrives())
                {
                    if (drive.IsReady && (drive.DriveType == DriveType.Fixed || drive.DriveType == DriveType.Network))
                    {
                        WatchedFolders.Add(drive.RootDirectory.FullName);
                    }
                }
            }
            catch
            {
                WatchedFolders = new List<string> { @"C:\" };
            }
        }

        // Populated entirely by appsettings.json or central API
        public HashSet<string> MonitoredExtensions { get; set; } = new(StringComparer.OrdinalIgnoreCase);

        // Populated entirely by appsettings.json or central API — MinimumSafetyExclusions
        // above is unioned in at resolve time regardless of what ends up here.
        public List<string> ExcludedFolders { get; set; } = new();

        public int LockPollIntervalSeconds { get; set; } = 2;
        public int MaxWatchMinutes { get; set; } = 240;

        public (List<string> AbsolutePrefixes, HashSet<string> NameMatches) ResolveExclusions()
        {
            var prefixes = new List<string>();
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            var combined = (ExcludedFolders ?? new List<string>()).Concat(MinimumSafetyExclusions);

            foreach (var entry in combined)
            {
                var expanded = Environment.ExpandEnvironmentVariables(entry);
                bool looksLikePath = expanded.Contains('\\') || expanded.Contains('/');
                if (looksLikePath)
                    prefixes.Add(Path.TrimEndingDirectorySeparator(Path.GetFullPath(expanded)));
                else
                    names.Add(expanded);
            }

            return (prefixes, names);
        }
    }
}