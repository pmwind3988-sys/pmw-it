using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace PmwDeviceScan
{
    // What IT set on the Settings screen, kept in PMW-Device-Scan.ini beside the
    // program. The file travels with the drive, so every PC scanned from it gets
    // the same behaviour, and the .bat and .ps1 read the same three lines.
    //
    //   ReportsFolder=            blank: the Reports folder beside the program
    //   ReportsFolder=\IT\Scans   a folder on this drive, kept WITHOUT its letter,
    //                             because the drive is E: on one PC and F: on the next
    //   ReportsFolder=\\server\PMW\IT\Device scans   a network folder
    //   AutoStart=yes             a machine scanned before starts at once
    //   AutoClose=yes             Completed closes itself after 10 seconds
    internal sealed class Settings
    {
        public const string FileName = "PMW-Device-Scan.ini";

        public string ReportsFolder = "";
        public bool AutoStart;
        public bool AutoClose = true;

        public static string ProgramFolder()
        {
            return AppDomain.CurrentDomain.BaseDirectory;
        }

        public static string FilePath()
        {
            return Path.Combine(ProgramFolder(), FileName);
        }

        public static string BesideProgram()
        {
            return Path.Combine(ProgramFolder(), "Reports");
        }

        // A missing or unreadable file is the defaults, never an error: a scan
        // must not stop because a settings file is damaged.
        public static Settings Load()
        {
            var s = new Settings();
            try
            {
                if (!File.Exists(FilePath())) return s;
                foreach (string raw in File.ReadAllLines(FilePath()))
                {
                    string line = raw.Trim();
                    if (line.Length == 0 || line.StartsWith(";") || line.StartsWith("#")) continue;
                    int eq = line.IndexOf('=');
                    if (eq < 0) continue;
                    string key = line.Substring(0, eq).Trim();
                    string value = line.Substring(eq + 1).Trim();
                    if (Is(key, "ReportsFolder")) s.ReportsFolder = value;
                    else if (Is(key, "AutoStart")) s.AutoStart = Yes(value);
                    else if (Is(key, "AutoClose")) s.AutoClose = Yes(value);
                }
            }
            catch (Exception) { }
            return s;
        }

        public void Save()
        {
            var lines = new List<string>
            {
                "; PMW Device Scan settings, written by the program's Settings screen.",
                "; The .bat and .ps1 on this drive read these too.",
                "ReportsFolder=" + ReportsFolder,
                "AutoStart=" + (AutoStart ? "yes" : "no"),
                "AutoClose=" + (AutoClose ? "yes" : "no"),
            };
            File.WriteAllText(FilePath(), string.Join("\r\n", lines) + "\r\n", new UTF8Encoding(false));
        }

        // The folder reports go to on THIS PC, with the drive letter it has here.
        public string ResolvedFolder()
        {
            string f = (ReportsFolder ?? "").Trim();
            if (f.Length == 0) return BesideProgram();
            if (f.StartsWith(@"\\")) return f;
            if (f.StartsWith(@"\")) return Path.Combine(Path.GetPathRoot(ProgramFolder()), f.TrimStart('\\'));
            if (Path.IsPathRooted(f)) return f;
            return Path.Combine(ProgramFolder(), f);
        }

        // How a folder picked with Browse is stored: on the program's own drive
        // it loses its letter; anywhere else it is kept whole.
        public static string ToStored(string picked)
        {
            if (string.IsNullOrEmpty(picked)) return "";
            string full = Path.GetFullPath(picked).TrimEnd('\\');
            if (string.Equals(full, BesideProgram().TrimEnd('\\'), StringComparison.OrdinalIgnoreCase)) return "";
            string root = Path.GetPathRoot(ProgramFolder());
            if (!full.StartsWith(@"\\") && root != null
                && full.StartsWith(root, StringComparison.OrdinalIgnoreCase))
                return @"\" + full.Substring(root.Length);
            return full;
        }

        private static bool Is(string a, string b)
        {
            return string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
        }

        private static bool Yes(string v)
        {
            return Is(v, "yes") || Is(v, "true") || v == "1" || Is(v, "on");
        }
    }
}
