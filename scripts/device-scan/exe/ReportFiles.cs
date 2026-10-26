using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace PmwDeviceScan
{
    // Thrown when the report could not be written to the Reports folder or the desktop.
    internal sealed class SaveFailedException : Exception
    {
        public SaveFailedException() : base("The report could not be saved.") { }
    }

    // Where reports live and how their names are read and written.
    internal static class ReportFiles
    {
        private static readonly UTF8Encoding Utf8 = new UTF8Encoding(false);

        public static string Computer()
        {
            return Environment.GetEnvironmentVariable("COMPUTERNAME") ?? "PC";
        }

        private static Settings _settings;

        // The settings in force; reloaded after the Settings screen saves.
        public static Settings Current
        {
            get { return _settings ?? (_settings = Settings.Load()); }
            set { _settings = value; }
        }

        // Where reports go, as set on the Settings screen (the Reports folder beside the program by default).
        public static string PrimaryFolder()
        {
            return Current.ResolvedFolder();
        }

        // Every folder a report of this PC might be in: the chosen one, then the
        // Reports folder beside the program and the desktop, which a save falls back to.
        private static List<string> SearchFolders()
        {
            var folders = new List<string>();
            foreach (string f in new[] { PrimaryFolder(), Settings.BesideProgram(), FallbackFolder() })
                if (!folders.Any(x => string.Equals(x.TrimEnd('\\'), f.TrimEnd('\\'), StringComparison.OrdinalIgnoreCase)))
                    folders.Add(f);
            return folders;
        }

        public static string FallbackFolder()
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "PMW Reports");
        }

        // Upper case, without the characters a file name cannot hold, single spaces.
        public static string CleanPlace(string raw)
        {
            string p = (raw ?? "").ToUpperInvariant();
            p = Regex.Replace(p, @"[\[\]\\/:*?""<>|]", "");
            p = Regex.Replace(p, @"\s+", " ");
            return p.Trim();
        }

        // "[F1 IT] PC_.txt", or "PC_.txt" when no place was given.
        public static string FileName(string place)
        {
            return place.Length > 0 ? "[" + place + "] " + Computer() + "_.txt" : Computer() + "_.txt";
        }

        // Every report name in the Reports folder and the desktop fallback.
        public static List<string> ReportNames()
        {
            var names = new List<string>();
            foreach (string folder in SearchFolders())
            {
                try
                {
                    if (!Directory.Exists(folder)) continue;
                    foreach (string f in Directory.GetFiles(folder, "*.txt"))
                        names.Add(Path.GetFileName(f));
                }
                catch (Exception) { }
            }
            return names;
        }

        private static string PlaceOf(string name)
        {
            Match m = Regex.Match(name, @"^\[(.+?)\]");
            return m.Success ? m.Groups[1].Value : null;
        }

        public static string PlaceOfFile(string path)
        {
            return PlaceOf(Path.GetFileName(path)) ?? "";
        }

        // Full path of this PC's earlier report on this drive, or null for a new machine.
        public static string FindPreviousFile()
        {
            string suffix = " " + Computer() + "_.txt";
            foreach (string folder in SearchFolders())
            {
                try
                {
                    if (!Directory.Exists(folder)) continue;
                    foreach (string f in Directory.GetFiles(folder, "*.txt"))
                    {
                        string name = Path.GetFileName(f);
                        if (!name.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)) continue;
                        if (!string.IsNullOrEmpty(PlaceOf(name))) return f;
                    }
                }
                catch (Exception) { }
            }
            return null;
        }

        // Every place on the drive, sorted, without repeats.
        public static List<string> KnownPlaces()
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var places = new List<string>();
            foreach (string name in ReportNames())
            {
                string place = PlaceOf(name);
                if (!string.IsNullOrEmpty(place) && seen.Add(place)) places.Add(place);
            }
            places.Sort(StringComparer.OrdinalIgnoreCase);
            return places;
        }

        // Writes the report to the chosen folder; if that cannot be reached (a
        // network folder off the network, a read-only drive), beside the program,
        // then to the desktop. The caller compares the path with PrimaryFolder()
        // to say on the Completed screen that it fell back.
        public static string Save(List<string> lines, string place)
        {
            string name = FileName(place);
            string text = string.Join("\r\n", lines) + "\r\n";
            foreach (string folder in SearchFolders())
            {
                try
                {
                    Directory.CreateDirectory(folder);
                    string path = Path.Combine(folder, name);
                    File.WriteAllText(path, text, Utf8);
                    return path;
                }
                catch (Exception) { }
            }
            throw new SaveFailedException();
        }

        public static bool InPrimaryFolder(string path)
        {
            return string.Equals(Path.GetDirectoryName(path).TrimEnd('\\'), PrimaryFolder().TrimEnd('\\'),
                StringComparison.OrdinalIgnoreCase);
        }

        // Command-line test runs: write into the folder given.
        public static string SaveTo(string folder, List<string> lines, string place)
        {
            Directory.CreateDirectory(folder);
            string path = Path.Combine(folder, FileName(place));
            File.WriteAllText(path, string.Join("\r\n", lines) + "\r\n", Utf8);
            return path;
        }

        // The path as the drive shows it: "Reports\[F1 IT] PC_.txt" when on the exe's drive.
        public static string OnDrive(string path)
        {
            string root = Path.GetPathRoot(path);
            string exeRoot = Path.GetPathRoot(AppDomain.CurrentDomain.BaseDirectory);
            if (root != null && string.Equals(root, exeRoot, StringComparison.OrdinalIgnoreCase))
                return path.Substring(root.Length);
            return path;
        }
    }
}
