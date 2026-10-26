using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace PmwDeviceScan
{
    // One licence row (SoftwareLicensingProduct or OfficeSoftwareProtectionProduct).
    internal sealed class Licence
    {
        public string Name;
        public string Description;
        public string AppId;
        public string PartialKey;
        public string ProductChannel;
        public int Status;
    }

    // One physical drive, from MSFT_PhysicalDisk or the Win32_DiskDrive fallback.
    internal sealed class Disk
    {
        public string Name;
        public string Media;
        public double Size;
        public string Serial;
        public string Bus;
        public string Firmware;
    }

    // One Programs and Features entry, read from the uninstall keys.
    internal sealed class UninstallEntry
    {
        public string DisplayName;
        public string DisplayVersion;
        public string Publisher;
        public string InstallDate;
        public string SystemComponent;
        public string ParentKeyName;
        public string ReleaseType;
    }

    // Builds the report lines, in the same order and format as PMW-Device-Scan.bat.
    // Each Block() is isolated: a block that throws writes its label and nothing else.
    internal sealed class ReportBuilder
    {
        private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        private const double GbBytes = 1073741824.0;
        private const string Cimv2 = "root\\cimv2";
        private const string Storage = "root\\Microsoft\\Windows\\Storage";
        private const string StandardCimv2 = "root\\StandardCimv2";

        private static readonly Dictionary<int, string> ChassisNames = new Dictionary<int, string>
        {
            { 3, "Desktop" }, { 4, "Low profile desktop" }, { 6, "Mini tower" }, { 7, "Tower" },
            { 8, "Portable" }, { 9, "Laptop" }, { 10, "Notebook" }, { 13, "All in one" },
            { 14, "Sub notebook" }, { 30, "Tablet" }, { 31, "Convertible" }, { 32, "Detachable" },
            { 35, "Mini PC" }, { 36, "Stick PC" }
        };

        private static readonly Dictionary<int, string> MemoryTypes = new Dictionary<int, string>
        {
            { 20, "DDR" }, { 21, "DDR2" }, { 24, "DDR3" }, { 26, "DDR4" }, { 34, "DDR5" }
        };

        private static readonly Dictionary<int, string> StatusWords = new Dictionary<int, string>
        {
            { 0, "Unlicensed" }, { 1, "Licensed" }, { 2, "Grace period (out of box)" },
            { 3, "Grace period (hardware changed)" }, { 4, "Grace period (not genuine)" },
            { 5, "Notification -- not activated" }, { 6, "Grace period (extended)" }
        };

        private static RegistryKey _lm;

        private readonly List<string> _out = new List<string>();
        private readonly Action<int> _onStep;

        private ReportBuilder(Action<int> onStep)
        {
            _onStep = onStep;
        }

        // Runs the whole scan. onStep (may be null) is called with 0..3 as each of the
        // four steps starts: Hardware, Serial and part numbers, Licences, Installed software.
        public static List<string> Build(Action<int> onStep)
        {
            var b = new ReportBuilder(onStep);
            b.Run();
            return b._out;
        }

        private void Step(int n)
        {
            if (_onStep != null) _onStep(n);
        }

        private void Block(string label, Func<List<string>> body)
        {
            _out.Add(label + ":");
            List<string> values;
            try { values = body() ?? new List<string>(); }
            catch (Exception) { values = new List<string>(); }
            foreach (string v in values)
            {
                if (v == null) continue;
                string t = v.Trim();
                if (t.Length > 0) _out.Add(t);
            }
            _out.Add("");
        }

        private void Banner(string title)
        {
            _out.Add("===============================");
            _out.Add(title);
            _out.Add("===============================");
        }

        private void Run()
        {
            // Licences are read in the background from the start, as the script does.
            Task<List<Licence>> licenceTask = Task.Factory.StartNew<List<Licence>>(() => QueryLicences());

            Step(0);
            _out.Add(("Name: " + Env("PMW_OWNER")).TrimEnd());
            _out.Add("Anydesk:");
            string anydesk = AnyDeskId();
            _out.Add(string.IsNullOrEmpty(anydesk) ? "None" : anydesk);
            _out.Add(("Antivirus status: " + Env("PMW_AV")).TrimEnd());
            _out.Add(("Remarks: " + Env("PMW_REMARKS")).TrimEnd());
            _out.Add("");
            _out.Add("");
            Banner("COMPUTER INFORMATION");
            _out.Add("");

            ManagementObject cs = First(Q(Cimv2, "Win32_ComputerSystem", null));
            ManagementObject bios = First(Q(Cimv2, "Win32_BIOS", null));
            ManagementObject cspr = First(Q(Cimv2, "Win32_ComputerSystemProduct", null));
            ManagementObject board = First(Q(Cimv2, "Win32_BaseBoard", null));
            ManagementObject os = First(Q(Cimv2, "Win32_OperatingSystem", null));
            Dictionary<string, string> biosReg = ReadKey(Lm(), "HARDWARE\\DESCRIPTION\\System\\BIOS");
            Dictionary<string, string> ntReg = ReadKey(Lm(), "SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion");
            Dictionary<string, string> c2r = ReadKey(Lm(), "SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration");

            Block("Computer Name", () => One(Env("COMPUTERNAME")));
            Block("Computer Model", () => One(S(P(cs, "Model"))));
            Block("Motherboard", () => One((S(P(board, "Manufacturer")) ?? "") + " | " + (S(P(board, "Product")) ?? "")));
            Block("Windows Version", () => One(S(P(os, "Caption"))));
            Block("Processor", () => Names(Q(Cimv2, "Win32_Processor", null)));
            Block("GPU", () => Names(Q(Cimv2, "Win32_VideoController", null)));
            Block("Total RAM", () => One(GbText(Num(P(cs, "TotalPhysicalMemory"))) + " GB"));

            // SMBIOSMemoryType arrived with Windows 10; Windows 7 only fills MemoryType.
            List<ManagementObject> ramModules = Q(Cimv2, "Win32_PhysicalMemory", null);
            Block("RAM Slot Info", () =>
            {
                var r = new List<string>();
                foreach (ManagementObject m in ramModules)
                {
                    double code = Num(P(m, "SMBIOSMemoryType"));
                    if (code == 0) code = Num(P(m, "MemoryType"));
                    string type;
                    if (!MemoryTypes.TryGetValue((int)code, out type)) type = "Unknown";
                    r.Add(GbText(Num(P(m, "Capacity"))) + " GB | " + type + " | " + (S(P(m, "Speed")) ?? "") + " MHz | "
                        + (S(P(m, "Manufacturer")) ?? "").Trim() + " | " + (S(P(m, "PartNumber")) ?? "").Trim());
                }
                // Measure-Object -Sum over no arrays gives nothing at all, not 0.
                List<ManagementObject> arrays = Q(Cimv2, "Win32_PhysicalMemoryArray", null);
                string slots = "";
                if (arrays.Count > 0)
                {
                    double total = 0;
                    foreach (ManagementObject a in arrays) total += Num(P(a, "MemoryDevices"));
                    slots = ((long)total).ToString(Inv);
                }
                r.Add("Total Slots: " + slots + " | Used Slots: " + ramModules.Count.ToString(Inv));
                return r;
            });

            // Get-PhysicalDisk arrived with Windows 8. Before that, Win32_DiskDrive.
            List<Disk> disks = LoadDisks();
            Block("Storage Drives", () =>
            {
                var r = new List<string>();
                foreach (Disk d in disks) r.Add(d.Name + " | " + d.Media + " | " + GbText(d.Size) + " GB");
                return r;
            });

            Block("Network Information", NetworkLines);
            Block("Antivirus", () =>
            {
                var r = new List<string>();
                foreach (ManagementObject av in Q("root\\SecurityCenter2", "AntivirusProduct", null))
                {
                    long ps = (long)Num(P(av, "productState"));
                    r.Add((S(P(av, "displayName")) ?? "") + " | " + (((ps & 0x1000) != 0) ? "Enabled" : "Disabled"));
                }
                return r;
            });
            Block("Monitor", () => Names(Q(Cimv2, "Win32_DesktopMonitor", null)));
            Block("PMW Server and credentials", CredentialLines);
            Block("Server folder", ServerFolderLines);
            Block("Microsoft Office", () => One(Reg(c2r, "ProductReleaseIds")));

            List<UninstallEntry> uninstall = ReadUninstall();
            Block("Adobe", () =>
            {
                var r = new List<string>();
                foreach (UninstallEntry e in uninstall)
                {
                    if (Regex.IsMatch(e.DisplayName ?? "", "Adobe Acrobat|Adobe Acrobat Reader", RegexOptions.IgnoreCase))
                        r.Add((e.DisplayName ?? "") + " | " + (e.DisplayVersion ?? ""));
                }
                if (r.Count == 0) r.Add("None");
                return r;
            });
            Block("Email data files found Active or Inactive account", MailFiles);

            // ---- New in version 2: identity, part numbers, serials ----
            Step(1);
            Block("Serial Number", () =>
            {
                string s = Clean(P(bios, "SerialNumber")) ?? Clean(P(cspr, "IdentifyingNumber")) ?? Clean(Reg(biosReg, "SystemSerialNumber"));
                return One(s ?? "None");
            });
            Block("Manufacturer", () => One(Clean(P(cs, "Manufacturer"))));
            Block("Product Number", () =>
            {
                string model = Clean(P(cs, "Model"));
                string mfr = S(P(cs, "Manufacturer")) ?? "";
                bool isLenovo = Regex.IsMatch(mfr, "Lenovo", RegexOptions.IgnoreCase);
                List<object> candidates = isLenovo
                    ? new List<object> { P(cs, "Model"), P(cspr, "Name"), P(cs, "SystemSKUNumber") }
                    : new List<object> { P(cs, "SystemSKUNumber"), Reg(biosReg, "SystemSKU"), P(cspr, "Name"), P(cspr, "Version") };
                string pick = null;
                foreach (object c in candidates)
                {
                    string v = Clean(c);
                    if (v == null) continue;
                    bool looksLikeName = Regex.IsMatch(v, @"^(HP|Dell|Lenovo|ASUS|Acer|MSI|Toshiba|Samsung|LG|Fujitsu|Microsoft|ThinkPad|ThinkCentre)\s", RegexOptions.IgnoreCase) || v.Length > 30;
                    if (pick == null && !looksLikeName && (isLenovo || !string.Equals(v, model, StringComparison.OrdinalIgnoreCase)))
                        pick = v;
                }
                return One(pick ?? "None");
            });
            Block("Model Family", () =>
            {
                string v = Clean(P(cspr, "Version"));
                if (v != null && !string.Equals(v, Clean(P(cs, "Model")), StringComparison.OrdinalIgnoreCase)) return One(v);
                return new List<string>();
            });
            Block("Motherboard Serial", () => One(Clean(P(board, "SerialNumber")) ?? "None"));
            Block("Device UUID", () => One(Clean(P(cspr, "UUID"))));
            Block("BIOS", () => One((S(P(bios, "Manufacturer")) ?? "").Trim() + " | " + (S(P(bios, "SMBIOSBIOSVersion")) ?? "").Trim() + " | " + Day(P(bios, "ReleaseDate"))));
            Block("Chassis", () =>
            {
                var r = new List<string>();
                foreach (ManagementObject e in Q(Cimv2, "Win32_SystemEnclosure", null))
                {
                    var kinds = new List<string>();
                    foreach (object c in Arr(P(e, "ChassisTypes")))
                    {
                        string k;
                        if (ChassisNames.TryGetValue((int)Num(c), out k)) kinds.Add(k);
                    }
                    string kind = string.Join(", ", kinds);
                    string tag = Clean(P(e, "SMBIOSAssetTag"));
                    r.Add((kind.Length > 0 ? kind : "Unknown") + " | Asset tag: " + (tag ?? "None"));
                }
                return r;
            });
            Block("Memory Modules", () =>
            {
                var r = new List<string>();
                foreach (ManagementObject m in ramModules)
                {
                    string sn = Clean(P(m, "SerialNumber"));
                    r.Add((S(P(m, "DeviceLocator")) ?? "").Trim() + " | " + GbText(Num(P(m, "Capacity"))) + " GB | "
                        + (S(P(m, "Manufacturer")) ?? "").Trim() + " | P/N " + (S(P(m, "PartNumber")) ?? "").Trim()
                        + " | S/N " + (sn ?? "None"));
                }
                return r;
            });
            Block("Drive Serials", () =>
            {
                var r = new List<string>();
                foreach (Disk d in disks)
                {
                    string sn = Clean(d.Serial);
                    r.Add(d.Name + " | S/N " + (sn ?? "None") + " | " + d.Bus + " | Firmware " + (d.Firmware ?? "").Trim());
                }
                return r;
            });
            Block("Monitor Details", () =>
            {
                var r = new List<string>();
                foreach (ManagementObject m in Q("root\\wmi", "WmiMonitorID", null))
                {
                    string sn = Clean(Decode(P(m, "SerialNumberID")));
                    r.Add(Decode(P(m, "ManufacturerName")) + " | " + Decode(P(m, "UserFriendlyName")) + " | S/N "
                        + (sn ?? "None") + " | Made " + (S(P(m, "YearOfManufacture")) ?? "") + " week " + (S(P(m, "WeekOfManufacture")) ?? ""));
                }
                return r;
            });
            Block("Battery", () =>
            {
                var r = new List<string>();
                foreach (ManagementObject b in Q(Cimv2, "Win32_Battery", null))
                    r.Add((S(P(b, "Name")) ?? "") + " | " + (S(P(b, "DeviceID")) ?? "") + " | Charge " + (S(P(b, "EstimatedChargeRemaining")) ?? "") + "%");
                return r;
            });

            // ---- New in version 2: Windows and Office licences ----
            Step(2);
            List<Licence> licences;
            try { licences = licenceTask.Wait(90000) ? licenceTask.Result : new List<Licence>(); }
            catch (Exception) { licences = new List<Licence>(); }
            string oemKey = S(P(First(Q(Cimv2, "SoftwareLicensingService", null)), "OA3xOriginalProductKeyDescription")) ?? "";

            const string windowsAppId = "55c92734-d682-4d71-983e-d6ec3f16059f";
            const string officeAppId = "0ff1ce15-a989-479d-af46-f275c6370663";

            Block("Windows Build", () =>
            {
                string display = Reg(ntReg, "DisplayVersion");
                if (string.IsNullOrEmpty(display)) display = Reg(ntReg, "ReleaseId");
                if (string.IsNullOrEmpty(display)) display = Reg(ntReg, "CSDVersion");
                string build = Reg(ntReg, "CurrentBuildNumber");
                if (ntReg.ContainsKey("UBR") && ntReg["UBR"] != null) build = build + "." + ntReg["UBR"];
                return One((display ?? "") + " | Build " + build + " | " + (S(P(os, "OSArchitecture")) ?? "") + " | Installed " + Day(P(os, "InstallDate")));
            });

            Block("Windows License", () =>
            {
                var r = new List<string>();
                var win = licences.Where(l => string.Equals(l.AppId, windowsAppId, StringComparison.OrdinalIgnoreCase)).ToList();
                if (win.Count == 0) r.Add("None found");
                foreach (Licence l in win)
                    r.Add(Channel(l) + " | " + StatusWord(l.Status) + " | Key ends " + (l.PartialKey ?? "") + " | " + (l.Name ?? ""));
                r.Add(oemKey.Length > 0 ? "Firmware key | " + oemKey : "Firmware key | None");
                return r;
            });

            Block("Office License", () =>
            {
                var r = new List<string>();
                bool found = false;
                foreach (string raw in (S(c2r.ContainsKey("ProductReleaseIds") ? c2r["ProductReleaseIds"] : null) ?? "").Split(','))
                {
                    if (raw.Length == 0) continue;
                    found = true;
                    string kind;
                    if (Regex.IsMatch(raw, "^O365|^M365", RegexOptions.IgnoreCase)) kind = "Subscription (Microsoft 365)";
                    else if (Regex.IsMatch(raw, "Volume$", RegexOptions.IgnoreCase)) kind = "Perpetual, volume";
                    else if (Regex.IsMatch(raw, "Retail$", RegexOptions.IgnoreCase)) kind = "Perpetual, retail";
                    else kind = "Unknown";
                    r.Add(raw + " | " + kind + " | Click-to-Run");
                }
                foreach (Licence l in licences.Where(x => string.Equals(x.AppId, officeAppId, StringComparison.OrdinalIgnoreCase)).ToList())
                {
                    found = true;
                    r.Add((l.Name ?? "") + " | " + Channel(l) + " | " + StatusWord(l.Status) + " | Key ends " + (l.PartialKey ?? ""));
                }
                // Office 2010 and older MSI installs use their own service.
                foreach (ManagementObject o in Q(Cimv2, "OfficeSoftwareProtectionProduct", "PartialProductKey IS NOT NULL"))
                {
                    Licence l = LicenceFrom(o);
                    found = true;
                    r.Add((l.Name ?? "") + " | " + Channel(l) + " | " + StatusWord(l.Status) + " | Key ends " + (l.PartialKey ?? ""));
                }
                // The products a signed-in Microsoft 365 account has licensed for this user.
                foreach (string name in ValueNames(Hkcu(), "Software\\Microsoft\\Office\\16.0\\Common\\Licensing\\LicensingNext"))
                {
                    if (Regex.IsMatch(name, "^PS|^InstalledGraceKey$|^MigrationToV5Done$", RegexOptions.IgnoreCase)) continue;
                    found = true;
                    r.Add("Signed-in licence | " + name);
                }
                if (!found) r.Add("None found");
                return r;
            });

            Block("Office Version", () =>
            {
                if (string.IsNullOrEmpty(Reg(c2r, "VersionToReport"))) return One("None");
                string[] keys = { "492350f6", "64256afe", "55336b82", "7ffbc6bf", "b8f9b850", "5440fd1f", "f2e724c1", "5030841d", "7983bac0" };
                string[] names = { "Current Channel", "Current Channel (Preview)", "Monthly Enterprise Channel",
                    "Semi-Annual Enterprise Channel", "Semi-Annual Enterprise Channel (Preview)", "Beta Channel",
                    "LTSC 2019", "LTSC 2021", "LTSC 2024" };
                string url = (Reg(c2r, "UpdateChannel") ?? "") + " " + (Reg(c2r, "CDNBaseUrl") ?? "");
                string channel = "Unknown channel";
                for (int i = 0; i < keys.Length; i++)
                    if (url.IndexOf(keys[i], StringComparison.OrdinalIgnoreCase) >= 0) channel = names[i];
                return One(Reg(c2r, "VersionToReport") + " | " + (Reg(c2r, "Platform") ?? "") + " | " + channel);
            });

            Block("Office Accounts", () =>
            {
                const string identity = "Software\\Microsoft\\Office\\16.0\\Common\\Identity";
                var lines = new List<string>();
                foreach (string id in SubKeyNames(Hkcu(), identity + "\\Identities"))
                {
                    Dictionary<string, string> p = ReadKey(Hkcu(), identity + "\\Identities\\" + id);
                    string email = (Reg(p, "EmailAddress") ?? "").Trim();
                    if (email.Length == 0) continue;
                    string provider = Reg(p, "ProviderId") ?? "";
                    string kind;
                    if (Regex.IsMatch(provider, "AD|OrgId", RegexOptions.IgnoreCase)) kind = "Work or school";
                    else if (Regex.IsMatch(provider, "LiveId|MSA", RegexOptions.IgnoreCase)) kind = "Personal Microsoft account";
                    else kind = "Unknown";
                    lines.Add(email + " | " + kind);
                }
                List<string> profiles = SubKeyNames(Hkcu(), identity + "\\Profiles").ToList();
                int work = profiles.Count(x => Regex.IsMatch(x, "_ADAL$", RegexOptions.IgnoreCase));
                int personal = profiles.Count(x => Regex.IsMatch(x, "_LiveId$", RegexOptions.IgnoreCase));
                if (work > 0) lines.Add("Work or school sign-in | " + work.ToString(Inv) + " profile(s)");
                if (personal > 0) lines.Add("Personal Microsoft account sign-in | " + personal.ToString(Inv) + " profile(s)");
                var unique = new List<string>();
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (string l in lines) if (seen.Add(l)) unique.Add(l);
                if (unique.Count == 0) unique.Add("None");
                return unique;
            });

            // Which account each Microsoft 365 licence on this PC belongs to. Same
            // rules as the script: one licence file per licence (base64 JSON in a
            // UTF-16 wrapper), the account's email from the registry. "Valid until"
            // is the next online check-in, not the end of the subscription.
            Block("Office Subscription", () =>
            {
                Dictionary<string, string> emails = ReadKey(Hkcu(),
                    "Software\\Microsoft\\Office\\16.0\\Common\\Licensing\\LicensingNext\\LicenseIdToEmailMapping");
                var lines = new List<string>();
                string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Microsoft\\Office\\Licenses");
                string[] files = new string[0];
                try { if (Directory.Exists(folder)) files = Directory.GetFiles(folder, "*", SearchOption.AllDirectories); }
                catch (Exception) { }
                foreach (string file in files)
                {
                    try
                    {
                        string raw = File.ReadAllText(file, Encoding.Unicode);
                        Match lic = Regex.Match(raw, "\"License\"\\s*:\\s*\"([^\"]+)\"");
                        if (!lic.Success) continue;
                        byte[] bytes = Convert.FromBase64String(lic.Groups[1].Value);
                        string json = Encoding.UTF8.GetString(bytes);
                        if (!Regex.IsMatch(json, "^\\s*\\{")) json = Encoding.Unicode.GetString(bytes);
                        string userId = JsonField(json, "UserId");
                        string email = Reg(emails, userId);
                        string tenant = JsonField(json, "TenantId");
                        Match day = Regex.Match(JsonField(json, "NotAfter"), "^(\\d{4}-\\d{2}-\\d{2})");
                        string line = JsonField(json, "ProductReleaseId") + " | " + JsonField(json, "LicenseType") + " | "
                            + (string.IsNullOrEmpty(email) ? "Unknown account" : email)
                            + " | Tenant " + (tenant.Length > 0 ? tenant : "Unknown")
                            + " | Valid until " + (day.Success ? day.Groups[1].Value : "Unknown")
                            + " | " + JsonField(json, "Status");
                        if (JsonField(json, "IsTrial") == "true") line += " | Trial";
                        lines.Add(line);
                    }
                    catch (Exception) { }
                }
                if (lines.Count == 0)
                {
                    foreach (string email in emails.Values.Distinct(StringComparer.OrdinalIgnoreCase))
                        lines.Add("Unknown | Unknown | " + email + " | Tenant Unknown | Valid until Unknown | Unknown");
                }
                if (lines.Count == 0) lines.Add("None");
                return lines;
            });

            Block("Device Join", () =>
            {
                var state = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (string line in RunLines("dsregcmd.exe", "/status"))
                {
                    Match m = Regex.Match(line, @"^\s*(AzureAdJoined|DomainJoined|WorkplaceJoined|TenantName)\s*:\s*(.*)$", RegexOptions.IgnoreCase);
                    if (m.Success) state[m.Groups[1].Value] = m.Groups[2].Value.Trim();
                }
                var r = new List<string>();
                if (state.Count > 0)
                {
                    r.Add("Azure AD joined | " + G(state, "AzureAdJoined"));
                    r.Add("Domain joined | " + G(state, "DomainJoined"));
                    r.Add("Workplace joined | " + G(state, "WorkplaceJoined"));
                    if (G(state, "TenantName").Length > 0) r.Add("Tenant | " + G(state, "TenantName"));
                }
                return r;
            });

            // ---- New in version 2: installed software ----
            Step(3);
            Block("Installed Software", () =>
            {
                var r = new List<string>();
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var rows = uninstall
                    .Where(e => e.SystemComponent != "1")
                    .Where(e => string.IsNullOrEmpty(e.ParentKeyName))
                    .Where(e => !Regex.IsMatch(e.ReleaseType ?? "", "Update|Hotfix|ServicePack", RegexOptions.IgnoreCase))
                    .Where(e => !Regex.IsMatch(e.DisplayName ?? "", @"^(Update for|Security Update|Hotfix)|\(KB\d+", RegexOptions.IgnoreCase))
                    .OrderBy(e => e.DisplayName ?? "", StringComparer.OrdinalIgnoreCase);
                foreach (UninstallEntry e in rows)
                {
                    string name = Regex.Replace((e.DisplayName ?? "").Trim(), @"\s*\|\s*", " / ").Replace(":", " -");
                    string key = name + "|" + (e.DisplayVersion ?? "");
                    if (!seen.Add(key)) continue;
                    r.Add(name + " | " + (e.DisplayVersion ?? "").Trim() + " | " + (e.Publisher ?? "").Trim().Replace("|", "/") + " | " + RegDay(e.InstallDate));
                }
                return r;
            });

            bool isAdmin = false;
            try { isAdmin = new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator); }
            catch (Exception) { isAdmin = false; }
            Block("Scan Info", () => One("Version 2 exe | " + DateTime.Now.ToString("yyyy-MM-dd HH:mm", Inv)
                + " | Run by " + Env("USERDOMAIN") + "\\" + Env("USERNAME") + " | Admin " + (isAdmin ? "Yes" : "No")));

            Banner("END OF REPORT");
        }

        // ------------------------------------------------------------------
        //  Small helpers
        // ------------------------------------------------------------------

        private static string Env(string name)
        {
            return Environment.GetEnvironmentVariable(name) ?? "";
        }

        private static List<string> One(string s)
        {
            var l = new List<string>();
            if (s != null) l.Add(s);
            return l;
        }

        private static List<string> Names(List<ManagementObject> rows)
        {
            var l = new List<string>();
            foreach (ManagementObject o in rows) l.Add(S(P(o, "Name")) ?? "");
            return l;
        }

        private static string S(object v)
        {
            if (v == null) return null;
            return Convert.ToString(v, Inv);
        }

        private static string Reg(Dictionary<string, string> d, string name)
        {
            string v;
            return d != null && d.TryGetValue(name, out v) ? v : null;
        }

        private static string G(Dictionary<string, string> d, string name)
        {
            string v = Reg(d, name);
            return v ?? "";
        }

        private static double Num(object v)
        {
            if (v == null) return 0;
            try { return Convert.ToDouble(v, Inv); }
            catch (Exception) { return 0; }
        }

        private static bool IsTrue(object v)
        {
            if (v is bool) return (bool)v;
            return Num(v) != 0;
        }

        private static string GbText(double bytes)
        {
            return ((long)Math.Round(bytes / GbBytes, 0)).ToString(Inv);
        }

        private static List<object> Arr(object v)
        {
            var l = new List<object>();
            Array a = v as Array;
            if (a != null) foreach (object x in a) l.Add(x);
            return l;
        }

        // A value as the scan writes it: BIOS filler and repeated characters are no value.
        private static string Clean(object value)
        {
            if (value == null) return null;
            string v = (S(value) ?? "").Trim().TrimEnd('.');
            string[] junk =
            {
                "To Be Filled By O.E.M.", "To be filled by O.E.M", "Default string", "Not Specified",
                "System Serial Number", "System Product Name", "System Version", "Unknown", "None",
                "N/A", "NA", "0", "00000000", "INVALID", "Not Available", "Not Supported",
                "OEM_Define0", "OEM_Define1", "OEM_Define2", "Chassis Serial Number", "Base Board Serial Number"
            };
            if (v.Length == 0) return null;
            foreach (string j in junk)
                if (string.Equals(j, v, StringComparison.OrdinalIgnoreCase)) return null;
            bool repeated = true;
            for (int i = 1; i < v.Length; i++)
                if (char.ToLowerInvariant(v[i]) != char.ToLowerInvariant(v[0])) { repeated = false; break; }
            if (repeated) return null;
            return v;
        }

        private static string Day(object value)
        {
            string s = S(value);
            if (string.IsNullOrEmpty(s)) return "";
            try
            {
                DateTime d = Regex.IsMatch(s, @"^\d{14}\.")
                    ? ManagementDateTimeConverter.ToDateTime(s)
                    : DateTime.Parse(s, Inv);
                return d.ToString("yyyy-MM-dd", Inv);
            }
            catch (Exception) { return ""; }
        }

        private static string RegDay(string value)
        {
            Match m = Regex.Match(value ?? "", @"^(\d{4})(\d{2})(\d{2})$");
            if (!m.Success) return "";
            return m.Groups[1].Value + "-" + m.Groups[2].Value + "-" + m.Groups[3].Value;
        }

        private static string Decode(object codes)
        {
            var sb = new StringBuilder();
            foreach (object c in Arr(codes))
            {
                int code = (int)Num(c);
                if (code > 0) sb.Append((char)code);
            }
            return sb.ToString().Trim();
        }

        private static ManagementObject First(List<ManagementObject> rows)
        {
            return rows != null && rows.Count > 0 ? rows[0] : null;
        }

        private static object P(ManagementBaseObject o, string name)
        {
            if (o == null) return null;
            try { return o[name]; }
            catch (Exception) { return null; }
        }

        // WMI rows, or null when the class or namespace is not there.
        private static List<ManagementObject> TryQ(string ns, string cls, string where)
        {
            var list = new List<ManagementObject>();
            try
            {
                string q = "SELECT * FROM " + cls;
                if (!string.IsNullOrEmpty(where)) q += " WHERE " + where;
                using (var searcher = new ManagementObjectSearcher(ns, q))
                using (ManagementObjectCollection rows = searcher.Get())
                {
                    foreach (ManagementObject o in rows)
                    {
                        // Load every property now: some objects read lazily, and the collection is disposed below.
                        try { o.Get(); } catch (Exception) { }
                        list.Add(o);
                    }
                }
            }
            catch (Exception)
            {
                return null;
            }
            return list;
        }

        private static List<ManagementObject> Q(string ns, string cls, string where)
        {
            return TryQ(ns, cls, where) ?? new List<ManagementObject>();
        }

        private static RegistryKey Lm()
        {
            if (_lm == null)
            {
                _lm = Environment.Is64BitOperatingSystem
                    ? RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64)
                    : Registry.LocalMachine;
            }
            return _lm;
        }

        private static RegistryKey Hkcu()
        {
            return Registry.CurrentUser;
        }

        // One value out of a flat JSON object, quoted or not; "" when absent.
        // Enough for Office's licence files, which are not nested where this reads.
        private static string JsonField(string json, string name)
        {
            Match m = Regex.Match(json, "\"" + Regex.Escape(name) + "\"\\s*:\\s*\"?([^\",}]*)\"?");
            return m.Success ? m.Groups[1].Value.Trim() : "";
        }

        private static Dictionary<string, string> ReadKey(RegistryKey root, string path)
        {
            var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using (RegistryKey k = root.OpenSubKey(path))
                {
                    if (k == null) return d;
                    foreach (string n in k.GetValueNames()) d[n] = S(k.GetValue(n));
                }
            }
            catch (Exception) { }
            return d;
        }

        private static List<string> ValueNames(RegistryKey root, string path)
        {
            try
            {
                using (RegistryKey k = root.OpenSubKey(path))
                {
                    if (k == null) return new List<string>();
                    return k.GetValueNames().Where(n => n.Length > 0).ToList();
                }
            }
            catch (Exception) { return new List<string>(); }
        }

        private static List<string> SubKeyNames(RegistryKey root, string path)
        {
            try
            {
                using (RegistryKey k = root.OpenSubKey(path))
                {
                    if (k == null) return new List<string>();
                    return k.GetSubKeyNames().ToList();
                }
            }
            catch (Exception) { return new List<string>(); }
        }

        private static Encoding OemEncoding()
        {
            try { return Encoding.GetEncoding((int)GetOEMCP()); }
            catch (Exception) { return Encoding.Default; }
        }

        [DllImport("kernel32.dll")]
        private static extern uint GetOEMCP();

        // Runs a Windows tool and returns its output lines, or nothing after 20 seconds.
        private static List<string> RunLines(string exe, string args)
        {
            var lines = new List<string>();
            try
            {
                var psi = new ProcessStartInfo(Path.Combine(Environment.SystemDirectory, exe), args);
                psi.UseShellExecute = false;
                psi.RedirectStandardOutput = true;
                psi.RedirectStandardError = true;
                psi.CreateNoWindow = true;
                psi.StandardOutputEncoding = OemEncoding();
                psi.StandardErrorEncoding = OemEncoding();
                using (Process p = Process.Start(psi))
                {
                    Task<string> outTask = Task.Factory.StartNew<string>(() => p.StandardOutput.ReadToEnd());
                    Task<string> errTask = Task.Factory.StartNew<string>(() => p.StandardError.ReadToEnd());
                    if (!p.WaitForExit(20000))
                    {
                        try { p.Kill(); } catch (Exception) { }
                    }
                    string text = outTask.Wait(5000) ? outTask.Result : "";
                    errTask.Wait(1000);
                    foreach (string l in text.Replace("\r", "").Split('\n')) lines.Add(l);
                }
            }
            catch (Exception) { }
            return lines;
        }

        // ------------------------------------------------------------------
        //  Blocks that need more than one line of code
        // ------------------------------------------------------------------

        private static string AnyDeskId()
        {
            string[] confs =
            {
                Env("ProgramData").Length > 0 ? Path.Combine(Env("ProgramData"), "AnyDesk\\system.conf") : null,
                Env("APPDATA").Length > 0 ? Path.Combine(Env("APPDATA"), "AnyDesk\\system.conf") : null
            };
            string value = null;
            foreach (string conf in confs)
            {
                if (!string.IsNullOrEmpty(value) || conf == null || !File.Exists(conf)) continue;
                string[] lines;
                try { lines = File.ReadAllLines(conf); }
                catch (Exception) { continue; }
                string line = lines.FirstOrDefault(l => Regex.IsMatch(l, "ad.anynet.id", RegexOptions.IgnoreCase));
                if (!string.IsNullOrEmpty(line)) value = Regex.Replace(line, ".*=", "").Trim();
            }
            return value;
        }

        private List<string> NetworkLines()
        {
            var lines = new List<string>();
            List<ManagementObject> profiles = TryQ(StandardCimv2, "MSFT_NetConnectionProfile", null);
            if (profiles != null)
            {
                foreach (ManagementObject p in profiles)
                {
                    if ((int)Num(P(p, "IPv4Connectivity")) == 0) continue; // Disconnected
                    string alias = S(P(p, "InterfaceAlias")) ?? "";
                    string index = S(P(p, "InterfaceIndex")) ?? "0";
                    string name = S(P(p, "Name")) ?? "";
                    var addrs = new List<string>();
                    foreach (ManagementObject a in Q(StandardCimv2, "MSFT_NetIPAddress", "AddressFamily=2 AND InterfaceIndex=" + index))
                    {
                        string ip = S(P(a, "IPAddress"));
                        if (!string.IsNullOrEmpty(ip)) addrs.Add(ip);
                    }
                    bool dhcp = false;
                    foreach (ManagementObject i in Q(StandardCimv2, "MSFT_NetIPInterface", "AddressFamily=2 AND InterfaceIndex=" + index))
                        if (Num(P(i, "Dhcp")) == 1) dhcp = true;
                    string how = dhcp ? "Dynamic" : "Static";
                    string ipText = string.Join(", ", addrs);
                    if (Regex.IsMatch(alias, "Wi[- ]?Fi", RegexOptions.IgnoreCase))
                        lines.Add("Wi-Fi | SSID: " + name + " | IP: " + ipText + " | " + how);
                    else
                        lines.Add("LAN | Network Name: " + name + " | IP: " + ipText + " | " + how);
                }
                return lines;
            }

            // Windows 7: adapter settings and netsh give the same facts.
            string ssid = "";
            foreach (string line in RunLines("netsh.exe", "wlan show interfaces"))
            {
                Match m = Regex.Match(line, @"^\s*SSID\s*:\s*(.+)$", RegexOptions.IgnoreCase);
                if (m.Success) ssid = m.Groups[1].Value.Trim();
            }
            foreach (ManagementObject nic in Q(Cimv2, "Win32_NetworkAdapterConfiguration", "IPEnabled = TRUE"))
            {
                var ips = new List<string>();
                foreach (object ip in Arr(P(nic, "IPAddress")))
                {
                    string s = S(ip);
                    if (s != null && Regex.IsMatch(s, @"^\d+\.\d+\.\d+\.\d+$")) ips.Add(s);
                }
                string how = IsTrue(P(nic, "DHCPEnabled")) ? "Dynamic" : "Static";
                string desc = S(P(nic, "Description")) ?? "";
                string ipText = string.Join(", ", ips);
                if (Regex.IsMatch(desc, "Wireless|Wi-?Fi|WLAN|802\\.11", RegexOptions.IgnoreCase))
                    lines.Add("Wi-Fi | SSID: " + ssid + " | IP: " + ipText + " | " + how);
                else
                    lines.Add("LAN | Network Name: | IP: " + ipText + " | " + how);
            }
            return lines;
        }

        private List<string> CredentialLines()
        {
            var lines = new List<string>();
            string[] servers = { "192.168.1.250", "server" };
            bool found = false;
            string target = null;
            foreach (string line in RunLines("cmdkey.exe", "/list"))
            {
                Match t = Regex.Match(line, @"^\s*Target:\s*(.+)$", RegexOptions.IgnoreCase);
                if (t.Success)
                {
                    target = Regex.Replace(t.Groups[1].Value, ".*target=", "", RegexOptions.IgnoreCase);
                    continue;
                }
                Match u = Regex.Match(line, @"^\s*User:\s*(.+)$", RegexOptions.IgnoreCase);
                if (u.Success && target != null && servers.Any(s => string.Equals(s, target, StringComparison.OrdinalIgnoreCase)))
                {
                    found = true;
                    lines.Add(target + " | " + u.Groups[1].Value);
                }
            }
            if (!found) lines.Add("None");
            return lines;
        }

        private List<string> ServerFolderLines()
        {
            var lines = new List<string>();
            foreach (string name in SubKeyNames(Hkcu(), "Network"))
            {
                Dictionary<string, string> k = ReadKey(Hkcu(), "Network\\" + name);
                lines.Add(name + ": | " + (G(k, "RemotePath")));
            }
            return lines;
        }

        private static List<UninstallEntry> ReadUninstall()
        {
            var list = new List<UninstallEntry>();
            AddUninstall(list, Lm(), "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall");
            AddUninstall(list, Lm(), "SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall");
            AddUninstall(list, Hkcu(), "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall");
            return list;
        }

        private static void AddUninstall(List<UninstallEntry> list, RegistryKey root, string path)
        {
            foreach (string id in SubKeyNames(root, path))
            {
                Dictionary<string, string> v = ReadKey(root, path + "\\" + id);
                string name = Reg(v, "DisplayName");
                if (string.IsNullOrEmpty(name)) continue;
                list.Add(new UninstallEntry
                {
                    DisplayName = name,
                    DisplayVersion = Reg(v, "DisplayVersion"),
                    Publisher = Reg(v, "Publisher"),
                    InstallDate = Reg(v, "InstallDate"),
                    SystemComponent = Reg(v, "SystemComponent"),
                    ParentKeyName = Reg(v, "ParentKeyName"),
                    ReleaseType = Reg(v, "ReleaseType")
                });
            }
        }

        private List<string> MailFiles()
        {
            var hits = new List<string>();
            string local = Env("LOCALAPPDATA");
            string profile = Env("USERPROFILE");
            if (local.Length > 0) Walk(Path.Combine(local, "Microsoft\\Outlook"), hits);
            if (profile.Length > 0) Walk(Path.Combine(profile, "Documents\\Outlook Files"), hits);
            if (hits.Count == 0) hits.Add("None");
            return hits;
        }

        // Same walk as Get-ChildItem -Recurse -Include *.ost, *.pst: entries in the order the file system gives them.
        private static void Walk(string dir, List<string> hits)
        {
            string[] entries;
            try { entries = Directory.GetFileSystemEntries(dir); }
            catch (Exception) { return; }
            foreach (string e in entries)
            {
                if (Directory.Exists(e)) Walk(e, hits);
                else if (e.EndsWith(".ost", StringComparison.OrdinalIgnoreCase) || e.EndsWith(".pst", StringComparison.OrdinalIgnoreCase))
                    hits.Add(Path.GetFileName(e) + " | " + e);
            }
        }

        private static string Channel(Licence l)
        {
            if (!string.IsNullOrEmpty(l.ProductChannel)) return l.ProductChannel;
            Match m = Regex.Match(l.Description ?? "", @",\s*(\S+)\s+channel", RegexOptions.IgnoreCase);
            if (m.Success) return m.Groups[1].Value;
            return "Unknown";
        }

        private static string StatusWord(int status)
        {
            string w;
            return StatusWords.TryGetValue(status, out w) ? w : "";
        }

        private static Licence LicenceFrom(ManagementObject o)
        {
            return new Licence
            {
                Name = S(P(o, "Name")),
                Description = S(P(o, "Description")),
                AppId = S(P(o, "ApplicationID")),
                PartialKey = S(P(o, "PartialProductKey")),
                ProductChannel = S(P(o, "ProductKeyChannel")),
                Status = (int)Num(P(o, "LicenseStatus"))
            };
        }

        private static List<Licence> QueryLicences()
        {
            var list = new List<Licence>();
            foreach (ManagementObject o in Q(Cimv2, "SoftwareLicensingProduct", "PartialProductKey IS NOT NULL"))
                list.Add(LicenceFrom(o));
            return list;
        }

        private static List<Disk> LoadDisks()
        {
            var list = new List<Disk>();
            foreach (ManagementObject o in Q(Storage, "MSFT_PhysicalDisk", null))
            {
                list.Add(new Disk
                {
                    Name = S(P(o, "FriendlyName")) ?? "",
                    Media = MediaName((int)Num(P(o, "MediaType"))),
                    Size = Num(P(o, "Size")),
                    Serial = S(P(o, "SerialNumber")),
                    Bus = BusName((int)Num(P(o, "BusType"))),
                    Firmware = S(P(o, "FirmwareVersion"))
                });
            }
            if (list.Count == 0)
            {
                foreach (ManagementObject o in Q(Cimv2, "Win32_DiskDrive", null))
                {
                    list.Add(new Disk
                    {
                        Name = S(P(o, "Model")) ?? "",
                        Media = "Unspecified",
                        Size = Num(P(o, "Size")),
                        Serial = S(P(o, "SerialNumber")),
                        Bus = S(P(o, "InterfaceType")) ?? "",
                        Firmware = S(P(o, "FirmwareRevision"))
                    });
                }
            }
            return list;
        }

        private static string MediaName(int code)
        {
            switch (code)
            {
                case 3: return "HDD";
                case 4: return "SSD";
                case 5: return "SCM";
                default: return "Unspecified";
            }
        }

        private static string BusName(int code)
        {
            switch (code)
            {
                case 17: return "NVMe";
                case 11: return "SATA";
                case 7: return "USB";
                case 8: return "RAID";
                case 10: return "SAS";
                case 3: return "ATA";
                case 1: return "SCSI";
                default: return "Unknown";
            }
        }
    }
}
