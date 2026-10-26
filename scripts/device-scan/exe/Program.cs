using System;
using System.Windows.Forms;

namespace PmwDeviceScan
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            // Hidden design check: --ui-demo <new|known|scanning|done|error|settings> shows one screen with
            // sample values. Nothing is scanned, saved or written.
            string demo = ArgValue(args, "--ui-demo");
            if (demo != null)
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                if (string.Equals(demo, "settings", StringComparison.OrdinalIgnoreCase)) Application.Run(new SettingsForm(true));
                else Application.Run(new MainForm(demo.ToLowerInvariant()));
                return 0;
            }

            // Hidden test switch: --report-to <folder> [--place "<PLACE>"] scans without a window.
            string folder = ArgValue(args, "--report-to");
            if (folder != null)
            {
                return HeadlessRun(folder, ArgValue(args, "--place") ?? "");
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
            return 0;
        }

        private static string ArgValue(string[] args, string name)
        {
            for (int i = 0; i < args.Length - 1; i++)
                if (string.Equals(args[i], name, StringComparison.OrdinalIgnoreCase)) return args[i + 1];
            return null;
        }

        private static int HeadlessRun(string folder, string place)
        {
            try
            {
                var lines = ReportBuilder.Build(null);
                ReportFiles.SaveTo(folder, lines, ReportFiles.CleanPlace(place));
                return 0;
            }
            catch (Exception)
            {
                return 1;
            }
        }
    }
}
