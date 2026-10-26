using System;
using System.Drawing;
using System.Windows.Forms;

namespace PmwDeviceScan
{
    // The Settings window: a sheet of its own, opened from the gear. Every change is saved at once
    // (PMW-Device-Scan.ini beside the program), so there is no OK button. Demo mode shows the window
    // with the current values and never writes anything.
    internal sealed class SettingsForm : Form
    {
        private readonly bool _demo;
        private Settings _current;
        private bool _loading;

        private NavItem _navGeneral, _navAbout;
        private Panel _general, _about;
        private Card _folderCard, _infoCard, _startCard, _closeCard;
        private Label _folderPath, _folderKept, _infoText, _startDesc, _startState, _closeDesc, _closeState;
        private RoundButton _change, _reset;
        private ToggleSwitch _startSwitch, _closeSwitch;
        private GlyphView _statusIcon;
        private Label _statusText;

        public SettingsForm(bool demo)
        {
            _demo = demo;
            Text = "PMW Device Scan settings";
            ClientSize = new Size(780, 568);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Ui.Tile;
            Font = Ui.Body(14);

            _current = Copy(ReportFiles.Current);
            BuildNav();
            BuildGeneral();
            BuildAbout();
            BuildStatus();
            ShowNav(true);
            ShowCurrent();
            SetStatus("Saved on this drive in PMW-Device-Scan.ini · the .bat and .ps1 follow it too", false);
        }

        private static Settings Copy(Settings s)
        {
            return new Settings { ReportsFolder = s.ReportsFolder ?? "", AutoStart = s.AutoStart, AutoClose = s.AutoClose };
        }

        private void BuildNav()
        {
            _navGeneral = new NavItem { Text = "General", Location = new Point(8, 8), Size = new Size(184, 36), Selected = true };
            _navAbout = new NavItem { Text = "About", Location = new Point(8, 46), Size = new Size(184, 36) };
            _navGeneral.Click += delegate { ShowNav(true); };
            _navAbout.Click += delegate { ShowNav(false); };
            Controls.Add(_navGeneral);
            Controls.Add(_navAbout);
        }

        private void ShowNav(bool general)
        {
            _general.Visible = general;
            _about.Visible = !general;
            _navGeneral.Selected = general;
            _navAbout.Selected = !general;
        }

        private void BuildGeneral()
        {
            _general = new Panel { Location = new Point(200, 0), Size = new Size(580, 520), BackColor = Ui.Tile };
            Controls.Add(_general);
            Ui.AddLabel(_general, "General", 8, 14, 548, 34, Ui.Semi(24), Ui.Ink);

            // Where reports are saved.
            _folderCard = new Card { Location = new Point(8, 64), Size = new Size(548, 68), Fill = Color.White, Edge = Ui.CardLine };
            _general.Controls.Add(_folderCard);
            Ui.AddLabel(_folderCard, "Save reports to", 16, 14, 340, 20, Ui.Body(14), Ui.Ink);
            _folderPath = Ui.AddLabel(_folderCard, "", 16, 36, 332, 18, Ui.Body(12), Ui.Soft);
            _folderPath.AutoEllipsis = true;
            _folderKept = Ui.AddLabel(_folderCard, "", 16, 56, 332, 18, Ui.Body(12), Ui.Soft);
            _change = new RoundButton { Text = "Change", Size = new Size(80, 32), Font = Ui.Body(14) };
            _reset = new RoundButton { Text = "Reset", Size = new Size(72, 32), Font = Ui.Body(14) };
            _change.Click += delegate { PickFolder(); };
            _reset.Click += delegate
            {
                Settings c = Copy(_current);
                c.ReportsFolder = "";
                Apply(c);
            };
            _folderCard.Controls.Add(_change);
            _folderCard.Controls.Add(_reset);

            // The note about network folders.
            _infoCard = new Card { Location = new Point(8, 140), Size = new Size(548, 60), Fill = Color.White, Edge = Ui.CardLine };
            _general.Controls.Add(_infoCard);
            _infoText = Ui.AddLabel(_infoCard, "A network folder such as \\\\server\\PMW\\IT works too. If it cannot be reached, "
                + "the report is saved beside the program and Completed says so. Reset goes back to the Reports folder beside the program.",
                16, 12, 516, 36, Ui.Body(12), Ui.Soft);

            // Start scanning by itself.
            _startCard = new Card { Location = new Point(8, 210), Size = new Size(548, 64), Fill = Color.White, Edge = Ui.CardLine };
            _general.Controls.Add(_startCard);
            Ui.AddLabel(_startCard, "Start scanning by itself", 16, 12, 400, 24, Ui.Body(14), Ui.Ink);
            _startDesc = Ui.AddLabel(_startCard, "A PC scanned before starts the moment the program opens. "
                + "A new PC still asks where it is.", 16, 38, 400, 18, Ui.Body(12), Ui.Soft);
            _startState = Ui.AddLabel(_startCard, "", 444, 22, 36, 20, Ui.Body(14), Ui.Ink);
            _startState.TextAlign = ContentAlignment.MiddleRight;
            _startSwitch = new ToggleSwitch { Location = new Point(492, 22), AccessibleName = "Start scanning by itself" };
            _startSwitch.CheckedChanged += OnStartChanged;
            _startCard.Controls.Add(_startSwitch);

            // Close by itself when done.
            _closeCard = new Card { Location = new Point(8, 282), Size = new Size(548, 64), Fill = Color.White, Edge = Ui.CardLine };
            _general.Controls.Add(_closeCard);
            Ui.AddLabel(_closeCard, "Close by itself when done", 16, 12, 400, 24, Ui.Body(14), Ui.Ink);
            _closeDesc = Ui.AddLabel(_closeCard, "Ten seconds after Completed.", 16, 38, 400, 18, Ui.Body(12), Ui.Soft);
            _closeState = Ui.AddLabel(_closeCard, "", 444, 22, 36, 20, Ui.Body(14), Ui.Ink);
            _closeState.TextAlign = ContentAlignment.MiddleRight;
            _closeSwitch = new ToggleSwitch { Location = new Point(492, 22), AccessibleName = "Close by itself when done" };
            _closeSwitch.CheckedChanged += OnCloseChanged;
            _closeCard.Controls.Add(_closeSwitch);
        }

        private void BuildAbout()
        {
            _about = new Panel { Location = new Point(200, 0), Size = new Size(580, 520), BackColor = Ui.Tile, Visible = false };
            Controls.Add(_about);
            Ui.AddLabel(_about, "About", 8, 14, 548, 34, Ui.Semi(24), Ui.Ink);

            var card = new Card { Location = new Point(8, 64), Size = new Size(548, 150), Fill = Color.White, Edge = Ui.CardLine };
            _about.Controls.Add(card);
            Ui.AddLabel(card, "PMW Device Scan", 16, 14, 516, 24, Ui.Semi(16), Ui.Ink);
            Ui.AddLabel(card, "Version " + AppInfo.Version, 16, 40, 516, 18, Ui.Body(12), Ui.Soft);
            Ui.AddLabel(card, "Scans this PC into a report on the drive: hardware, serial and part numbers, "
                + "Windows and Office licences, and installed software. It writes the same report as PMW-Device-Scan.bat.",
                16, 68, 516, 60, Ui.Body(14), Ui.Ink);
        }

        // The line at the bottom of the window: a green tick when saved, a red "!" when not.
        private void BuildStatus()
        {
            _statusIcon = new GlyphView { Glyph = "check", Tint = Ui.Green, Px = 16, Location = new Point(208, 530), Size = new Size(16, 16), BackColor = Ui.Tile };
            _statusText = new Label { Location = new Point(230, 528), Size = new Size(540, 20), Font = Ui.Body(12), ForeColor = Ui.Soft, BackColor = Ui.Tile };
            Controls.Add(_statusIcon);
            Controls.Add(_statusText);
        }

        private void SetStatus(string text, bool error)
        {
            _statusText.Text = text;
            _statusText.ForeColor = error ? Ui.Red : Ui.Soft;
            _statusIcon.Glyph = error ? "!" : "check";
            _statusIcon.Tint = error ? Ui.Red : Ui.Green;
            _statusIcon.Invalidate();
        }

        // Puts the saved values on the controls. Runs with the switches' events muted.
        private void ShowCurrent()
        {
            _loading = true;
            _folderPath.Text = _current.ResolvedFolder();
            _folderKept.Text = "Kept as " + _current.ReportsFolder + ", so it works whatever letter the drive gets.";
            _startSwitch.Checked = _current.AutoStart;
            _closeSwitch.Checked = _current.AutoClose;
            _startState.Text = _current.AutoStart ? "On" : "Off";
            _closeState.Text = _current.AutoClose ? "On" : "Off";
            _loading = false;
            LayoutGeneral();
        }

        // Stacks the General cards from their text, so a long path or a note never clips.
        private void LayoutGeneral()
        {
            string folder = _current.ReportsFolder ?? "";
            bool kept = folder.StartsWith("\\", StringComparison.Ordinal) && !folder.StartsWith("\\\\", StringComparison.Ordinal);
            _folderKept.Visible = kept;
            int keptH = kept ? Ui.Lines(_folderKept.Text, _folderKept.Font, 332) : 0;
            _folderKept.Size = new Size(332, Math.Max(18, keptH));
            int folderH = kept ? 56 + Math.Max(18, keptH) + 14 : 68;
            _folderCard.Size = new Size(548, folderH);
            _reset.Location = new Point(548 - 16 - 72, (folderH - 32) / 2);
            _change.Location = new Point(_reset.Left - 8 - 80, (folderH - 32) / 2);

            int y = _folderCard.Bottom + 8;
            int infoTextH = Ui.Lines(_infoText.Text, _infoText.Font, 516);
            _infoText.Size = new Size(516, infoTextH);
            _infoCard.Location = new Point(8, y);
            _infoCard.Size = new Size(548, infoTextH + 24);

            y = _infoCard.Bottom + 10;
            int descH = Ui.Lines(_startDesc.Text, _startDesc.Font, 400);
            int startH = Math.Max(64, 38 + descH + 14);
            _startCard.Location = new Point(8, y);
            _startCard.Size = new Size(548, startH);
            _startDesc.Size = new Size(400, Math.Max(18, descH));
            _startSwitch.Top = (startH - 20) / 2;
            _startState.Top = (startH - 20) / 2;

            y = _startCard.Bottom + 8;
            int closeDescH = Ui.Lines(_closeDesc.Text, _closeDesc.Font, 400);
            int closeH = Math.Max(64, 38 + closeDescH + 14);
            _closeCard.Location = new Point(8, y);
            _closeCard.Size = new Size(548, closeH);
            _closeDesc.Size = new Size(400, Math.Max(18, closeDescH));
            _closeSwitch.Top = (closeH - 20) / 2;
            _closeState.Top = (closeH - 20) / 2;
        }

        private void OnStartChanged(object sender, EventArgs e)
        {
            if (_loading) return;
            Settings c = Copy(_current);
            c.AutoStart = _startSwitch.Checked;
            Apply(c);
        }

        private void OnCloseChanged(object sender, EventArgs e)
        {
            if (_loading) return;
            Settings c = Copy(_current);
            c.AutoClose = _closeSwitch.Checked;
            Apply(c);
        }

        private void PickFolder()
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Where should reports be saved?";
                try { dialog.SelectedPath = _current.ResolvedFolder(); } catch (Exception) { }
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                Settings c = Copy(_current);
                c.ReportsFolder = Settings.ToStored(dialog.SelectedPath);
                Apply(c);
            }
        }

        // Saves the candidate at once. A drive that cannot take the file keeps the old values on screen.
        private void Apply(Settings candidate)
        {
            if (!_demo)
            {
                try
                {
                    candidate.Save();
                }
                catch (Exception ex)
                {
                    ShowCurrent();
                    SetStatus("The settings could not be saved on this drive: " + ex.Message, true);
                    return;
                }
                ReportFiles.Current = candidate;
            }
            _current = candidate;
            ShowCurrent();
            SetStatus("Saved on this drive in PMW-Device-Scan.ini · the .bat and .ps1 follow it too", false);
        }
    }
}
